package controller;

import model.*;
import repository.ExcelRepository;
import service.RicaviExcelService;
import view.*;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.apache.poi.ss.usermodel.*;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.data.category.DefaultCategoryDataset;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.Desktop;
import java.io.File;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.*;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;

import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;

public class MainController {

    private static final Logger log = LogManager.getLogger(MainController.class);

    private final AppModel model;
    private final MainFrame view;
    private final ExcelRepository excelRepo;
    private PremioCompController premioController;

    private RicaviExcelService ricaviService;
    private List<ArticleRow> cachedArticles = new ArrayList<>();

    private static final DecimalFormat DF_INT = new DecimalFormat("#,##0");
    private static final DecimalFormat DF_3   = new DecimalFormat("#,##0.000");
    private static final DecimalFormat DF_2   = new DecimalFormat("#,##0.00");

    /**
     * Riga HTML standard per la sezione "Dettagli".
     * Mantiene un look pulito senza cambiare la logica dei calcoli.
     */
    private static String rowHtml(String k, String v) {
        String key = (k == null) ? "" : k;
        String val = (v == null) ? "" : v;
        return "<tr>" +
                "<td style='padding:4px 8px;border-top:1px solid #eee;color:#333;white-space:nowrap;'><b>" + key + "</b></td>" +
                "<td style='padding:4px 8px;border-top:1px solid #eee;color:#111;text-align:right;'>" + val + "</td>" +
                "</tr>";
    }

    // ==========================================================
    // ✅ FIX EXCEL: helper per aggiornare celle senza rompere formule
    // ==========================================================
    private void writeNumericIfNotFormula(Sheet sheet, int rowIdx, int colIdx, double value) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) row = sheet.createRow(rowIdx);

        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

        // NON sovrascrivere formule
        if (cell.getCellType() == CellType.FORMULA) return;

        cell.setCellValue(value);
    }

    /**
     * Se la cella è una formula, valuta e aggiorna il cached result
     * (serve per vedere valori aggiornati aprendo la copia in Excel).
     */
    private void evaluateFormulaCellIfPresent(Sheet sheet, FormulaEvaluator eval, int rowIdx, Integer colIdx) {
        if (colIdx == null) return;

        Row row = sheet.getRow(rowIdx);
        if (row == null) return;

        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return;

        if (cell.getCellType() == CellType.FORMULA) {
            eval.evaluateFormulaCell(cell);
        }
    }

    /**
     * Forza Excel a ricalcolare all'apertura del file.
     */
    private void forceExcelRecalcOnOpen(Workbook wb) {
        wb.setForceFormulaRecalculation(true);
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            wb.getSheetAt(i).setForceFormulaRecalculation(true);
        }
    }

    // ===========================
    // CE Budget 2022: mapping righe/colonna
    // (valori in colonna J: J5..J13)
    // ===========================
    private static final int CE_COL_J = 9; // J = index 9 (0-based)

    private static final int CE_ROW_RICAVI_PF      = 4;  // J5
    private static final int CE_ROW_RICAVI_MP      = 5;  // J6
    private static final int CE_ROW_RICAVI_CLAV    = 6;  // J7
    private static final int CE_ROW_ALTRI_RICAVI   = 7;  // J8
    private static final int CE_ROW_VAR_PF         = 8;  // J9
    private static final int CE_ROW_ACQUISTO_MP    = 10; // J11
    private static final int CE_ROW_VAR_SCORTE     = 11; // J12

    private static final String K_RICAVI_PF    = "RICAVI_PF";
    private static final String K_RICAVI_MP    = "RICAVI_MP";
    private static final String K_RICAVI_CLAV  = "RICAVI_CLAV";
    private static final String K_ALTRI_RICAVI = "ALTRI_RICAVI";
    private static final String K_VAR_PF       = "VAR_PF";
    private static final String K_ACQUISTO_MP  = "ACQUISTO_MP";
    private static final String K_VAR_SCORTE   = "VAR_SCORTE";

    public MainController(AppModel model, MainFrame view, ExcelRepository excelRepo) {
        this.model = model;
        this.view = view;
        this.excelRepo = excelRepo;

        log.info("MainController inizializzato.");
        initListeners();
    }

    private void initListeners() {
        view.getBtnResetExcel().addActionListener(e -> onResetExcel());
        view.getBtnLoadExcel().addActionListener(e -> onLoadExcel());
        view.getBtnExit().addActionListener(e -> onExit());
        view.getBtnOpenWorkingCopy().addActionListener(e -> onOpenWorkingCopy());
        view.getControlsPanel().getBtnSimulate().addActionListener(e -> onSimulate());

        // ✅ nuovo listener: CE Budget 2022 (base fisso)
        view.getBtnShowCeBudget().addActionListener(e -> onShowCeBudgetBase());

        log.debug("Listener UI registrati.");
    }

    private void onOpenPremioComp() {
        if (model.getWorkingExcelCopy() == null) {
            JOptionPane.showMessageDialog(view, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
            return;
        }
        if (premioController == null) {
            premioController = new PremioCompController(model, view, excelRepo);
        }
        premioController.open();
    }

    private void onResetExcel() {
        if (model.getWorkingExcelCopy() == null) {
            JOptionPane.showMessageDialog(view, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
            return;
        }

        int ok = JOptionPane.showConfirmDialog(
                view,
                "Vuoi ripristinare la copia di lavoro allo stato iniziale?\n" +
                        "Perderai tutte le simulazioni fatte finora nella sessione.",
                "Reset Excel",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.WARNING_MESSAGE
        );
        if (ok != JOptionPane.YES_OPTION) return;

        try {
            // ✅ reset working copy dal base snapshot
            File wc = excelRepo.resetWorkingCopyToBase();
            model.setWorkingExcelCopy(wc);

            // ✅ ricreo service e ricarico articoli (così riparti da base)
            ricaviService = new RicaviExcelService(wc);
            cachedArticles = ricaviService.loadArticles();
            view.getControlsPanel().setArticles(cachedArticles);

            // ✅ pulisco output (dettagli + grafici)
            view.getControlsPanel().setDetails("");
            view.getChartsPanel().setPosChart(null);
            view.getChartsPanel().setCompChart(null);
            view.getChartsPanel().setCeBaseChart(null);
            view.getChartsPanel().setCeVarChart(null);
            view.getChartsPanel().setCeDeltaChart(null);

            JOptionPane.showMessageDialog(view, "Reset completato. Riparti dai valori base.", "OK", JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception ex) {
            log.error("Errore reset working copy", ex);
            JOptionPane.showMessageDialog(
                    view,
                    "Errore reset: " + ex.getMessage() + "\n\n" +
                            "Se hai aperto la copia in Excel, chiudila e riprova.",
                    "Errore",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }

    private void onLoadExcel() {
        log.info("Click: Carica Excel");

        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Seleziona il file Excel (.xlsx)");
        chooser.setFileFilter(new FileNameExtensionFilter("Excel (*.xlsx)", "xlsx"));

        int result = chooser.showOpenDialog(view);
        if (result != JFileChooser.APPROVE_OPTION) return;

        File original = chooser.getSelectedFile();
        log.info("File selezionato: {}", original.getAbsolutePath());

        try {
            File workingCopy = excelRepo.createWorkingCopy(original);

            model.setOriginalExcel(original);
            model.setWorkingExcelCopy(workingCopy);
            view.setExcelLoaded(original.getName());

            ricaviService = new RicaviExcelService(workingCopy);
            cachedArticles = ricaviService.loadArticles();
            view.getControlsPanel().setArticles(cachedArticles);

            JOptionPane.showMessageDialog(
                    view,
                    "Excel caricato.\nArticoli modificabili trovati: " + cachedArticles.size(),
                    "OK",
                    JOptionPane.INFORMATION_MESSAGE
            );

        } catch (Exception ex) {
            log.error("Errore nel caricamento Excel", ex);
            JOptionPane.showMessageDialog(view, "Errore: " + ex.getMessage(), "Errore", JOptionPane.ERROR_MESSAGE);
            view.setExcelNotLoaded();
        }
    }

    private void onOpenWorkingCopy() {
        try {
            File f = model.getWorkingExcelCopy();
            if (f == null) return;

            if (!f.exists() || f.length() == 0) {
                JOptionPane.showMessageDialog(view, "Copia di lavoro non valida.\nRicarica l'Excel.", "Errore",
                        JOptionPane.ERROR_MESSAGE);
                return;
            }

            Desktop.getDesktop().open(f);

        } catch (Exception ex) {
            log.error("Impossibile aprire working copy", ex);
            JOptionPane.showMessageDialog(view, "Impossibile aprire la copia: " + ex.getMessage(), "Errore",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private Map<String, Integer> buildRowIndexMap(Sheet ricaviSheet, DataFormatter fmt, int headerRowIdx, int colCat, int colArt) {
        Map<String, Integer> map = new HashMap<>();
        for (int r = headerRowIdx + 1; r <= ricaviSheet.getLastRowNum(); r++) {
            Row rr = ricaviSheet.getRow(r);
            if (rr == null) continue;

            String catTxt = fmt.formatCellValue(rr.getCell(colCat)).trim().toUpperCase();
            String artTxt = fmt.formatCellValue(rr.getCell(colArt)).trim().toUpperCase();
            if (catTxt.isEmpty() || artTxt.isEmpty()) continue;

            map.put(catTxt + "||" + artTxt, r);
        }
        return map;
    }

    private void onSimulate() {

    if (model.getWorkingExcelCopy() == null || ricaviService == null) {
        JOptionPane.showMessageDialog(view, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
        return;
    }

    List<SimulationControlsPanel.SimRequest> requests = view.getControlsPanel().getSimulationRequests();
    if (requests == null || requests.isEmpty()) {
        JOptionPane.showMessageDialog(view, "Seleziona almeno un articolo (colonna 'Sel').", "Attenzione", JOptionPane.WARNING_MESSAGE);
        return;
    }

    try (Workbook wb = WorkbookFactory.create(model.getWorkingExcelCopy())) {

        Sheet ricaviSheet = wb.getSheet("Ricavi");
        if (ricaviSheet == null) throw new IllegalStateException("Foglio 'Ricavi' non trovato.");

        FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
        DataFormatter fmt = new DataFormatter();

        // =========================================================
        // 1) Trovo header tabella destra e colonne chiave (UNA VOLTA)
        // =========================================================
        int headerRowIdx = -1;
        for (int r = 0; r <= Math.min(ricaviSheet.getLastRowNum(), 200); r++) {
            Row row = ricaviSheet.getRow(r);
            if (row == null) continue;

            boolean hasCat = false, hasArt = false, hasQty = false, hasPos = false;
            for (int c = 0; c < Math.min(row.getLastCellNum(), 200); c++) {
                String v = fmt.formatCellValue(row.getCell(c)).trim();
                if (v.equalsIgnoreCase("Cat")) hasCat = true;
                if (v.equalsIgnoreCase("Articolo")) hasArt = true;
                if (v.toLowerCase().contains("quantità")) hasQty = true;
                if (v.equalsIgnoreCase("POS")) hasPos = true;
            }
            if (hasCat && hasArt && hasQty && hasPos) {
                headerRowIdx = r;
                break;
            }
        }
        if (headerRowIdx < 0) {
            throw new IllegalStateException("Header tabella destra non trovato (Cat/Articolo/Quantità/POS).");
        }

        Row header = ricaviSheet.getRow(headerRowIdx);

        java.util.function.BiFunction<String, int[], Integer> findExactInWindow = (exact, win) -> {
            int start = win[0], end = win[1];
            for (int c = start; c <= end; c++) {
                String v = fmt.formatCellValue(header.getCell(c)).trim();
                if (v.equalsIgnoreCase(exact)) return c;
            }
            return null;
        };

        java.util.function.BiFunction<String, int[], Integer> findLastExactInWindow = (exact, win) -> {
            int start = win[0], end = win[1];
            Integer found = null;
            for (int c = start; c <= end; c++) {
                String v = fmt.formatCellValue(header.getCell(c)).trim();
                if (v.equalsIgnoreCase(exact)) found = c;
            }
            return found;
        };

        java.util.function.BiFunction<String, int[], Integer> findContainsInWindow = (needleLower, win) -> {
            String needle = needleLower.toLowerCase();
            int start = win[0], end = win[1];
            for (int c = start; c <= end; c++) {
                String v = fmt.formatCellValue(header.getCell(c)).trim().toLowerCase();
                if (v.contains(needle)) return c;
            }
            return null;
        };

        List<Integer> catCandidates = new ArrayList<>();
        for (int c = 0; c < header.getLastCellNum(); c++) {
            String v = fmt.formatCellValue(header.getCell(c)).trim();
            if (v.equalsIgnoreCase("Cat")) catCandidates.add(c);
        }
        if (catCandidates.isEmpty()) throw new IllegalStateException("Header: 'Cat' non trovato.");

        Integer colCat = null, colArt = null, colQty = null, colPos = null;
        Integer colPeur = null, colCMPeur = null;
        Integer colFatt = null;
        Integer colCogs = null;

        boolean foundTable = false;

        for (int catColCandidate : catCandidates) {
            int start = catColCandidate;
            int end = Math.min(header.getLastCellNum() - 1, catColCandidate + 50);
            int[] win = new int[]{start, end};

            Integer a = findExactInWindow.apply("Articolo", win);
            Integer q = findContainsInWindow.apply("quantità", win);
            Integer p = findLastExactInWindow.apply("POS", win);

            Integer pe = findContainsInWindow.apply("p medio (€/kg)", win);
            Integer cmpe = findContainsInWindow.apply("cmp medio (€/kg)", win);

            Integer fatt = findContainsInWindow.apply("fatturato", win);
            Integer cogs = findContainsInWindow.apply("cogs", win);
            if (cogs == null) cogs = findContainsInWindow.apply("costo del venduto", win);

            if (a != null && q != null && p != null && pe != null && cmpe != null) {
                colCat = catColCandidate;
                colArt = a;
                colQty = q;
                colPos = p;
                colPeur = pe;
                colCMPeur = cmpe;
                colFatt = fatt;
                colCogs = cogs;
                foundTable = true;
                break;
            }
        }

        if (!foundTable) {
            throw new IllegalStateException(
                    "Impossibile identificare le colonne necessarie nella tabella destra.\n" +
                            "Servono: Cat, Articolo, Quantità, P medio (€/kg), CMP medio (€/kg), POS."
            );
        }

        // =========================================================
        // 2) Mappa (Cat,Articolo) -> rowIndex (UNA VOLTA)
        // =========================================================
        Map<String, Integer> rowMap = buildRowIndexMap(ricaviSheet, fmt, headerRowIdx, colCat, colArt);

        // =========================================================
        // 3) Pre-lettura valori base per ogni articolo (per grafici coerenti)
        // =========================================================
        eval.evaluateAll();

        class Base {
            int rowIdx;
            String cat, art;
            double q0, p0, cmp0, pos0Excel, fatt0, cogs0, pos0Calc;
        }
        Map<String, Base> baseByKey = new LinkedHashMap<>();

        for (SimulationControlsPanel.SimRequest req : requests) {
            String cat = (req.article.getCat() == null) ? "" : req.article.getCat().trim().toUpperCase();
            String art = (req.article.getArticolo() == null) ? "" : req.article.getArticolo().trim().toUpperCase();
            String key = cat + "||" + art;

            Integer rowIdx = rowMap.get(key);
            if (rowIdx == null) {
                throw new IllegalStateException("Non trovo la riga per Cat='" + cat + "' e Articolo='" + art + "'.");
            }

            double q0 = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colQty);
            double p0 = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPeur);
            double cmp0 = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colCMPeur);
            double pos0Excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);

            if (!(q0 > 0)) throw new IllegalStateException("Q0 non valida letta da Excel: " + q0);
            if (!(p0 > 0)) throw new IllegalStateException("P0 (€/kg) non valido letto da Excel: " + p0);

            // CMP può essere 0 (trattino nel file). Vietato solo <0 o NaN.
            if (Double.isNaN(cmp0) || cmp0 < 0) {
                throw new IllegalStateException("CMP0 (€/kg) non valido letto da Excel: " + cmp0);
            }


            Base b = new Base();
            b.rowIdx = rowIdx;
            b.cat = cat;
            b.art = art;
            b.q0 = q0;
            b.p0 = p0;
            b.cmp0 = cmp0;
            b.pos0Excel = pos0Excel;
            b.fatt0 = q0 * p0;
            b.cogs0 = q0 * cmp0;
            b.pos0Calc = b.fatt0 - b.cogs0;
            baseByKey.put(key, b);
        }

        // pulisco grafici multi
        view.getChartsPanel().clearArticleCharts();

        // dettagli: sezioni multiple
        StringBuilder html = new StringBuilder();
        html.append("<html><body style='font-family:SansSerif;font-size:12px;'>");
        html.append("<div style='font-size:13px;'><b>Simulazione multi-articolo</b></div>");
        html.append("<div style='color:#666;'>Articoli selezionati: ").append(requests.size()).append("</div>");
        html.append("<hr style='border:none;border-top:1px solid #ddd;margin:10px 0;' />");

        // =========================================================
        // 4) Applico modifiche una per una sulla STESSA working copy
        // =========================================================
        for (SimulationControlsPanel.SimRequest req : requests) {

            String cat = (req.article.getCat() == null) ? "" : req.article.getCat().trim().toUpperCase();
            String art = (req.article.getArticolo() == null) ? "" : req.article.getArticolo().trim().toUpperCase();
            String key = cat + "||" + art;

            Base b = baseByKey.get(key);
            int rowIdx = b.rowIdx;

            double percent = req.percent;
            SimulationMode mode = req.mode;
            boolean doCompensate = req.compensate;

            // reset alla base per questo articolo (consistenza)
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, b.q0);
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, b.p0);
            eval.evaluateAll();

            // Step1 variazione
            double q1 = b.q0;
            double p1 = b.p0;

            if (mode == SimulationMode.QUANTITY) {
                q1 = b.q0 * (1.0 + percent / 100.0);
                ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, q1);
            } else {
                p1 = b.p0 * (1.0 + percent / 100.0);
                ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, p1);
            }

            eval.evaluateAll();
            double pos1Excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);

            double fatt1 = q1 * p1;
            double cogs1 = q1 * b.cmp0;
            double pos1Calc = fatt1 - cogs1;

            // aggiorna celle nella working copy (fatt/cogs) se presenti + cache formula + recalc
            if (colFatt != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colFatt, fatt1);
            if (colCogs != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colCogs, cogs1);

            evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colFatt);
            evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colCogs);
            evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colPos);

            forceExcelRecalcOnOpen(wb);
            eval.evaluateAll();

            // Step2 compensazione (opzionale): POS costante
            Double compValue = null;
            String compensatedVarLabel = null;
            double pos2Excel = Double.NaN;
            double qStar = Double.NaN;
            double pStar = Double.NaN;
            double fattStar = Double.NaN;
            double cogsStar = Double.NaN;
            double posStarCalc = Double.NaN;

            if (doCompensate) {

                if (mode == SimulationMode.QUANTITY) {
                    compValue = b.cmp0 + (b.pos0Calc / q1);
                    compensatedVarLabel = "P medio (€/kg)";
                    ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, compValue);
                    qStar = q1;
                    pStar = compValue;
                } else {
                    double denom = (p1 - b.cmp0);
                    if (Math.abs(denom) < 1e-12) throw new IllegalStateException("Compensazione impossibile: P1 - CMP0 = 0.");
                    compValue = b.pos0Calc / denom;
                    if (compValue <= 0) throw new IllegalStateException("Compensazione impossibile: Q* <= 0 (" + compValue + ").");
                    compensatedVarLabel = "Quantità (kg)";
                    ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, compValue);
                    qStar = compValue;
                    pStar = p1;
                }

                eval.evaluateAll();
                pos2Excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);

                fattStar = qStar * pStar;
                cogsStar = qStar * b.cmp0;
                posStarCalc = fattStar - cogsStar;

                if (colFatt != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colFatt, fattStar);
                if (colCogs != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colCogs, cogsStar);

                evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colFatt);
                evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colCogs);
                evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colPos);

                forceExcelRecalcOnOpen(wb);
                eval.evaluateAll();
            }

            // ===== Dettagli sezione articolo =====
            html.append("<div style='font-size:13px;'><b>")
                    .append(art).append("</b> <span style='color:#666;'>[").append(cat).append("]</span></div>");
            html.append("<div><b>Leva:</b> ").append(mode == SimulationMode.QUANTITY ? "Quantità" : "Prezzo")
                    .append(" &nbsp; <b>%:</b> ").append(String.format(java.util.Locale.US, "%.2f", percent))
                    .append("% &nbsp; <b>Compensa:</b> ")
                    .append(doCompensate ? "<span style='color:#1b5e20;'><b>SI</b></span>" : "<span style='color:#b71c1c;'><b>NO</b></span>")
                    .append("</div>");

            html.append("<table style='border-collapse:collapse;width:100%;margin-top:6px;'>");
            html.append(rowHtml("Q0", String.format(java.util.Locale.US, "%,.0f kg", b.q0)));
            html.append(rowHtml("P0", String.format(java.util.Locale.US, "%,.3f €/kg", b.p0)));
            html.append(rowHtml("CMP0", String.format(java.util.Locale.US, "%,.3f €/kg", b.cmp0)));
            html.append(rowHtml("POS0 (calc)", String.format(java.util.Locale.US, "%,.0f", b.pos0Calc)));
            html.append(rowHtml("Q1", String.format(java.util.Locale.US, "%,.0f kg", q1)));
            html.append(rowHtml("P1", String.format(java.util.Locale.US, "%,.3f €/kg", p1)));
            html.append(rowHtml("POS1 (calc)", String.format(java.util.Locale.US, "%,.0f", pos1Calc)));
            if (doCompensate) {
                html.append(rowHtml("Variabile compensata", compensatedVarLabel));
                html.append(rowHtml("Valore compensazione", (mode == SimulationMode.QUANTITY)
                        ? String.format(java.util.Locale.US, "%,.3f €/kg", compValue)
                        : String.format(java.util.Locale.US, "%,.0f kg", compValue)));
                html.append(rowHtml("POS* (calc)", String.format(java.util.Locale.US, "%,.0f", posStarCalc)));
            }
            html.append("</table>");
            html.append("<hr style='border:none;border-top:1px solid #eee;margin:10px 0;' />");

            // ===== Grafici per articolo (tab) =====
            DefaultCategoryDataset posDS = new DefaultCategoryDataset();
            DefaultCategoryDataset compDS = new DefaultCategoryDataset();

            posDS.addValue(b.pos0Excel, "POS (Excel)", "Originale");
            posDS.addValue(pos1Excel, "POS (Excel)", "Dopo variazione");
            posDS.addValue(b.pos0Calc, "POS (calc)", "Originale");
            posDS.addValue(pos1Calc, "POS (calc)", "Dopo variazione");
            if (doCompensate) {
                posDS.addValue(pos2Excel, "POS (Excel)", "Dopo compensazione");
                posDS.addValue(posStarCalc, "POS (calc)", "Dopo compensazione");
            }

            if (mode == SimulationMode.QUANTITY) {
                compDS.addValue(b.q0, "Quantità (kg)", "Originale");
                compDS.addValue(q1, "Quantità (kg)", "Dopo variazione");
                compDS.addValue(b.p0, "P medio (€/kg)", "Originale");
                compDS.addValue(b.p0, "P medio (€/kg)", "Dopo variazione");
                if (doCompensate) {
                    compDS.addValue(q1, "Quantità (kg)", "Dopo compensazione");
                    compDS.addValue(compValue, "P medio (€/kg)", "Dopo compensazione");
                }
            } else {
                compDS.addValue(b.p0, "P medio (€/kg)", "Originale");
                compDS.addValue(p1, "P medio (€/kg)", "Dopo variazione");
                compDS.addValue(b.q0, "Quantità (kg)", "Originale");
                compDS.addValue(b.q0, "Quantità (kg)", "Dopo variazione");
                if (doCompensate) {
                    compDS.addValue(p1, "P medio (€/kg)", "Dopo compensazione");
                    compDS.addValue(compValue, "Quantità (kg)", "Dopo compensazione");
                }
            }

            JFreeChart posChart = ChartFactory.createBarChart(
                    "POS – " + art,
                    "Scenario",
                    "POS",
                    posDS
            );

            String compTitle = doCompensate ? ("Compensazione – " + art) : ("Variazione – " + art);
            JFreeChart compChart = ChartFactory.createBarChart(
                    compTitle,
                    "Scenario",
                    "Valori (Q e P)",
                    compDS
            );

            configureCategoryChart(posChart, true);
            applySeriesLabelFormatting(posChart);

            configureCategoryChart(compChart, false);
            applySeriesLabelFormatting(compChart);


            // tabKey stabile: cat||art
            view.getChartsPanel().addOrReplaceArticleCharts(key, art, posChart, compChart);
        }

        html.append("</body></html>");
        view.getControlsPanel().setDetails(html.toString());

        // salva workbook finale batch
        forceExcelRecalcOnOpen(wb);
        eval.evaluateAll();
        excelRepo.safeSaveWorkbook(wb);

    } catch (Exception ex) {
        log.error("Errore simulazione multi", ex);
        JOptionPane.showMessageDialog(view, "Errore simulazione: " + ex.getMessage(), "Errore",
                JOptionPane.ERROR_MESSAGE);
    }
}

    // ===========================
    // Chart config
    // ===========================
   private void configureCategoryChart(JFreeChart chart, boolean integerValues) {
    CategoryPlot plot = chart.getCategoryPlot();

    // --- estetica base
    chart.setBackgroundPaint(Color.WHITE);
    plot.setBackgroundPaint(new Color(250, 250, 250));
    plot.setOutlineVisible(false);
    plot.setRangeGridlinePaint(new Color(220, 220, 220));
    plot.setDomainGridlinePaint(new Color(220, 220, 220));
    plot.setRangeGridlinesVisible(true);
    plot.setDomainGridlinesVisible(true);

    // --- font
    Font axisFont = new Font("SansSerif", Font.PLAIN, 12);
    Font tickFont = new Font("SansSerif", Font.PLAIN, 11);

    CategoryAxis domain = plot.getDomainAxis();
    domain.setLabelFont(axisFont);
    domain.setTickLabelFont(tickFont);

    int cols = (plot.getDataset() != null) ? plot.getDataset().getColumnCount() : 0;
    if (cols > 4) {
        domain.setCategoryLabelPositions(
                CategoryLabelPositions.createUpRotationLabelPositions(Math.PI / 8.0)
        );
    } else {
        domain.setCategoryLabelPositions(CategoryLabelPositions.STANDARD);
    }

    // ==========================================================
    // ✅ FIX FORMATO: niente ".000" -> usa #,##0.###
    // ==========================================================
    NumberFormat fmt;
    if (integerValues) {
        DecimalFormat df = new DecimalFormat("#,##0");
        df.setGroupingUsed(true);
        fmt = df;
    } else {
        // max 3 decimali, min 0 => niente zeri finali
        DecimalFormat df = new DecimalFormat("#,##0.###");
        df.setGroupingUsed(true);
        df.setMinimumFractionDigits(0);
        df.setMaximumFractionDigits(3);
        fmt = df;
    }

    // --- asse Y
    if (plot.getRangeAxis() instanceof NumberAxis) {
        NumberAxis range = (NumberAxis) plot.getRangeAxis();
        range.setLabelFont(axisFont);
        range.setTickLabelFont(tickFont);

        range.setNumberFormatOverride(fmt);
        range.setAutoRangeIncludesZero(true);

        if (integerValues) {
            range.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
        }

        range.setUpperMargin(0.15);
        range.setLowerMargin(0.10);
    }

    // --- renderer: Bar chart
    if (plot.getRenderer() instanceof BarRenderer) {
        BarRenderer r = (BarRenderer) plot.getRenderer();
        r.setShadowVisible(false);
        r.setBarPainter(new StandardBarPainter());

        r.setDefaultItemLabelsVisible(true);
        r.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator("{2}", fmt));
        r.setDefaultItemLabelFont(new Font("SansSerif", Font.PLAIN, 11));
    }

    // --- renderer: Line chart
    if (plot.getRenderer() instanceof LineAndShapeRenderer) {
        LineAndShapeRenderer r = (LineAndShapeRenderer) plot.getRenderer();

        r.setDefaultShapesVisible(true);
        r.setDefaultItemLabelsVisible(true);
        r.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator("{2}", fmt));
        r.setDefaultStroke(new BasicStroke(2.0f));
    }
}

   private void applySeriesLabelFormatting(JFreeChart chart) {
	    if (chart == null) return;
	    if (chart.getPlot() == null) return;

	    if (chart.getPlot() instanceof org.jfree.chart.plot.CategoryPlot) {
	        org.jfree.chart.plot.CategoryPlot plot = (org.jfree.chart.plot.CategoryPlot) chart.getPlot();
	        if (plot.getRenderer() == null) return;

	        plot.getRenderer().setDefaultItemLabelsVisible(true);

	        plot.getRenderer().setDefaultItemLabelGenerator(new org.jfree.chart.labels.CategoryItemLabelGenerator() {
	            @Override
	            public String generateLabel(org.jfree.data.category.CategoryDataset dataset, int row, int column) {
	                if (dataset == null) return "";
	                Comparable<?> rowKey = dataset.getRowKey(row);
	                Number v = dataset.getValue(row, column);
	                if (v == null) return "";

	                String series = (rowKey == null) ? "" : rowKey.toString().toLowerCase();

	                // Quantità -> intero (come nei dettagli)
	                if (series.contains("quant") || series.contains("(kg)")) {
	                    return DF_INT.format(v.doubleValue());
	                }

	                // Prezzi / CMP -> 3 decimali fissi (come nei dettagli: 1,640)
	                if (series.contains("p medio") || series.contains("€/kg") || series.contains("cmp")) {
	                    return DF_3.format(v.doubleValue());
	                }

	                // Tutto il resto (POS, fatturato, cogs, ecc.) -> intero
	                return DF_INT.format(v.doubleValue());
	            }

	            @Override
	            public String generateRowLabel(org.jfree.data.category.CategoryDataset dataset, int row) {
	                return dataset.getRowKey(row).toString();
	            }

	            @Override
	            public String generateColumnLabel(org.jfree.data.category.CategoryDataset dataset, int column) {
	                return dataset.getColumnKey(column).toString();
	            }
	        });
	    }
	}

    // ===========================
    // CE Budget 2022 reading: SOLO colonna J
    // ===========================
    private Map<String, Double> readCeBudgetSnapshot(Workbook wb, FormulaEvaluator eval) {
        Sheet ce = findCeBudgetSheet(wb);
        if (ce == null) {
            throw new IllegalStateException("Foglio CE Budget 2022 non trovato (nome contenente 'CE' e 'BUDGET').");
        }

        eval.evaluateAll();

        Map<String, Double> out = new HashMap<>();
        out.put(K_RICAVI_PF,    readNumericCell(ce, eval, CE_ROW_RICAVI_PF,   CE_COL_J));
        out.put(K_RICAVI_MP,    readNumericCell(ce, eval, CE_ROW_RICAVI_MP,   CE_COL_J));
        out.put(K_RICAVI_CLAV,  readNumericCell(ce, eval, CE_ROW_RICAVI_CLAV, CE_COL_J));
        out.put(K_ALTRI_RICAVI, readNumericCell(ce, eval, CE_ROW_ALTRI_RICAVI,CE_COL_J));
        out.put(K_VAR_PF,       readNumericCell(ce, eval, CE_ROW_VAR_PF,      CE_COL_J));
        out.put(K_ACQUISTO_MP,  readNumericCell(ce, eval, CE_ROW_ACQUISTO_MP, CE_COL_J));
        out.put(K_VAR_SCORTE,   readNumericCell(ce, eval, CE_ROW_VAR_SCORTE,  CE_COL_J));

        return out;
    }

    private double readNumericCell(Sheet sh, FormulaEvaluator eval, int rowIdx, int colIdx) {
        Row row = sh.getRow(rowIdx);
        if (row == null) return 0.0;

        Cell c = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (c == null) return 0.0;

        if (c.getCellType() == CellType.NUMERIC) return c.getNumericCellValue();

        if (c.getCellType() == CellType.FORMULA) {
            CellValue cv = eval.evaluate(c);
            if (cv != null && cv.getCellType() == CellType.NUMERIC) return cv.getNumberValue();
            return 0.0;
        }

        if (c.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(c.getStringCellValue().trim().replace(".", "").replace(",", "."));
            } catch (Exception ignore) {
                return 0.0;
            }
        }

        return 0.0;
    }

    private Sheet findCeBudgetSheet(Workbook wb) {

        String[] candidates = {
                "CE-Budget-2022",
                "CE BUDGET 2022",
                "CE_BUDGET_2022",
                "CE Budget 2022",
                "CE BUDGET2022",
                "CEBudget2022"
        };
        for (String n : candidates) {
            Sheet s = wb.getSheet(n);
            if (s != null) return s;
        }

        DataFormatter fmt = new DataFormatter();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet s = wb.getSheetAt(i);
            if (sheetLooksLikeCeBudget2022(s, fmt)) return s;
        }

        return null;
    }

    private boolean sheetLooksLikeCeBudget2022(Sheet s, DataFormatter fmt) {
        int maxRows = Math.min(s.getLastRowNum(), 40);
        int maxCols = 25;

        for (int r = 0; r <= maxRows; r++) {
            Row row = s.getRow(r);
            if (row == null) continue;

            int last = Math.min(row.getLastCellNum(), maxCols);
            for (int c = 0; c < last; c++) {
                String txt = fmt.formatCellValue(row.getCell(c)).trim().toUpperCase();
                if (txt.isEmpty()) continue;

                String norm = txt.replaceAll("\\s+", " ");
                if (norm.contains("CE") && norm.contains("BUDGET") && norm.contains("2022")) return true;
            }
        }
        return false;
    }

    private String prettifyCeKey(String k) {
        if (K_RICAVI_PF.equals(k))    return "Ricavi PF";
        if (K_RICAVI_MP.equals(k))    return "Ricavi MP";
        if (K_RICAVI_CLAV.equals(k))  return "Ricavi C/Lav.";
        if (K_ALTRI_RICAVI.equals(k)) return "Altri ricavi";
        if (K_VAR_PF.equals(k))       return "Var. PF";
        if (K_ACQUISTO_MP.equals(k))  return "Acquisto MP";
        if (K_VAR_SCORTE.equals(k))   return "Var. scorte";
        return k;
    }

    // ===========================
    // FIX: CE "Dopo variazione" calcolato (non dipende dal foglio CE)
    // NB: per MP: Acquisto MP va con segno coerente ai valori in CE (costi negativi)
    // ===========================
    private Map<String, Double> computeCeAfterVar(
            Map<String, Double> ceBase,
            String targetCat,
            double fatt0, double fatt1,
            double cogs0, double cogs1
    ) {
        Map<String, Double> out = new HashMap<>(ceBase);

        double dRicavi = fatt1 - fatt0;
        double dCosti  = cogs1 - cogs0;

        String cat = (targetCat == null) ? "" : targetCat.trim().toUpperCase();

        if (cat.contains("MP")) {
            out.put(K_RICAVI_MP,   ceBase.getOrDefault(K_RICAVI_MP, 0.0) + dRicavi);
            out.put(K_ACQUISTO_MP, ceBase.getOrDefault(K_ACQUISTO_MP, 0.0) - dCosti);
        } else {
            out.put(K_RICAVI_PF, ceBase.getOrDefault(K_RICAVI_PF, 0.0) + dRicavi);
        }

        return out;
    }

    private void onExit() {
        excelRepo.cleanup();
        System.exit(0);
    }

 // =====================================================================
 // ✅ CE Budget 2022: finestra con grafico aggiornato
 // - legge dalla WORKING COPY (quella modificata)
 // - così il CE riflette le simulazioni appena fatte
 // =====================================================================
 private void onShowCeBudgetBase() {

     File working = model.getWorkingExcelCopy();
     if (working == null || !working.exists()) {
         JOptionPane.showMessageDialog(view, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
         return;
     }

     try (Workbook wb = WorkbookFactory.create(working)) {

         Sheet ce = findCeBudgetSheet(wb);
         if (ce == null) throw new IllegalStateException("Foglio CE Budget 2022 non trovato.");

         FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
         DataFormatter fmt = new DataFormatter();

         // ✅ Importantissimo: rivaluta le formule del file (se il CE contiene formule)
         eval.evaluateAll();

         LinkedHashMap<String, Double> v = new LinkedHashMap<>();

         v.put("Ricavi PF", findValueByRowLabel(ce, eval, fmt, "RICAVI DELLE VENDITE DI PRODOTTI FINITI"));
         v.put("Ricavi MP", findValueByRowLabel(ce, eval, fmt, "RICAVI DELLE VENDITE DI MATERIE PRIME"));
         v.put("Ricavi C/Lav.", findValueByRowLabel(ce, eval, fmt, "RICAVI CONTO LAVORAZIONE"));
         v.put("Altri ricavi", findValueByRowLabel(ce, eval, fmt, "ALTRI RICAVI"));
         v.put("Var. PF", findValueByRowLabel(ce, eval, fmt, "VARIAZIONE PRODOTTI FINITI"));
         v.put("Tot. Ricavi produzione (A)", findValueByRowLabel(ce, eval, fmt, "TOTALE RICAVI PRODUZIONE"));

         v.put("Acquisto MP", findValueByRowLabel(ce, eval, fmt, "ACQUISTO MATERIE PRIME"));
         v.put("Var. scorte", findValueByRowLabel(ce, eval, fmt, "VARIAZIONE SCORTE"));
         v.put("Tot. Costi MP (B)", findValueByRowLabel(ce, eval, fmt, "TOTALE COSTI MATERIE PRIME"));

         v.put("Costo energia", findValueByRowLabel(ce, eval, fmt, "COSTO ENERGIA"));
         v.put("Materiali di consumo", findValueByRowLabel(ce, eval, fmt, "MATERIALI DI CONSUMO"));
         v.put("Pulizia/smaltimento", findValueByRowLabel(ce, eval, fmt, "PULIZIA"));
         v.put("Tot. Costi variabili prod. (C)", findValueByRowLabel(ce, eval, fmt, "COSTI VARIABILI DI PRODUZIONE"));

         v.put("Trasporti/oneri vendita+acquisto", findValueByRowLabel(ce, eval, fmt, "TRASPORTI"));
         v.put("Provvigioni/Enasarco", findValueByRowLabel(ce, eval, fmt, "PROVVIGIONI"));
         v.put("Tot. Costi di vendita (D)", findValueByRowLabel(ce, eval, fmt, "TOTALE COSTI DI VENDITA"));

         v.put("MOL (A-B-C-D)", findValueByRowLabel(ce, eval, fmt, "MARGINE OPERATIVO LORDO"));

         DefaultCategoryDataset ds = new DefaultCategoryDataset();
         for (Map.Entry<String, Double> e : v.entrySet()) {
             ds.addValue(e.getValue(), "CE Budget 2022 (working copy)", e.getKey());
         }

         JFreeChart chart = ChartFactory.createLineChart(
                 "CE Budget 2022 – Aggiornato (Working copy)",
                 "Voce",
                 "Valore",
                 ds,
                 PlotOrientation.VERTICAL,
                 true,
                 true,
                 false
         );
         configureCategoryChart(chart, true);

         CeBudgetFrame ceFrame = new CeBudgetFrame();
         ceFrame.setChart(chart);

         view.setVisible(false);

         ceFrame.getBtnBack().addActionListener(ev -> {
             ceFrame.dispose();
             view.setVisible(true);
         });

         ceFrame.addWindowListener(new java.awt.event.WindowAdapter() {
             @Override
             public void windowClosed(java.awt.event.WindowEvent e) {
                 view.setVisible(true);
             }
         });

         ceFrame.setVisible(true);

     } catch (Exception ex) {
         log.error("Errore apertura CE Budget (working copy)", ex);
         JOptionPane.showMessageDialog(view, "Errore: " + ex.getMessage(), "Errore", JOptionPane.ERROR_MESSAGE);
     }
 }


    private double findValueByRowLabel(Sheet sh, FormulaEvaluator eval, DataFormatter fmt, String labelNeedle) {

        String needle = labelNeedle.trim().toUpperCase();
        int maxRows = Math.min(sh.getLastRowNum(), 200);

        for (int r = 0; r <= maxRows; r++) {
            Row row = sh.getRow(r);
            if (row == null) continue;

            for (int c = 0; c <= 8; c++) {
                String txt = fmt.formatCellValue(row.getCell(c)).trim().toUpperCase();
                if (txt.isEmpty()) continue;

                if (txt.contains(needle)) {
                    return readNumericCell(sh, eval, r, 9);
                }
            }
        }

        log.warn("Voce CE non trovata nel foglio: '{}'", labelNeedle);
        return 0.0;
    }
}
