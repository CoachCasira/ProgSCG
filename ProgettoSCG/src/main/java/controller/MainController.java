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
        view.getBtnShowPremioComp().addActionListener(e -> onOpenPremioComp());

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

    private void onSimulate() {

        if (model.getWorkingExcelCopy() == null || ricaviService == null) {
            JOptionPane.showMessageDialog(view, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
            return;
        }

        ArticleRow ar = view.getControlsPanel().getSelectedArticle();
        if (ar == null) return;

        SimulationMode mode = view.getControlsPanel().getMode();

        // ✅ nuova opzione: esegui o meno la compensazione
        boolean doCompensate = view.getControlsPanel().isCompensateSelected();

        double percent;
        try {
            percent = view.getControlsPanel().getPercent();
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view, e.getMessage(), "Input non valido", JOptionPane.WARNING_MESSAGE);
            return;
        }

        log.info("Simula articolo={} mode={} percent={} compensate={}", ar, mode, percent, doCompensate);

        try (Workbook wb = WorkbookFactory.create(model.getWorkingExcelCopy())) {

            Sheet ricaviSheet = wb.getSheet("Ricavi");
            if (ricaviSheet == null) throw new IllegalStateException("Foglio 'Ricavi' non trovato.");

            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
            DataFormatter fmt = new DataFormatter();

            // =========================================================
            // 0) Snapshot CE "BUDGET BASE" (prima di toccare Ricavi)
            // =========================================================
            Map<String, Double> ceBase = readCeBudgetSnapshot(wb, eval);

            // =========================================================
            // 1) Trovo header tabella destra e colonne chiave
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

                // facoltativi: se presenti, li aggiorniamo in working copy
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
            // 2) Trovo la riga corretta per Cat/Articolo
            // =========================================================
            String targetCat = (ar.getCat() == null) ? "" : ar.getCat().trim().toUpperCase();
            String targetArt = (ar.getArticolo() == null) ? "" : ar.getArticolo().trim().toUpperCase();

            int rowIdx = -1;
            for (int r = headerRowIdx + 1; r <= ricaviSheet.getLastRowNum(); r++) {
                Row rr = ricaviSheet.getRow(r);
                if (rr == null) continue;

                String catTxt = fmt.formatCellValue(rr.getCell(colCat)).trim().toUpperCase();
                String artTxt = fmt.formatCellValue(rr.getCell(colArt)).trim().toUpperCase();

                if (catTxt.equals(targetCat) && artTxt.equals(targetArt)) {
                    rowIdx = r;
                    break;
                }
            }

            if (rowIdx < 0) {
                throw new IllegalStateException(
                        "Non trovo la riga per Cat='" + targetCat + "' e Articolo='" + targetArt + "' nella tabella destra."
                );
            }

            // =========================================================
            // 3) Leggo valori base (EUR) + POS Excel
            // =========================================================
            eval.evaluateAll();

            double q0   = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colQty);
            double p0   = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPeur);
            double cmp0 = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colCMPeur);
            double pos0Excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);

            if (q0 <= 0)   throw new IllegalStateException("Q0 non valida letta da Excel: " + q0);
            if (p0 <= 0)   throw new IllegalStateException("P0 (€/kg) non valido letto da Excel: " + p0);
            if (cmp0 <= 0) throw new IllegalStateException("CMP0 (€/kg) non valido letto da Excel: " + cmp0);

            double fatt0 = q0 * p0;
            double cogs0 = q0 * cmp0;
            double pos0Calc = fatt0 - cogs0;

            // =========================================================
            // 4) Step 1: applico variazione
            // =========================================================
            // riallineo prima alle condizioni base
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, q0);
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, p0);
            eval.evaluateAll();

            double q1 = q0;
            double p1 = p0;

            if (mode == SimulationMode.QUANTITY) {
                q1 = q0 * (1.0 + percent / 100.0);
                ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, q1);
            } else {
                p1 = p0 * (1.0 + percent / 100.0);
                ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, p1);
            }

            eval.evaluateAll();

            double pos1Excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);

            double fatt1 = q1 * p1;
            double cogs1 = q1 * cmp0;
            double pos1Calc = fatt1 - cogs1;

            // =========================================================
            // ✅ FIX EXCEL: dopo VARIAZIONE aggiorna celle & cache
            // =========================================================
            if (colFatt != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colFatt, fatt1);
            if (colCogs != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colCogs, cogs1);

            evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colFatt);
            evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colCogs);
            evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colPos);

            forceExcelRecalcOnOpen(wb);
            eval.evaluateAll();

            // =========================================================
            // 4B) CE dopo variazione (FIX): calcolato, non riletto da Excel
            // =========================================================
            Map<String, Double> ceAfterVar = computeCeAfterVar(ceBase, targetCat, fatt0, fatt1, cogs0, cogs1);

            // =========================================================
            // 5) Step 2: compensazione per mantenere POS(calc) costante
            // =========================================================
            Double compValue = null;
            String compensatedVarLabel = null;
            double pos2Excel = Double.NaN;
            double qStar = Double.NaN;
            double pStar = Double.NaN;
            double fattStar = Double.NaN;
            double cogsStar = Double.NaN;
            double posStarCalc = Double.NaN;

            if (doCompensate) {
                ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, q1);
                ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, p1);
                eval.evaluateAll();

                if (mode == SimulationMode.QUANTITY) {
                    // compenso su Prezzo
                    compValue = cmp0 + (pos0Calc / q1);
                    compensatedVarLabel = "P medio (€/kg)";
                    ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, compValue);
                } else {
                    // compenso su Quantità
                    double denom = (p1 - cmp0);
                    if (Math.abs(denom) < 1e-12) {
                        throw new IllegalStateException("Compensazione impossibile: P1 - CMP0 = 0.");
                    }
                    compValue = pos0Calc / denom;
                    if (compValue <= 0) {
                        throw new IllegalStateException("Compensazione impossibile: Q* <= 0 (" + compValue + ").");
                    }
                    compensatedVarLabel = "Quantità (kg)";
                    ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, compValue);
                }

                eval.evaluateAll();
                pos2Excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);

                qStar = (mode == SimulationMode.QUANTITY) ? q1 : compValue;
                pStar = (mode == SimulationMode.QUANTITY) ? compValue : p1;

                fattStar = qStar * pStar;
                cogsStar = qStar * cmp0;
                posStarCalc = fattStar - cogsStar;

                // =========================================================
                // ✅ FIX EXCEL: dopo COMPENSAZIONE aggiorna celle & cache
                // =========================================================
                if (colFatt != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colFatt, fattStar);
                if (colCogs != null) writeNumericIfNotFormula(ricaviSheet, rowIdx, colCogs, cogsStar);

                evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colFatt);
                evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colCogs);
                evaluateFormulaCellIfPresent(ricaviSheet, eval, rowIdx, colPos);

                forceExcelRecalcOnOpen(wb);
                eval.evaluateAll();
            }

            // =========================================================
            // 6) Dettagli (HTML)
            // =========================================================
            String whatChanged = (mode == SimulationMode.QUANTITY) ? "Quantità (kg)" : "Prezzo (€/kg)";
            String unitPrice = "€/kg";

            StringBuilder html = new StringBuilder();
            html.append("<html><body style='font-family:SansSerif;font-size:12px;'>");
            html.append("<div style='font-size:13px;'><b>Articolo:</b> ").append(targetArt)
                    .append(" <span style='color:#666;'>[").append(targetCat).append("]</span></div>");
            html.append("<div><b>Percentuale applicata:</b> ").append(DF_2.format(percent)).append("%</div>");
            html.append("<div><b>Compensazione:</b> ")
                    .append(doCompensate ? "<span style='color:#1b5e20;'><b>SI</b></span>" : "<span style='color:#b71c1c;'><b>NO</b></span>")
                    .append("</div>");
            html.append("<hr style='border:none;border-top:1px solid #ddd;margin:10px 0;' />");

            // Tabella base
            html.append("<div style='margin-bottom:6px;'><b>Valori originali</b></div>");
            html.append("<table style='border-collapse:collapse;width:100%;'>");
            html.append(rowHtml("Q0", DF_INT.format(q0) + " kg"));
            html.append(rowHtml("P0", DF_3.format(p0) + " " + unitPrice));
            html.append(rowHtml("CMP0", DF_3.format(cmp0) + " " + unitPrice));
            html.append(rowHtml("Fatturato0 = Q0·P0", DF_INT.format(fatt0)));
            html.append(rowHtml("COGS0 = Q0·CMP0", DF_INT.format(cogs0)));
            html.append(rowHtml("POS0 (Excel)", DF_INT.format(pos0Excel)));
            html.append(rowHtml("POS0 (calc)", DF_INT.format(pos0Calc)));
            html.append("</table>");

            // Step 1
            html.append("<hr style='border:none;border-top:1px solid #eee;margin:10px 0;' />");
            html.append("<div style='margin-bottom:6px;'><b>Step 1 — Dopo variazione</b> <span style='color:#666;'>(senza compensazione)</span></div>");
            html.append("<table style='border-collapse:collapse;width:100%;'>");
            html.append(rowHtml("Leva modificata", whatChanged));
            html.append(rowHtml("Q1", DF_INT.format(q1) + " kg"));
            html.append(rowHtml("P1", DF_3.format(p1) + " " + unitPrice));
            html.append(rowHtml("Fatturato1 = Q1·P1", DF_INT.format(fatt1)));
            html.append(rowHtml("COGS1 = Q1·CMP0", DF_INT.format(cogs1)));
            html.append(rowHtml("POS1 (Excel)", DF_INT.format(pos1Excel)));
            html.append(rowHtml("POS1 (calc)", DF_INT.format(pos1Calc)));
            html.append("</table>");

            // Step 2 (solo se richiesto)
            if (doCompensate) {
                html.append("<hr style='border:none;border-top:1px solid #eee;margin:10px 0;' />");
                html.append("<div style='margin-bottom:6px;'><b>Step 2 — Compensazione</b> <span style='color:#666;'>(mantieni POS costante)</span></div>");
                html.append("<table style='border-collapse:collapse;width:100%;'>");
                html.append(rowHtml("Target POS(calc)", DF_INT.format(pos0Calc)));
                html.append(rowHtml("Variabile compensata", compensatedVarLabel));

                if (mode == SimulationMode.QUANTITY) {
                    double deltaPct = ((compValue / p0) - 1.0) * 100.0;
                    html.append(rowHtml("P*", DF_3.format(compValue) + " " + unitPrice));
                    html.append(rowHtml("ΔP%", DF_2.format(deltaPct) + "%"));
                } else {
                    double deltaPct = ((compValue / q0) - 1.0) * 100.0;
                    html.append(rowHtml("Q*", DF_INT.format(compValue) + " kg"));
                    html.append(rowHtml("ΔQ%", DF_2.format(deltaPct) + "%"));
                }

                html.append("</table>");

                html.append("<hr style='border:none;border-top:1px solid #eee;margin:10px 0;' />");
                html.append("<div style='margin-bottom:6px;'><b>Check finale</b></div>");
                html.append("<table style='border-collapse:collapse;width:100%;'>");
                html.append(rowHtml("Fatturato* = Q*·P*", DF_INT.format(fattStar)));
                html.append(rowHtml("COGS* = Q*·CMP0", DF_INT.format(cogsStar)));
                html.append(rowHtml("POS dopo compensazione (Excel)", DF_INT.format(pos2Excel)));
                html.append(rowHtml("POS dopo compensazione (calc)", DF_INT.format(posStarCalc)));
                html.append(rowHtml("Errore |POS(calc) - POS0(calc)|", DF_INT.format(Math.abs(posStarCalc - pos0Calc))));
                html.append("</table>");
            }

            html.append("</body></html>");
            view.getControlsPanel().setDetails(html.toString());

            // =========================================================
            // 7) Grafici POS + Comp
            // =========================================================
            DefaultCategoryDataset posDS = new DefaultCategoryDataset();
            DefaultCategoryDataset compDS = new DefaultCategoryDataset();

            posDS.addValue(pos0Excel, "POS (Excel)", "Originale");
            posDS.addValue(pos1Excel, "POS (Excel)", "Dopo variazione");

            posDS.addValue(pos0Calc, "POS (calc)", "Originale");
            posDS.addValue(pos1Calc, "POS (calc)", "Dopo variazione");

            if (doCompensate) {
                posDS.addValue(pos2Excel, "POS (Excel)", "Dopo compensazione");
                posDS.addValue(posStarCalc, "POS (calc)", "Dopo compensazione");
            }

            if (mode == SimulationMode.QUANTITY) {
                compDS.addValue(q0, "Quantità (kg)", "Originale");
                compDS.addValue(q1, "Quantità (kg)", "Dopo variazione");

                compDS.addValue(p0, "P medio (€/kg)", "Originale");
                compDS.addValue(p0, "P medio (€/kg)", "Dopo variazione");

                if (doCompensate) {
                    compDS.addValue(q1, "Quantità (kg)", "Dopo compensazione");
                    compDS.addValue(compValue, "P medio (€/kg)", "Dopo compensazione");
                }
            } else {
                compDS.addValue(p0, "P medio (€/kg)", "Originale");
                compDS.addValue(p1, "P medio (€/kg)", "Dopo variazione");

                compDS.addValue(q0, "Quantità (kg)", "Originale");
                compDS.addValue(q0, "Quantità (kg)", "Dopo variazione");

                if (doCompensate) {
                    compDS.addValue(p1, "P medio (€/kg)", "Dopo compensazione");
                    compDS.addValue(compValue, "Quantità (kg)", "Dopo compensazione");
                }
            }

            JFreeChart posChart = ChartFactory.createBarChart(
                    "POS – " + targetArt,
                    "Scenario",
                    "POS",
                    posDS
            );

            String compTitle = doCompensate ? ("Compensazione – " + targetArt) : ("Variazione – " + targetArt);
            JFreeChart compChart = ChartFactory.createBarChart(
                    compTitle,
                    "Scenario",
                    "Valori (Q e P)",
                    compDS
            );

            configureCategoryChart(posChart, true);
            configureCategoryChart(compChart, false);

            view.getChartsPanel().setPosChart(posChart);
            view.getChartsPanel().setCompChart(compChart);

            // =========================================================
            // 8) CE: tre grafici separati (BASE, DOPO VAR, DELTA)
            // =========================================================
            List<String> keysOrdered = Arrays.asList(
                    K_RICAVI_PF,
                    K_RICAVI_MP,
                    K_RICAVI_CLAV,
                    K_ALTRI_RICAVI,
                    K_VAR_PF,
                    K_ACQUISTO_MP,
                    K_VAR_SCORTE
            );

            // --- CE BASE
            DefaultCategoryDataset ceBaseDS = new DefaultCategoryDataset();
            for (String k : keysOrdered) {
                String catLabel = prettifyCeKey(k);
                ceBaseDS.addValue(ceBase.getOrDefault(k, 0.0), "Budget base", catLabel);
            }
            JFreeChart ceBaseChart = ChartFactory.createLineChart(
                    "CE Budget 2022 – Valori (Budget base)",
                    "Voce",
                    "Valore",
                    ceBaseDS
            );
            configureCategoryChart(ceBaseChart, true);
            view.getChartsPanel().setCeBaseChart(ceBaseChart);

            // --- CE DOPO VARIAZIONE
            DefaultCategoryDataset ceVarDS = new DefaultCategoryDataset();
            for (String k : keysOrdered) {
                String catLabel = prettifyCeKey(k);
                ceVarDS.addValue(ceAfterVar.getOrDefault(k, 0.0), "Dopo variazione", catLabel);
            }
            JFreeChart ceVarChart = ChartFactory.createLineChart(
                    "CE Budget 2022 – Valori (Dopo variazione)",
                    "Voce",
                    "Valore",
                    ceVarDS
            );
            configureCategoryChart(ceVarChart, true);
            view.getChartsPanel().setCeVarChart(ceVarChart);

            // --- CE DELTA (dopo - base)
            DefaultCategoryDataset ceDeltaDS = new DefaultCategoryDataset();
            for (String k : keysOrdered) {
                String catLabel = prettifyCeKey(k);
                double baseV = ceBase.getOrDefault(k, 0.0);
                double varV  = ceAfterVar.getOrDefault(k, 0.0);
                ceDeltaDS.addValue(varV - baseV, "Δ (Dopo - Base)", catLabel);
            }
            JFreeChart ceDeltaChart = ChartFactory.createLineChart(
                    "CE Budget 2022 – Variazioni (Δ)",
                    "Voce",
                    "Δ Valore",
                    ceDeltaDS
            );
            configureCategoryChart(ceDeltaChart, false);
            view.getChartsPanel().setCeDeltaChart(ceDeltaChart);

            // =========================================================
            // 9) Salvo (forza recalc prima di salvare)
            // =========================================================
            forceExcelRecalcOnOpen(wb);
            eval.evaluateAll();

            excelRepo.safeSaveWorkbook(wb);

        } catch (Exception ex) {
            log.error("Errore simulazione", ex);
            JOptionPane.showMessageDialog(view, "Errore simulazione: " + ex.getMessage(), "Errore",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    // ===========================
    // Chart config
    // ===========================
    private void configureCategoryChart(JFreeChart chart, boolean integerValues) {
        CategoryPlot plot = chart.getCategoryPlot();

        chart.setBackgroundPaint(Color.WHITE);
        plot.setBackgroundPaint(new Color(250, 250, 250));
        plot.setOutlineVisible(false);
        plot.setRangeGridlinePaint(new Color(220, 220, 220));
        plot.setDomainGridlinePaint(new Color(220, 220, 220));
        plot.setRangeGridlinesVisible(true);
        plot.setDomainGridlinesVisible(true);

        Font axisFont = new Font("SansSerif", Font.PLAIN, 12);
        Font tickFont = new Font(" >> ",Font.PLAIN,12);

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

        NumberFormat fmt = integerValues
                ? new DecimalFormat("#,##0")
                : new DecimalFormat("#,##0.000");

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

        if (plot.getRenderer() instanceof BarRenderer) {
            BarRenderer r = (BarRenderer) plot.getRenderer();
            r.setShadowVisible(false);
            r.setBarPainter(new StandardBarPainter());
            r.setDefaultItemLabelsVisible(true);
            r.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator("{2}", fmt));
            r.setDefaultItemLabelFont(new Font("SansSerif", Font.PLAIN, 11));
        }

        if (plot.getRenderer() instanceof LineAndShapeRenderer) {
            LineAndShapeRenderer r = (LineAndShapeRenderer) plot.getRenderer();
            r.setDefaultShapesVisible(true);
            r.setDefaultItemLabelsVisible(true);
            r.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator("{2}", fmt));
            r.setDefaultStroke(new BasicStroke(2.0f));
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
