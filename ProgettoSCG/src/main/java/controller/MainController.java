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
import org.jfree.chart.axis.ValueAxis;
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

public class MainController {

    private static final Logger log = LogManager.getLogger(MainController.class);

    private final AppModel model;
    private final MainFrame view;
    private final ExcelRepository excelRepo;

    private RicaviExcelService ricaviService;
    private List<ArticleRow> cachedArticles = new ArrayList<>();

    private static final DecimalFormat DF_INT = new DecimalFormat("#,##0");
    private static final DecimalFormat DF_3   = new DecimalFormat("#,##0.000");
    private static final DecimalFormat DF_2   = new DecimalFormat("#,##0.00");

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
        view.getBtnLoadExcel().addActionListener(e -> onLoadExcel());
        view.getBtnExit().addActionListener(e -> onExit());
        view.getBtnOpenWorkingCopy().addActionListener(e -> onOpenWorkingCopy());
        view.getControlsPanel().getBtnSimulate().addActionListener(e -> onSimulate());

        // ✅ nuovo listener: CE Budget 2022 (base fisso)
        view.getBtnShowCeBudget().addActionListener(e -> onShowCeBudgetBase());

        log.debug("Listener UI registrati.");
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

        int percent;
        try {
            percent = view.getControlsPanel().getPercent();
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view, e.getMessage(), "Input non valido", JOptionPane.WARNING_MESSAGE);
            return;
        }

        log.info("Simula articolo={} mode={} percent={}", ar, mode, percent);

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

                if (a != null && q != null && p != null && pe != null && cmpe != null) {
                    colCat = catColCandidate;
                    colArt = a;
                    colQty = q;
                    colPos = p;
                    colPeur = pe;
                    colCMPeur = cmpe;
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
            // 4B) CE dopo variazione (FIX): calcolato, non riletto da Excel
            // =========================================================
            Map<String, Double> ceAfterVar = computeCeAfterVar(ceBase, targetCat, fatt0, fatt1, cogs0, cogs1);

            // =========================================================
            // 5) Step 2: compensazione per mantenere POS(calc) costante
            // =========================================================
            double compValue;
            String compensatedVarLabel;

            ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, q1);
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, p1);
            eval.evaluateAll();

            if (mode == SimulationMode.QUANTITY) {
                compValue = cmp0 + (pos0Calc / q1);
                compensatedVarLabel = "P medio (€/kg)";
                ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, compValue);
            } else {
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
            double pos2Excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);

            double qStar = (mode == SimulationMode.QUANTITY) ? q1 : compValue;
            double pStar = (mode == SimulationMode.QUANTITY) ? compValue : p1;

            double fattStar = qStar * pStar;
            double cogsStar = qStar * cmp0;
            double posStarCalc = fattStar - cogsStar;

            // =========================================================
            // 6) Dettagli
            // =========================================================
            String whatChanged = (mode == SimulationMode.QUANTITY) ? "Quantità (kg)" : "Prezzo (€/kg)";
            String unitPrice = "€/kg";

            StringBuilder sb = new StringBuilder();
            sb.append("Articolo: ").append(targetArt).append(" [").append(targetCat).append("]\n");
            sb.append("Percentuale applicata: ").append(percent).append("%\n\n");

            sb.append("Valori originali (tabella destra):\n");
            sb.append("  Q0 = ").append(DF_INT.format(q0)).append(" kg\n");
            sb.append("  P0 = ").append(DF_3.format(p0)).append(" ").append(unitPrice).append("\n");
            sb.append("  CMP0 = ").append(DF_3.format(cmp0)).append(" ").append(unitPrice).append("\n");
            sb.append("  Fatturato0 = Q0*P0 = ").append(DF_INT.format(fatt0)).append("\n");
            sb.append("  COGS0 = Q0*CMP0 = ").append(DF_INT.format(cogs0)).append("\n");
            sb.append("  POS0 (target, Excel) = ").append(DF_INT.format(pos0Excel)).append("\n");
            sb.append("  POS0 (calc) = ").append(DF_INT.format(pos0Calc)).append("\n\n");

            sb.append("Step 1 — Dopo variazione (senza compensazione):\n");
            sb.append("  Leva modificata: ").append(whatChanged).append("\n");
            sb.append("  Q1 = ").append(DF_INT.format(q1)).append(" kg\n");
            sb.append("  P1 = ").append(DF_3.format(p1)).append(" ").append(unitPrice).append("\n");
            sb.append("  POS senza compensazione (Excel) = ").append(DF_INT.format(pos1Excel)).append("\n");
            sb.append("  POS senza compensazione (calc) = ").append(DF_INT.format(pos1Calc)).append("\n\n");

            sb.append("Step 2 — Compensazione per mantenere POS costante:\n");
            sb.append("  Target: POS(calc) = POS0(calc) = ").append(DF_INT.format(pos0Calc)).append("\n");
            sb.append("  Variabile compensata: ").append(compensatedVarLabel).append("\n");

            if (mode == SimulationMode.QUANTITY) {
                double deltaPct = ((compValue / p0) - 1.0) * 100.0;
                sb.append("  P* = ").append(DF_3.format(compValue)).append(" ").append(unitPrice).append("\n");
                sb.append("  ΔP% = ").append(DF_2.format(deltaPct)).append("%\n\n");
            } else {
                double deltaPct = ((compValue / q0) - 1.0) * 100.0;
                sb.append("  Q* = ").append(DF_INT.format(compValue)).append(" kg\n");
                sb.append("  ΔQ% = ").append(DF_2.format(deltaPct)).append("%\n\n");
            }

            sb.append("Check finale:\n");
            sb.append("  POS dopo compensazione (Excel) = ").append(DF_INT.format(pos2Excel)).append("\n");
            sb.append("  POS dopo compensazione (calc) = ").append(DF_INT.format(posStarCalc)).append("\n");
            sb.append("  Errore |POS(calc) - POS0(calc)| = ").append(DF_INT.format(Math.abs(posStarCalc - pos0Calc))).append("\n");

            view.getControlsPanel().setDetails(sb.toString());

            // =========================================================
            // 7) Grafici POS + Comp
            // =========================================================
            DefaultCategoryDataset posDS = new DefaultCategoryDataset();
            DefaultCategoryDataset compDS = new DefaultCategoryDataset();

            posDS.addValue(pos0Excel, "POS (Excel)", "Originale");
            posDS.addValue(pos1Excel, "POS (Excel)", "Dopo variazione");
            posDS.addValue(pos2Excel, "POS (Excel)", "Dopo compensazione");

            posDS.addValue(pos0Calc, "POS (calc)", "Originale");
            posDS.addValue(pos1Calc, "POS (calc)", "Dopo variazione");
            posDS.addValue(posStarCalc, "POS (calc)", "Dopo compensazione");

            if (mode == SimulationMode.QUANTITY) {
                compDS.addValue(q0, "Quantità (kg)", "Originale");
                compDS.addValue(q1, "Quantità (kg)", "Dopo variazione");
                compDS.addValue(q1, "Quantità (kg)", "Dopo compensazione");

                compDS.addValue(p0, "P medio (€/kg)", "Originale");
                compDS.addValue(p0, "P medio (€/kg)", "Dopo variazione");
                compDS.addValue(compValue, "P medio (€/kg)", "Dopo compensazione");
            } else {
                compDS.addValue(p0, "P medio (€/kg)", "Originale");
                compDS.addValue(p1, "P medio (€/kg)", "Dopo variazione");
                compDS.addValue(p1, "P medio (€/kg)", "Dopo compensazione");

                compDS.addValue(q0, "Quantità (kg)", "Originale");
                compDS.addValue(q0, "Quantità (kg)", "Dopo variazione");
                compDS.addValue(compValue, "Quantità (kg)", "Dopo compensazione");
            }

            JFreeChart posChart = ChartFactory.createLineChart(
                    "POS – " + targetArt,
                    "Scenario",
                    "POS",
                    posDS
            );

            JFreeChart compChart = ChartFactory.createLineChart(
                    "Compensazione – " + targetArt,
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
            // ✅ delta: formato più fine (non intero)
            configureCategoryChart(ceDeltaChart, false);
            view.getChartsPanel().setCeDeltaChart(ceDeltaChart);

            // =========================================================
            // 9) Salvo
            // =========================================================
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
        LineAndShapeRenderer r = (LineAndShapeRenderer) plot.getRenderer();

        r.setDefaultShapesVisible(true);
        r.setDefaultItemLabelsVisible(true);

        NumberFormat fmt = integerValues
                ? new DecimalFormat("#,##0")
                : new DecimalFormat("#,##0.000");

        r.setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator("{2}", fmt));

        ValueAxis range = plot.getRangeAxis();
        range.setUpperMargin(0.20);
        range.setLowerMargin(0.10);
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
            // ✅ come nel tuo controller “buono”
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
    // ✅ NUOVA FEATURE: finestra con grafico CE Budget 2022 base (fisso)
    // - legge dall'EXCEL ORIGINALE
    // - NON usa la working copy
    // =====================================================================
    private void onShowCeBudgetBase() {

        File original = model.getOriginalExcel();
        if (original == null || !original.exists()) {
            JOptionPane.showMessageDialog(view, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Workbook wb = WorkbookFactory.create(original)) {

            Sheet ce = findCeBudgetSheet(wb);
            if (ce == null) throw new IllegalStateException("Foglio CE Budget 2022 non trovato.");

            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
            DataFormatter fmt = new DataFormatter();

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
                ds.addValue(e.getValue(), "CE Budget 2022 (base)", e.getKey());
            }

            JFreeChart chart = ChartFactory.createLineChart(
                    "CE Budget 2022 – Base (Excel originale)",
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
            log.error("Errore apertura CE Budget base", ex);
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
