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
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.data.category.DefaultCategoryDataset;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.Desktop;
import java.io.File;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

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

            Sheet s = wb.getSheet("Ricavi");
            if (s == null) throw new IllegalStateException("Foglio 'Ricavi' non trovato.");

            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
            DataFormatter fmt = new DataFormatter();

            // =========================================================
            // 1) Trovo header tabella destra (Cat / Articolo / Quantità / POS)
            //    e le colonne P medio (€/kg) e CMP medio (€/kg)
            // =========================================================
            int headerRowIdx = -1;
            for (int r = 0; r <= Math.min(s.getLastRowNum(), 200); r++) {
                Row row = s.getRow(r);
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

            Row header = s.getRow(headerRowIdx);

            // helper ricerca nella finestra
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

            // candidati colonne "Cat" (può comparire più volte)
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

                // SOLO quelle che ci servono davvero per la tua logica
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
            // 2) Trovo la RIGA corretta (NON mi fido di ar.getRowIndex())
            // =========================================================
            String targetCat = (ar.getCat() == null) ? "" : ar.getCat().trim().toUpperCase();
            String targetArt = (ar.getArticolo() == null) ? "" : ar.getArticolo().trim().toUpperCase();

            int rowIdx = -1;
            for (int r = headerRowIdx + 1; r <= s.getLastRowNum(); r++) {
                Row rr = s.getRow(r);
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
                        "Non trovo la riga per Cat='" + targetCat + "' e Articolo='" + targetArt + "' nella tabella destra.\n" +
                        "Quindi ar.getRowIndex() non è affidabile e non c'è match testuale."
                );
            }

            // =========================================================
            // 3) Leggo valori base (EUR) + POS Excel
            // =========================================================
            eval.evaluateAll();

            double q0   = ricaviService.readNumeric(s, eval, rowIdx, colQty);
            double p0   = ricaviService.readNumeric(s, eval, rowIdx, colPeur);     // P medio (€/kg)
            double cmp0 = ricaviService.readNumeric(s, eval, rowIdx, colCMPeur);   // CMP medio (€/kg) -> già con dazio se presente
            double pos0Excel = ricaviService.readNumeric(s, eval, rowIdx, colPos);

            if (q0 <= 0)   throw new IllegalStateException("Q0 non valida letta da Excel: " + q0);
            if (p0 <= 0)   throw new IllegalStateException("P0 (€/kg) non valido letto da Excel: " + p0);
            if (cmp0 <= 0) throw new IllegalStateException("CMP0 (€/kg) non valido letto da Excel: " + cmp0);

            double fatt0 = q0 * p0;
            double cogs0 = q0 * cmp0;
            double pos0Calc = fatt0 - cogs0;

            // =========================================================
            // 4) Step 1: applico variazione (Q o P) e leggo POS Excel + calcolo POS(calc)
            // =========================================================
            // ripristino base su Excel (sicurezza)
            ricaviService.writeNumeric(s, rowIdx, colQty, q0);
            ricaviService.writeNumeric(s, rowIdx, colPeur, p0);
            eval.evaluateAll();

            double q1 = q0;
            double p1 = p0;

            if (mode == SimulationMode.QUANTITY) {
                q1 = q0 * (1.0 + percent / 100.0);
                ricaviService.writeNumeric(s, rowIdx, colQty, q1);
            } else { // PRICE
                p1 = p0 * (1.0 + percent / 100.0);
                ricaviService.writeNumeric(s, rowIdx, colPeur, p1);
            }

            eval.evaluateAll();
            double pos1Excel = ricaviService.readNumeric(s, eval, rowIdx, colPos);

            double fatt1 = q1 * p1;
            double cogs1 = q1 * cmp0;        // CMP rimane costante, cambia solo Q
            double pos1Calc = fatt1 - cogs1; // tua definizione

            // =========================================================
            // 5) Step 2: compensazione per mantenere POS(calc) = POS0(calc)
            //    (formula chiusa, niente bisection inutile)
            // =========================================================
            double compValue;            // valore compensato (P* oppure Q*)
            String compensatedVarLabel;  // testo

            // ripristino scenario step1 su Excel prima di compensare
            ricaviService.writeNumeric(s, rowIdx, colQty, q1);
            ricaviService.writeNumeric(s, rowIdx, colPeur, p1);
            eval.evaluateAll();

            if (mode == SimulationMode.QUANTITY) {
                // ho cambiato Q => compenso P
                // POS0 = Q1*(P* - CMP0)  => P* = CMP0 + POS0/Q1
                compValue = cmp0 + (pos0Calc / q1);
                compensatedVarLabel = "P medio (€/kg)";

                ricaviService.writeNumeric(s, rowIdx, colPeur, compValue);

            } else {
                // ho cambiato P => compenso Q
                // POS0 = Q*(P1 - CMP0) => Q* = POS0/(P1 - CMP0)
                double denom = (p1 - cmp0);
                if (Math.abs(denom) < 1e-12) {
                    throw new IllegalStateException("Compensazione impossibile: P1 - CMP0 = 0 (denominatore nullo).");
                }
                compValue = pos0Calc / denom;
                if (compValue <= 0) {
                    throw new IllegalStateException("Compensazione impossibile: Q* <= 0 (" + compValue + ").");
                }
                compensatedVarLabel = "Quantità (kg)";

                ricaviService.writeNumeric(s, rowIdx, colQty, compValue);
            }

            eval.evaluateAll();
            double pos2Excel = ricaviService.readNumeric(s, eval, rowIdx, colPos);

            double qStar = (mode == SimulationMode.QUANTITY) ? q1 : compValue;
            double pStar = (mode == SimulationMode.QUANTITY) ? compValue : p1;

            double fattStar = qStar * pStar;
            double cogsStar = qStar * cmp0;
            double posStarCalc = fattStar - cogsStar;

            // =========================================================
            // 6) Dettagli (chiari e coerenti)
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
            sb.append("  Fatturato1 = ").append(DF_INT.format(fatt1)).append("\n");
            sb.append("  COGS1 = ").append(DF_INT.format(cogs1)).append("\n");
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
            sb.append("  Fatturato* = ").append(DF_INT.format(fattStar)).append("\n");
            sb.append("  COGS* = ").append(DF_INT.format(cogsStar)).append("\n");
            sb.append("  POS dopo compensazione (Excel) = ").append(DF_INT.format(pos2Excel)).append("\n");
            sb.append("  POS dopo compensazione (calc) = ").append(DF_INT.format(posStarCalc)).append("\n");
            sb.append("  Errore |POS(calc) - POS0(calc)| = ").append(DF_INT.format(Math.abs(posStarCalc - pos0Calc))).append("\n");

            view.getControlsPanel().setDetails(sb.toString());

            // =========================================================
            // 7) Grafici: due serie (Excel vs calc) + compensazione (var modificata vs compensata)
            //    (se vuoi li rifiniamo dopo, intanto sono leggibili)
            // =========================================================
            DefaultCategoryDataset posDS = new DefaultCategoryDataset();
            DefaultCategoryDataset compDS = new DefaultCategoryDataset();

            // POS: Excel vs Calc
            posDS.addValue(pos0Excel, "POS (Excel)", "Originale");
            posDS.addValue(pos1Excel, "POS (Excel)", "Dopo variazione");
            posDS.addValue(pos2Excel, "POS (Excel)", "Dopo compensazione");

            posDS.addValue(pos0Calc, "POS (calc)", "Originale");
            posDS.addValue(pos1Calc, "POS (calc)", "Dopo variazione");
            posDS.addValue(posStarCalc, "POS (calc)", "Dopo compensazione");

            // Comp: variabile modificata + variabile compensata
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

            String compYLabel = (mode == SimulationMode.QUANTITY) ? "Valori (Q e P)" : "Valori (P e Q)";
            JFreeChart compChart = ChartFactory.createLineChart(
                    "Compensazione – " + targetArt,
                    "Scenario",
                    compYLabel,
                    compDS
            );

            // meglio: niente labels sovrapposte
            configureCategoryChart(posChart, true);
            configureCategoryChart(compChart, false);

            view.getChartsPanel().setPosChart(posChart);
            view.getChartsPanel().setCompChart(compChart);

            // salvo
            excelRepo.safeSaveWorkbook(wb);

        } catch (Exception ex) {
            log.error("Errore simulazione", ex);
            JOptionPane.showMessageDialog(view, "Errore simulazione: " + ex.getMessage(), "Errore",
                    JOptionPane.ERROR_MESSAGE);
        }
    }









    // ===========================
    // Chart config (compatibile)
    // ===========================

    private void configureCategoryChart(JFreeChart chart, boolean integerValues) {
        CategoryPlot plot = chart.getCategoryPlot();
        LineAndShapeRenderer r = (LineAndShapeRenderer) plot.getRenderer();

        // compatibile con JFreeChart 1.5.x
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

    private void animateScenarioCharts(DefaultCategoryDataset posDS,
                                       DefaultCategoryDataset compDS,
                                       SimulationMode mode,
                                       double q0, double p0,
                                       double q1, double p1,
                                       double comp,
                                       double pos0, double posNoFix) {

        posDS.clear();
        compDS.clear();

        // Originale
        posDS.addValue(pos0, "POS target (da Excel)", "Originale");
        posDS.addValue(pos0, "POS senza compensazione", "Originale");

        if (mode == SimulationMode.QUANTITY) {
            compDS.addValue(p0, "Prezzo originale (P0)", "Originale");
            compDS.addValue(p0, "Prezzo dopo variazione (P1)", "Originale"); // uguale
            compDS.addValue(p0, "Prezzo compensato (P*)", "Originale");
        } else {
            compDS.addValue(q0, "Quantità originale (Q0)", "Originale");
            compDS.addValue(q0, "Quantità dopo variazione (Q1)", "Originale"); // uguale
            compDS.addValue(q0, "Quantità compensata (Q*)", "Originale");
        }

        // Dopo variazione (comparsa)
        Timer t = new Timer(450, null);
        t.addActionListener(e -> {

            // ✅ qui ora c’è posNoFix, quindi il grafico POS mostra davvero la variazione
            posDS.addValue(pos0, "POS target (da Excel)", "Dopo variazione");
            posDS.addValue(posNoFix, "POS senza compensazione", "Dopo variazione");

            if (mode == SimulationMode.QUANTITY) {
                compDS.addValue(p0, "Prezzo originale (P0)", "Dopo variazione");
                compDS.addValue(p1, "Prezzo dopo variazione (P1)", "Dopo variazione");
                compDS.addValue(comp, "Prezzo compensato (P*)", "Dopo variazione");
            } else {
                compDS.addValue(q0, "Quantità originale (Q0)", "Dopo variazione");
                compDS.addValue(q1, "Quantità dopo variazione (Q1)", "Dopo variazione");
                compDS.addValue(comp, "Quantità compensata (Q*)", "Dopo variazione");
            }

            t.stop();
        });
        t.setRepeats(false);
        t.start();
    }

    // ===========================
    // Calcoli (Excel)
    // ===========================

    private double computePosNoFix(Workbook wb, ArticleRow ar, SimulationMode mode, double percent,
                                  double qty0, double price0, PriceColChoice priceChoice) {

        Sheet s = wb.getSheet("Ricavi");
        FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();

        if (mode == SimulationMode.QUANTITY) {
            double qty1 = qty0 * (1.0 + percent / 100.0);
            ricaviService.writeNumeric(s, ar.getRowIndex(), ar.getColQty(), qty1);
        } else {
            double p1 = price0 * (1.0 + percent / 100.0);
            ricaviService.writeNumeric(s, ar.getRowIndex(), priceChoice.colIndex, p1);
        }

        eval.evaluateAll();
        return ricaviService.readNumeric(s, eval, ar.getRowIndex(), ar.getColPos());
    }

    private double computePosAfterFix(Workbook wb, ArticleRow ar, SimulationMode mode, double percent,
                                     double qty0, double price0, PriceColChoice priceChoice, double comp) {

        Sheet s = wb.getSheet("Ricavi");
        FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();

        if (mode == SimulationMode.QUANTITY) {
            double qty1 = qty0 * (1.0 + percent / 100.0);
            ricaviService.writeNumeric(s, ar.getRowIndex(), ar.getColQty(), qty1);
            ricaviService.writeNumeric(s, ar.getRowIndex(), priceChoice.colIndex, comp);
        } else {
            double p1 = price0 * (1.0 + percent / 100.0);
            ricaviService.writeNumeric(s, ar.getRowIndex(), priceChoice.colIndex, p1);
            ricaviService.writeNumeric(s, ar.getRowIndex(), ar.getColQty(), comp);
        }

        eval.evaluateAll();
        return ricaviService.readNumeric(s, eval, ar.getRowIndex(), ar.getColPos());
    }

    private CompensationResult computeCompensationRobust(Workbook wb, ArticleRow ar, SimulationMode mode, double percent,
                                                        double posTarget, double qty0, double price0, PriceColChoice priceChoice) {

        Sheet s = wb.getSheet("Ricavi");
        FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();

        // applico prima la leva
        if (mode == SimulationMode.QUANTITY) {
            double qty1 = qty0 * (1.0 + percent / 100.0);
            ricaviService.writeNumeric(s, ar.getRowIndex(), ar.getColQty(), qty1);
        } else {
            double p1 = price0 * (1.0 + percent / 100.0);
            ricaviService.writeNumeric(s, ar.getRowIndex(), priceChoice.colIndex, p1);
        }
        eval.evaluateAll();

        boolean compensatePrice = (mode == SimulationMode.QUANTITY);
        double base = compensatePrice ? price0 : qty0;
        if (base <= 0) base = 1.0;

        double lo = Math.max(0.000001, base * 0.05);
        double hi = base * 5.0;

        java.util.function.DoubleFunction<Double> f = (x) -> {
            if (compensatePrice) {
                ricaviService.writeNumeric(s, ar.getRowIndex(), priceChoice.colIndex, x);
            } else {
                ricaviService.writeNumeric(s, ar.getRowIndex(), ar.getColQty(), x);
            }
            eval.evaluateAll();
            double pos = ricaviService.readNumeric(s, eval, ar.getRowIndex(), ar.getColPos());
            return pos - posTarget;
        };

        double flo = f.apply(lo);
        double fhi = f.apply(hi);

        int expandMax = 30;
        int expand = 0;
        while (flo * fhi > 0 && expand < expandMax) {
            lo = Math.max(0.000001, lo / 2.0);
            hi = hi * 2.0;
            flo = f.apply(lo);
            fhi = f.apply(hi);
            expand++;
        }

        if (flo * fhi > 0) {
            throw new IllegalStateException(
                    "Compensazione non raggiungibile: non trovo intervallo [lo,hi] che riporti POS al target.\n" +
                    "Questo accade se la variabile compensata non impatta davvero il POS o se il target non è raggiungibile."
            );
        }

        double mid = 0.0;
        for (int i = 0; i < 60; i++) {
            mid = (lo + hi) / 2.0;
            double fmid = f.apply(mid);
            if (flo * fmid <= 0) {
                hi = mid;
                fhi = fmid;
            } else {
                lo = mid;
                flo = fmid;
            }
        }

        return new CompensationResult(lo, hi, mid);
    }

    private void restoreBase(Sheet s, FormulaEvaluator eval, ArticleRow ar, double q0, double p0, PriceColChoice priceChoice) {
        ricaviService.writeNumeric(s, ar.getRowIndex(), ar.getColQty(), q0);
        ricaviService.writeNumeric(s, ar.getRowIndex(), priceChoice.colIndex, p0);
        eval.evaluateAll();
    }

    // ===========================
    // Scelta colonna prezzo (EUR vs USD) che cambia davvero il POS
    // ===========================

    private PriceColChoice detectPriceColumnThatAffectsPos(Workbook wb, Sheet s, FormulaEvaluator eval,
            ArticleRow ar, double q0, double pos0) {

// candidati: (colonna, label)
List<int[]> cols = new ArrayList<>();
List<String> labels = new ArrayList<>();

if (ar.getColPmedioEUR() >= 0) { cols.add(new int[]{ar.getColPmedioEUR()}); labels.add("P medio (€/kg)"); }
if (ar.getColPmedioUSD() != null) { cols.add(new int[]{ar.getColPmedioUSD()}); labels.add("P medio ($/kg)"); }
if (ar.getColCMPmedioEUR() >= 0) { cols.add(new int[]{ar.getColCMPmedioEUR()}); labels.add("CMP medio (€/kg)"); }
if (ar.getColCMPmedioUSD() != null) { cols.add(new int[]{ar.getColCMPmedioUSD()}); labels.add("CMP medio ($/kg)"); }

if (cols.isEmpty()) {
throw new IllegalStateException("Nessuna colonna prezzo/CMP trovata per questa riga.");
}

double bestDelta = -1.0;
int bestCol = -1;
String bestLabel = null;

for (int i = 0; i < cols.size(); i++) {
int c = cols.get(i)[0];
String lab = labels.get(i);

double p0 = ricaviService.readNumeric(s, eval, ar.getRowIndex(), c);
double delta = testPosDeltaByPerturbingPrice(s, eval, ar, c, p0, pos0);

if (delta > bestDelta) {
bestDelta = delta;
bestCol = c;
bestLabel = lab;
}
}

// restore quantità (paranoia)
ricaviService.writeNumeric(s, ar.getRowIndex(), ar.getColQty(), q0);
eval.evaluateAll();

// se bestDelta ~ 0: NESSUNA colonna impatta davvero il POS
if (bestDelta < 1e-9) {
throw new IllegalStateException(
"Impossibile individuare una colonna prezzo/CMP che impatti il POS per " + ar.getArticolo() + ".\n" +
"Probabile: il POS dipende da altre celle o la riga non è quella corretta nella tabella."
);
}

return new PriceColChoice(bestCol, bestLabel);
}


    private double testPosDeltaByPerturbingPrice(Sheet s, FormulaEvaluator eval, ArticleRow ar,
                                                 int priceCol, double p0, double pos0) {

        if (p0 <= 0) return 0.0;

        // salva originale
        double original = p0;

        // perturbo +1%
        double p1 = original * 1.01;
        ricaviService.writeNumeric(s, ar.getRowIndex(), priceCol, p1);
        eval.evaluateAll();
        double pos1 = ricaviService.readNumeric(s, eval, ar.getRowIndex(), ar.getColPos());

        // restore
        ricaviService.writeNumeric(s, ar.getRowIndex(), priceCol, original);
        eval.evaluateAll();

        return Math.abs(pos1 - pos0);
    }

    private void onExit() {
        excelRepo.cleanup();
        System.exit(0);
    }

    // ===========================
    // DTO
    // ===========================

    private static class CompensationResult {
        final double lo;
        final double hi;
        final double compValue;
        CompensationResult(double lo, double hi, double compValue) {
            this.lo = lo;
            this.hi = hi;
            this.compValue = compValue;
        }
    }

    private static class PriceColChoice {
        final int colIndex;
        final String label;   // "P medio (€/kg)" ecc
        PriceColChoice(int colIndex, String label) {
            this.colIndex = colIndex;
            this.label = label;
            if (this.colIndex < 0) throw new IllegalStateException("Colonna prezzo scelta non valida.");
        }
    }

}
