package controller;

import model.*;
import repository.ExcelRepository;
import service.RicaviExcelService;
import view.MainFrame;
import view.PremioCompFrame;

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
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

public class PremioCompController {

    private static final Logger log = LogManager.getLogger(PremioCompController.class);

    private final AppModel model;
    private final MainFrame mainView;
    private final ExcelRepository excelRepo;

    private final PremioCompFrame premioView;

    private RicaviExcelService ricaviService;
    private List<ArticleRow> cachedArticles = new ArrayList<>();

    private static final DecimalFormat DF_INT = new DecimalFormat("#,##0");
    private static final DecimalFormat DF_3   = new DecimalFormat("#,##0.000");
    private static final DecimalFormat DF_2   = new DecimalFormat("#,##0.00");

    // ===========================
    // RIFERIMENTI FOGLIO "Ricavi"
    // Premio mensile: Q66
    // Mensilità:      P66
    // Premio annuo:   W66 = Q66 * P66
    // Premio in somma POS: X66 (tipicamente -W66 o +W66)
    // POS totale:     X67
    // ===========================
    private static final int ROW_66 = 65; // riga 66 -> index 65
    private static final int ROW_67 = 66; // riga 67 -> index 66

    private static final int COL_P  = 15; // P -> index 15
    private static final int COL_Q  = 16; // Q -> index 16
    private static final int COL_W  = 22; // W -> index 22
    private static final int COL_X  = 23; // X -> index 23  (X66, X67)

    public PremioCompController(AppModel model, MainFrame mainView, ExcelRepository excelRepo) {
        this.model = model;
        this.mainView = mainView;
        this.excelRepo = excelRepo;
        this.premioView = new PremioCompFrame();
        initListeners();
    }

    private void initListeners() {
    	premioView.getBtnReset().addActionListener(e -> onResetExcelPremio());

        premioView.getControlsPanel().getBtnSimulate().addActionListener(e -> onSimulatePremio());

        premioView.getBtnBack().addActionListener(e -> {
            premioView.dispose();
            mainView.setVisible(true);
        });

        premioView.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosed(WindowEvent e) {
                mainView.setVisible(true);
            }
        });
    }
    
    private void onResetExcelPremio() {
        if (model.getWorkingExcelCopy() == null) {
            JOptionPane.showMessageDialog(premioView, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
            return;
        }

        int ok = JOptionPane.showConfirmDialog(
                premioView,
                "Vuoi ripristinare la copia di lavoro allo stato iniziale?\n" +
                "Perderai tutte le simulazioni fatte finora nella sessione.",
                "Reset Excel",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.WARNING_MESSAGE
        );
        if (ok != JOptionPane.YES_OPTION) return;

        try {
            File wc = excelRepo.resetWorkingCopyToBase();
            model.setWorkingExcelCopy(wc);

            // ✅ ricreo service e ricarico articoli
            ricaviService = new RicaviExcelService(wc);
            cachedArticles = ricaviService.loadArticles();
            premioView.getControlsPanel().setArticles(cachedArticles);

            // ✅ pulisco dettagli e grafici
            premioView.getControlsPanel().setDetails("");
            premioView.setPosChart(null);
            premioView.setPremioChart(null);

            JOptionPane.showMessageDialog(premioView, "Reset completato. Riparti dai valori base.", "OK", JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception ex) {
            log.error("[PREMIO] Errore reset working copy", ex);
            JOptionPane.showMessageDialog(
                    premioView,
                    "Errore reset: " + ex.getMessage() + "\n\n" +
                    "Se hai aperto la copia in Excel, chiudila e riprova.",
                    "Errore",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }


    public void open() {
        if (model.getWorkingExcelCopy() == null) {
            JOptionPane.showMessageDialog(mainView, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try {
            if (ricaviService == null) {
                ricaviService = new RicaviExcelService(model.getWorkingExcelCopy());
                cachedArticles = ricaviService.loadArticles();
                premioView.getControlsPanel().setArticles(cachedArticles);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(mainView, "Errore inizializzazione: " + ex.getMessage(), "Errore", JOptionPane.ERROR_MESSAGE);
            return;
        }

        mainView.setVisible(false);
        premioView.setVisible(true);
    }

    private void onSimulatePremio() {

        if (model.getWorkingExcelCopy() == null || ricaviService == null) {
            JOptionPane.showMessageDialog(premioView, "Carica prima un file Excel.", "Attenzione", JOptionPane.WARNING_MESSAGE);
            return;
        }

        ArticleRow ar = premioView.getControlsPanel().getSelectedArticle();
        if (ar == null) return;

        SimulationMode mode = premioView.getControlsPanel().getMode();

        // Percentuale robusta: accetta "3,5" e "3.5"
        double percent;
        try {
            String raw = String.valueOf(premioView.getControlsPanel().getPercent()).trim();
            raw = raw.replace(",", ".");
            percent = Double.parseDouble(raw);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(premioView, "Percentuale non valida.", "Input non valido", JOptionPane.WARNING_MESSAGE);
            return;
        }

        String targetCat = (ar.getCat() == null) ? "" : ar.getCat().trim().toUpperCase();
        String targetArt = (ar.getArticolo() == null) ? "" : ar.getArticolo().trim().toUpperCase();

        log.info("[PREMIO] Simula articolo={} cat={} mode={} percent={}", targetArt, targetCat, mode, percent);

        try (Workbook wb = WorkbookFactory.create(model.getWorkingExcelCopy())) {

            Sheet ricaviSheet = wb.getSheet("Ricavi");
            if (ricaviSheet == null) throw new IllegalStateException("Foglio 'Ricavi' non trovato.");

            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
            DataFormatter fmt = new DataFormatter();

            // =========================================================
            // 1) trova header tabella destra
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
                if (hasCat && hasArt && hasQty && hasPos) { headerRowIdx = r; break; }
            }
            if (headerRowIdx < 0) throw new IllegalStateException("Header tabella destra non trovato (Cat/Articolo/Quantità/POS).");

            Row header = ricaviSheet.getRow(headerRowIdx);

            java.util.function.BiFunction<String, int[], Integer> findExactInWindow = (exact, win) -> {
                for (int c = win[0]; c <= win[1]; c++) {
                    String v = fmt.formatCellValue(header.getCell(c)).trim();
                    if (v.equalsIgnoreCase(exact)) return c;
                }
                return null;
            };

            java.util.function.BiFunction<String, int[], Integer> findLastExactInWindow = (exact, win) -> {
                Integer found = null;
                for (int c = win[0]; c <= win[1]; c++) {
                    String v = fmt.formatCellValue(header.getCell(c)).trim();
                    if (v.equalsIgnoreCase(exact)) found = c;
                }
                return found;
            };

            java.util.function.BiFunction<String, int[], Integer> findContainsInWindow = (needleLower, win) -> {
                String needle = needleLower.toLowerCase();
                for (int c = win[0]; c <= win[1]; c++) {
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
            if (!foundTable) throw new IllegalStateException("Impossibile identificare colonne (Cat/Articolo/Quantità/P medio/CMP medio/POS).");

            // =========================================================
            // 2) trova riga articolo
            // =========================================================
            int rowIdx = -1;
            for (int r = headerRowIdx + 1; r <= ricaviSheet.getLastRowNum(); r++) {
                Row rr = ricaviSheet.getRow(r);
                if (rr == null) continue;

                String catTxt = fmt.formatCellValue(rr.getCell(colCat)).trim().toUpperCase();
                String artTxt = fmt.formatCellValue(rr.getCell(colArt)).trim().toUpperCase();

                if (catTxt.equals(targetCat) && artTxt.equals(targetArt)) { rowIdx = r; break; }
            }
            if (rowIdx < 0) throw new IllegalStateException("Riga non trovata per Cat='" + targetCat + "' Articolo='" + targetArt + "'.");

            // =========================================================
            // 3) Letture base (riga + premio + POS totale)
            // =========================================================
            eval.evaluateAll();

            double q0   = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colQty);
            double p0   = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPeur);
            double cmp0 = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colCMPeur);

            // POS riga: uso Excel per baseline, ma controllo anche formula robusta
            double posRow0_excel = ricaviService.readNumeric(ricaviSheet, eval, rowIdx, colPos);
            double posRow0_calc  = (p0 - cmp0) * q0;

            // Se Excel non è coerente (o non ricalcola), baseline prendo quello Excel se è “sensato”, altrimenti il calcolato
            double posRow0 = (Math.abs(posRow0_excel) > 1e-9) ? posRow0_excel : posRow0_calc;

            double months = readNumericCell(ricaviSheet, eval, ROW_66, COL_P);
            if (months <= 0) throw new IllegalStateException("Mensilità P66 non valida: " + months);

            double premioMens0 = readNumericCell(ricaviSheet, eval, ROW_66, COL_Q);
            double premioAnn0  = readNumericCell(ricaviSheet, eval, ROW_66, COL_W);
            double x66_0        = readNumericCell(ricaviSheet, eval, ROW_66, COL_X);

            // POS totale baseline (può non ricalcolare dopo, ma baseline la prendiamo dal file)
            double totPos0 = readNumericCell(ricaviSheet, eval, ROW_67, COL_X);

            if (q0 <= 0) throw new IllegalStateException("Q0 non valida: " + q0);
            if (p0 <= 0) throw new IllegalStateException("P0 non valido: " + p0);
            if (cmp0 <= 0) throw new IllegalStateException("CMP0 non valido: " + cmp0);

            // segno di X66 rispetto a W66
            int signX66 = detectSignForX66(x66_0, premioAnn0);

            // =========================================================
            // 4) Step 1: applico variazione (Q o P) - PREMIO invariato
            // =========================================================
            // reset riga
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colQty, q0);
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colPeur, p0);

            // riallineo premio base Q66/W66/X66
            setPremioMensile(ricaviSheet, eval, premioMens0, months, signX66);
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

            // ====== QUI È IL FIX: POS riga e POS totale calcolati (NO dipendenza da ricalcolo Excel) ======
            double posRow1_calc = (p1 - cmp0) * q1;

            // opzionale: scrivo POS della riga anche in Excel così "si vede" cambiare
            ricaviService.writeNumeric(ricaviSheet, rowIdx, colPos, posRow1_calc);

            // POS totale dopo variazione: baseline + delta del POS riga selezionata
            double totPos1_calc = totPos0 + (posRow1_calc - posRow0);

            // =========================================================
            // 5) Step 2: compenso PREMIO per riportare POS totale a totPos0
            //    Compenso agendo su X66 (che è ±W66) in modo deterministico
            // =========================================================
            double deltaPosRow = posRow1_calc - posRow0;

            // voglio annullare deltaPosRow, quindi modifico X66 di -deltaPosRow
            double x66_star = x66_0 - deltaPosRow;

            double denom = signX66 * months;
            if (Math.abs(denom) < 1e-9) {
                throw new IllegalStateException("Compensazione impossibile: months o segno non valido.");
            }

            double premioMensStar = x66_star / denom;
            if (premioMensStar < 0) {
                throw new IllegalStateException("Compensazione impossibile: Premio mensile* < 0 (" + premioMensStar + ").");
            }

            // scrivo premio compensato Q66/W66/X66
            setPremioMensile(ricaviSheet, eval, premioMensStar, months, signX66);
            eval.evaluateAll();

            double premioAnnStar = readNumericCell(ricaviSheet, eval, ROW_66, COL_W);
            double x66_star1     = readNumericCell(ricaviSheet, eval, ROW_66, COL_X);

            // POS totale dopo compensazione: per costruzione torna a totPos0
            double totPos2_calc = totPos1_calc + (x66_star1 - x66_0);

            // =========================================================
            // 6) Dettagli
            // =========================================================
            String whatChanged = (mode == SimulationMode.QUANTITY) ? "Quantità (kg)" : "Prezzo (€/kg)";

            StringBuilder sb = new StringBuilder();
            sb.append("Articolo: ").append(targetArt).append(" [").append(targetCat).append("]\n");
            sb.append("Percentuale applicata: ").append(DF_2.format(percent)).append("%\n\n");

            sb.append("Valori originali:\n");
            sb.append("  Q0 = ").append(DF_INT.format(q0)).append(" kg\n");
            sb.append("  P0 = ").append(DF_3.format(p0)).append(" €/kg\n");
            sb.append("  CMP0 = ").append(DF_3.format(cmp0)).append(" €/kg\n");
            sb.append("  Premio mensile (Q66) = ").append(DF_INT.format(premioMens0)).append("\n");
            sb.append("  Mensilità (P66) = ").append(DF_INT.format(months)).append("\n");
            sb.append("  Premio annuo (W66) = ").append(DF_INT.format(premioAnn0)).append("\n");
            sb.append("  Premio in somma (X66) = ").append(DF_INT.format(x66_0)).append("\n");
            sb.append("  POS riga (baseline) = ").append(DF_INT.format(posRow0)).append("\n");
            sb.append("  POS TOTALE (baseline) = ").append(DF_INT.format(totPos0)).append("\n\n");

            sb.append("Step 1 — Dopo variazione (premio invariato):\n");
            sb.append("  Leva modificata: ").append(whatChanged).append("\n");
            sb.append("  Q1 = ").append(DF_INT.format(q1)).append(" kg\n");
            sb.append("  P1 = ").append(DF_3.format(p1)).append(" €/kg\n");
            sb.append("  Premio mensile invariato (Q66) = ").append(DF_INT.format(premioMens0)).append("\n");
            sb.append("  POS riga (calcolato) = ").append(DF_INT.format(posRow1_calc)).append("\n");
            sb.append("  POS TOTALE (calcolato) = ").append(DF_INT.format(totPos1_calc)).append("\n\n");

            sb.append("Step 2 — Compensazione PREMIO per mantenere POS TOTALE costante:\n");
            sb.append("  Target: POS TOTALE = ").append(DF_INT.format(totPos0)).append("\n");
            sb.append("  Premio mensile* (Q66) = ").append(DF_INT.format(premioMensStar)).append("\n");
            sb.append("  Premio annuo* (W66) = ").append(DF_INT.format(premioAnnStar)).append("\n");
            sb.append("  Premio in somma* (X66) = ").append(DF_INT.format(x66_star1)).append("\n");
            double deltaPremioPct = (premioMens0 == 0) ? 0 : ((premioMensStar / premioMens0) - 1.0) * 100.0;
            sb.append("  ΔPremio% = ").append(DF_2.format(deltaPremioPct)).append("%\n\n");

            sb.append("Check finale:\n");
            sb.append("  POS TOTALE dopo compensazione (calcolato) = ").append(DF_INT.format(totPos2_calc)).append("\n");
            sb.append("  Errore |totPos2 - totPos0| = ").append(DF_INT.format(Math.abs(totPos2_calc - totPos0))).append("\n");

            premioView.getControlsPanel().setDetails(sb.toString());

            // =========================================================
            // 7) Grafico POS totale (usa valori calcolati)
            // =========================================================
            DefaultCategoryDataset posTotDS = new DefaultCategoryDataset();
            posTotDS.addValue(totPos0,      "POS Totale", "Originale");
            posTotDS.addValue(totPos1_calc, "POS Totale", "Dopo variazione");
            posTotDS.addValue(totPos2_calc, "POS Totale", "Dopo compensazione");

            JFreeChart posChart = ChartFactory.createLineChart(
                    "POS Totale – " + targetArt + " (compenso con PREMIO)",
                    "Scenario",
                    "POS Totale",
                    posTotDS
            );
            configureCategoryChart(posChart, true);
            premioView.setPosChart(posChart);

            // =========================================================
            // 8) Grafico Leva + Premio
            // =========================================================
            DefaultCategoryDataset premioDS = new DefaultCategoryDataset();

            if (mode == SimulationMode.QUANTITY) {
                premioDS.addValue(q0, "Quantità (kg)", "Originale");
                premioDS.addValue(q1, "Quantità (kg)", "Dopo variazione");
                premioDS.addValue(q1, "Quantità (kg)", "Dopo compensazione");
            } else {
                premioDS.addValue(p0, "P medio (€/kg)", "Originale");
                premioDS.addValue(p1, "P medio (€/kg)", "Dopo variazione");
                premioDS.addValue(p1, "P medio (€/kg)", "Dopo compensazione");
            }

            premioDS.addValue(premioMens0,    "Premio mensile (Q66)", "Originale");
            premioDS.addValue(premioMens0,    "Premio mensile (Q66)", "Dopo variazione");
            premioDS.addValue(premioMensStar, "Premio mensile (Q66)", "Dopo compensazione");

            premioDS.addValue(premioAnn0,     "Premio annuo (W66)", "Originale");
            premioDS.addValue(premioAnn0,     "Premio annuo (W66)", "Dopo variazione");
            premioDS.addValue(premioAnnStar,  "Premio annuo (W66)", "Dopo compensazione");

            JFreeChart premioChart = ChartFactory.createLineChart(
                    "Leva + Premio – " + targetArt,
                    "Scenario",
                    "Valore",
                    premioDS
            );
            configureCategoryChart(premioChart, false);
            premioView.setPremioChart(premioChart);

            // =========================================================
            // 9) Salva
            // =========================================================
            excelRepo.safeSaveWorkbook(wb);

        } catch (Exception ex) {
            log.error("[PREMIO] Errore simulazione", ex);
            JOptionPane.showMessageDialog(premioView, "Errore simulazione: " + ex.getMessage(), "Errore", JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Imposta Q66 (premio mensile) e aggiorna coerentemente:
     * - W66 = Q66 * P66 (forzato)
     * - X66 = ±W66 (forzato)
     */
    private void setPremioMensile(Sheet sh, FormulaEvaluator eval, double premioMensile, double months, int signX66) {
        // Q66
        writeNumericCell(sh, ROW_66, COL_Q, premioMensile);

        // W66 = Q66 * P66 (forzato)
        double premioAnnuo = premioMensile * months;
        writeNumericCell(sh, ROW_66, COL_W, premioAnnuo);

        // X66 = ±W66 (forzato)
        writeNumericCell(sh, ROW_66, COL_X, signX66 * premioAnnuo);

        if (eval != null) eval.evaluateAll();
    }

    /**
     * Se X66 è circa -W66 => sign = -1
     * Se X66 è circa +W66 => sign = +1
     * (default -1 perché tipicamente il premio è un costo)
     */
    private int detectSignForX66(double x66, double w66) {
        if (w66 == 0) return -1;

        double dPlus  = Math.abs(x66 - w66);
        double dMinus = Math.abs(x66 + w66);

        if (dMinus < dPlus) return -1;
        if (dPlus < dMinus) return +1;

        return -1;
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

    private void writeNumericCell(Sheet sh, int rowIdx, int colIdx, double value) {
        Row row = sh.getRow(rowIdx);
        if (row == null) row = sh.createRow(rowIdx);

        Cell c = row.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        c.setCellType(CellType.NUMERIC);
        c.setCellValue(value);
    }

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
}
