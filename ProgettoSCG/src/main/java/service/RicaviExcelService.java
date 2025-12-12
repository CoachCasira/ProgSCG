package service;

import java.io.File;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import model.ArticleRow;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;


public class RicaviExcelService {
	
	private static class TableCols {
	    final int colCat, colArticolo, colCategoria, colQty;
	    final int colPmedioEUR, colPmedioUSD;
	    final int colCMPmedioEUR, colCMPmedioUSD;
	    final int colPos;

	    TableCols(int colCat, int colArticolo, int colCategoria, int colQty,
	              int colPmedioEUR, int colPmedioUSD,
	              int colCMPmedioEUR, int colCMPmedioUSD,
	              int colPos) {
	        this.colCat = colCat;
	        this.colArticolo = colArticolo;
	        this.colCategoria = colCategoria;
	        this.colQty = colQty;
	        this.colPmedioEUR = colPmedioEUR;
	        this.colPmedioUSD = colPmedioUSD;
	        this.colCMPmedioEUR = colCMPmedioEUR;
	        this.colCMPmedioUSD = colCMPmedioUSD;
	        this.colPos = colPos;
	    }
	}


    private static final Logger log = LogManager.getLogger(RicaviExcelService.class);

    private final File workingCopy;

    // Pattern robusti: accettano "PF P1" / "PFP1" / "PFV 12" ecc.
    private static final Pattern MP_PATTERN  = Pattern.compile("^MP\\s*(\\d{1,2})$", Pattern.CASE_INSENSITIVE);
    private static final Pattern PFP_PATTERN = Pattern.compile("^(?:PF\\s*P\\s*|PFP\\s*)(\\d{1,2})$", Pattern.CASE_INSENSITIVE);
    private static final Pattern PFV_PATTERN = Pattern.compile("^(?:PF\\s*V\\s*|PFV\\s*)(\\d{1,2})$", Pattern.CASE_INSENSITIVE);
    private static final Pattern PFA_PATTERN = Pattern.compile("^(?:PF\\s*A\\s*|PFA\\s*)(\\d{1,2})$", Pattern.CASE_INSENSITIVE);

    public RicaviExcelService(File workingCopy) {
        this.workingCopy = workingCopy;
    }

    public List<ArticleRow> loadArticles() throws Exception {
        try (Workbook wb = WorkbookFactory.create(workingCopy)) {

            Sheet sheet = wb.getSheet("Ricavi");
            if (sheet == null) throw new IllegalStateException("Foglio 'Ricavi' non trovato.");

            DataFormatter fmt = new DataFormatter();

            int headerRowIdx = findHeaderRow(sheet, fmt);
            if (headerRowIdx < 0) {
                throw new IllegalStateException("Header tabella destra non trovato (Cat/Articolo/Quantità/P medio/POS).");
            }

            Row header = sheet.getRow(headerRowIdx);
            TableCols cols = detectRightTableColumns(header, fmt);

            log.info("Tabella DESTRA: riga {} Cat={} Articolo={} Categoria={} Qty={} P€={} P$={} POS={}",
                    headerRowIdx + 1,
                    cols.colCat, cols.colArticolo, cols.colCategoria, cols.colQty,
                    cols.colPmedioEUR, cols.colPmedioUSD, cols.colPos);

            List<ArticleRow> out = new ArrayList<>();

            for (int r = headerRowIdx + 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String articoloRaw = fmt.formatCellValue(row.getCell(cols.colArticolo)).trim();
                if (articoloRaw.isEmpty()) continue;

                String articoloNorm = normalizeSpaces(articoloRaw);
                String up = articoloNorm.toUpperCase();

                // Escludo PCL sempre
                if (up.startsWith("PCL")) continue;

                // Classifico SOLO in base all'articolo (robusto)
                String exposedCat = classifyExposedCat(articoloNorm);
                if (exposedCat == null) continue;

                String categoria = fmt.formatCellValue(row.getCell(cols.colCategoria)).trim();

                out.add(new ArticleRow(
                        exposedCat,
                        articoloNorm,
                        categoria,
                        r,
                        cols.colQty,
                        cols.colPmedioEUR,
                        cols.colPmedioUSD,
                        cols.colCMPmedioEUR,
                        cols.colCMPmedioUSD,
                        cols.colPos
                ));

            }

            // Ordine naturale: MP -> PFP -> PFV -> PFA, poi numerico
            out.sort(Comparator
                    .comparing(ArticleRow::getCat)
                    .thenComparing(a -> naturalKey(a.getArticolo()))
            );

            log.info("Articoli caricati (filtrati): {}", out.size());
            return out;
        }
    }

    // ===========================
    // Lettura/scrittura numerica
    // ===========================

    public double readNumeric(Sheet sheet, FormulaEvaluator eval, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return 0.0;

        Cell c = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (c == null) return 0.0;

        if (c.getCellType() == CellType.FORMULA) {
            CellValue cv = eval.evaluate(c);
            if (cv != null && cv.getCellType() == CellType.NUMERIC) return cv.getNumberValue();
            return 0.0;
        }

        if (c.getCellType() == CellType.NUMERIC) return c.getNumericCellValue();

        if (c.getCellType() == CellType.STRING) {
            try { return Double.parseDouble(c.getStringCellValue().trim().replace(",", ".")); }
            catch (Exception ignore) { return 0.0; }
        }

        return 0.0;
    }

    public void writeNumeric(Sheet sheet, int rowIdx, int colIdx, double value) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) row = sheet.createRow(rowIdx);

        Cell c = row.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        c.setCellValue(value);
    }

    // ===========================
    // Header / Tabella destra
    // ===========================

    private int findHeaderRow(Sheet sheet, DataFormatter fmt) {
        for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 140); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            boolean hasCat = false, hasArt = false, hasQty = false, hasPos = false;

            for (int c = 0; c < Math.min(row.getLastCellNum(), 120); c++) {
                String v = fmt.formatCellValue(row.getCell(c)).trim();
                if (v.equalsIgnoreCase("Cat")) hasCat = true;
                if (v.equalsIgnoreCase("Articolo")) hasArt = true;
                if (v.toLowerCase().contains("quantità")) hasQty = true;
                if (v.equalsIgnoreCase("POS")) hasPos = true;
            }

            if (hasCat && hasArt && hasQty && hasPos) return r;
        }
        return -1;
    }

    private TableCols detectRightTableColumns(Row header, DataFormatter fmt) {

        List<Integer> catCols = new ArrayList<>();
        for (int c = 0; c < header.getLastCellNum(); c++) {
            String v = fmt.formatCellValue(header.getCell(c)).trim();
            if (v.equalsIgnoreCase("Cat")) catCols.add(c);
        }
        if (catCols.isEmpty()) throw new IllegalStateException("Header: colonna 'Cat' non trovata.");

        for (int catCol : catCols) {
            int start = catCol;
            int end = Math.min(header.getLastCellNum() - 1, catCol + 40);

            Integer colArt = findExactInWindow(header, fmt, start, end, "Articolo");
            Integer colCategoria = findExactInWindow(header, fmt, start, end, "Categoria");
            Integer colQty = findContainsInWindow(header, fmt, start, end, "quantità");

            // P medio
            Integer colPeur = findContainsInWindow(header, fmt, start, end, "p medio (€/kg)");
            Integer colPusd = findContainsInWindow(header, fmt, start, end, "p medio ($/kg)");

            // CMP medio
            Integer colCe = findContainsInWindow(header, fmt, start, end, "cmp medio (€/kg)");
            Integer colCu = findContainsInWindow(header, fmt, start, end, "cmp medio ($/kg)");

            Integer colPos = findLastExactInWindow(header, fmt, start, end, "POS");

            if (colArt != null && colCategoria != null && colQty != null && colPos != null) {
                int pE = (colPeur != null) ? colPeur : -1;
                int pU = (colPusd != null) ? colPusd : -1;
                int cE = (colCe != null) ? colCe : -1;
                int cU = (colCu != null) ? colCu : -1;

                // Basta che almeno UNA tra P medio o CMP medio esista (nel tuo file ci sono entrambe)
                if (pE >= 0 || cE >= 0) {
                    return new TableCols(catCol, colArt, colCategoria, colQty, pE, pU, cE, cU, colPos);
                }
            }
        }

        throw new IllegalStateException("Impossibile identificare colonne tabella destra (Cat/Articolo/Categoria/Quantità/P medio/CMP medio/POS).");
    }


    private Integer findExactInWindow(Row header, DataFormatter fmt, int start, int end, String exact) {
        for (int c = start; c <= end; c++) {
            String v = fmt.formatCellValue(header.getCell(c)).trim();
            if (v.equalsIgnoreCase(exact)) return c;
        }
        return null;
    }

    private Integer findLastExactInWindow(Row header, DataFormatter fmt, int start, int end, String exact) {
        Integer found = null;
        for (int c = start; c <= end; c++) {
            String v = fmt.formatCellValue(header.getCell(c)).trim();
            if (v.equalsIgnoreCase(exact)) found = c;
        }
        return found;
    }

    private Integer findContainsInWindow(Row header, DataFormatter fmt, int start, int end, String needleLower) {
        String needle = needleLower.toLowerCase();
        for (int c = start; c <= end; c++) {
            String v = fmt.formatCellValue(header.getCell(c)).trim().toLowerCase();
            if (v.contains(needle)) return c;
        }
        return null;
    }

    // ===========================
    // Classificazione / Ordinamento
    // ===========================

    private String classifyExposedCat(String articoloNorm) {
        String up = articoloNorm.trim().toUpperCase();

        Matcher mMP = MP_PATTERN.matcher(up);
        if (mMP.matches()) {
            int n = Integer.parseInt(mMP.group(1));
            return (n >= 1 && n <= 14) ? "MP" : null;
        }

        Matcher mP = PFP_PATTERN.matcher(up);
        if (mP.matches()) {
            int n = Integer.parseInt(mP.group(1));
            return (n >= 1 && n <= 7) ? "PFP" : null;
        }

        Matcher mV = PFV_PATTERN.matcher(up);
        if (mV.matches()) {
            int n = Integer.parseInt(mV.group(1));
            return (n >= 1 && n <= 16) ? "PFV" : null;
        }

        Matcher mA = PFA_PATTERN.matcher(up);
        if (mA.matches()) {
            int n = Integer.parseInt(mA.group(1));
            return (n >= 1 && n <= 19) ? "PFA" : null;
        }

        return null;
    }

    private String normalizeSpaces(String s) {
        return s.trim().replaceAll("\\s+", " ");
    }

    private String naturalKey(String articolo) {
        String up = articolo.toUpperCase().trim();

        Matcher m = MP_PATTERN.matcher(up);
        if (m.matches()) return "MP|" + String.format("%03d", Integer.parseInt(m.group(1)));

        Matcher p = PFP_PATTERN.matcher(up);
        if (p.matches()) return "PFP|" + String.format("%03d", Integer.parseInt(p.group(1)));

        Matcher v = PFV_PATTERN.matcher(up);
        if (v.matches()) return "PFV|" + String.format("%03d", Integer.parseInt(v.group(1)));

        Matcher a = PFA_PATTERN.matcher(up);
        if (a.matches()) return "PFA|" + String.format("%03d", Integer.parseInt(a.group(1)));

        return "ZZZ|" + up;
    }

    private FletcherCols unused() { return null; } // (ignora: placeholder per evitare warning in certi IDE)

   

    // Dummy class solo per compatibilità in alcuni build che “rompono” se file non cambia spesso.
    private static class FletcherCols {}
}
