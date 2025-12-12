package model;

public class ArticleRow {

    private final String cat;        // MP / PFP / PFV / PFA
    private final String articolo;
    private final String categoria;
    private final int rowIndex;

    private final int colQty;

    // colonne candidate (possono essere -1 se non trovate)
    private final int colPmedioEUR;
    private final int colPmedioUSD;
    private final int colCMPmedioEUR;
    private final int colCMPmedioUSD;

    private final int colPos;

    public ArticleRow(String cat,
                      String articolo,
                      String categoria,
                      int rowIndex,
                      int colQty,
                      int colPmedioEUR,
                      int colPmedioUSD,
                      int colCMPmedioEUR,
                      int colCMPmedioUSD,
                      int colPos) {
        this.cat = cat;
        this.articolo = articolo;
        this.categoria = categoria;
        this.rowIndex = rowIndex;
        this.colQty = colQty;
        this.colPmedioEUR = colPmedioEUR;
        this.colPmedioUSD = colPmedioUSD;
        this.colCMPmedioEUR = colCMPmedioEUR;
        this.colCMPmedioUSD = colCMPmedioUSD;
        this.colPos = colPos;
    }

    public String getCat() { return cat; }
    public String getArticolo() { return articolo; }
    public String getCategoria() { return categoria; }
    public int getRowIndex() { return rowIndex; }

    public int getColQty() { return colQty; }
    public int getColPos() { return colPos; }

    public int getColPmedioEUR() { return colPmedioEUR; }
    public Integer getColPmedioUSD() { return colPmedioUSD >= 0 ? colPmedioUSD : null; }

    public int getColCMPmedioEUR() { return colCMPmedioEUR; }
    public Integer getColCMPmedioUSD() { return colCMPmedioUSD >= 0 ? colCMPmedioUSD : null; }

    @Override
    public String toString() {
        return articolo + " (" + cat + ")";
    }

	public boolean isPriceIsUSDInput() {
		// TODO Auto-generated method stub
		return false;
	}
}
