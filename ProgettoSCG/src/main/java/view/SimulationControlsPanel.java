package view;

import model.ArticleRow;
import model.SimulationMode;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.TableColumn;
import java.awt.*;
import java.util.ArrayList;
import java.util.List;

public class SimulationControlsPanel extends JPanel {

    public static class SimRequest {
        public final ArticleRow article;
        public final SimulationMode mode;
        public final double percent;
        public final boolean compensate;

        public SimRequest(ArticleRow article, SimulationMode mode, double percent, boolean compensate) {
            this.article = article;
            this.mode = mode;
            this.percent = percent;
            this.compensate = compensate;
        }
    }

    private JTable table;
    private ArticlesTableModel tableModel;

    private JButton btnSelectAll;
    private JButton btnSelectNone;
    private JButton btnSimulate;

    private JEditorPane detailsPane;

    public SimulationControlsPanel() {
        super(new BorderLayout(10, 10));
        setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEtchedBorder(),
                "Controlli simulazione",
                TitledBorder.LEFT,
                TitledBorder.TOP
        ));

        // ===== TOP: tabella multi-articolo =====
        JPanel top = new JPanel(new BorderLayout(8, 8));
        top.setBorder(BorderFactory.createEmptyBorder(8, 10, 0, 10));

        JLabel hint = new JLabel("Seleziona uno o più articoli e imposta Leva/Percentuale/Compensa per ciascuno.");
        hint.setFont(new Font("SansSerif", Font.PLAIN, 12));
        top.add(hint, BorderLayout.NORTH);

        tableModel = new ArticlesTableModel();
        table = new JTable(tableModel);
        table.setRowHeight(24);
        table.setFillsViewportHeight(true);
        table.setAutoCreateRowSorter(true);

        // Colonna percentuale allineata a destra
        DefaultTableCellRenderer right = new DefaultTableCellRenderer();
        right.setHorizontalAlignment(SwingConstants.RIGHT);
        table.getColumnModel().getColumn(4).setCellRenderer(right);

        // Editor per "Leva"
        TableColumn modeCol = table.getColumnModel().getColumn(3);
        JComboBox<String> modeBox = new JComboBox<String>(new String[]{"QUANTITA", "PREZZO"});
        modeCol.setCellEditor(new DefaultCellEditor(modeBox));

        JScrollPane tableScroll = new JScrollPane(table,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        tableScroll.setBorder(BorderFactory.createTitledBorder("Articoli"));

        top.add(tableScroll, BorderLayout.CENTER);

        // pulsanti rapidi + simula
        JPanel actions = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        btnSelectAll = new JButton("Seleziona tutti");
        btnSelectNone = new JButton("Deseleziona tutti");
        btnSimulate = new JButton("Simula");

        actions.add(btnSelectAll);
        actions.add(btnSelectNone);
        actions.add(btnSimulate);

        top.add(actions, BorderLayout.SOUTH);

        add(top, BorderLayout.NORTH);

        btnSelectAll.addActionListener(e -> tableModel.selectAll(true));
        btnSelectNone.addActionListener(e -> tableModel.selectAll(false));

        // ===== CENTER: dettagli (HTML) =====
        JPanel detailsOuter = new JPanel(new BorderLayout());
        detailsOuter.setBorder(BorderFactory.createTitledBorder("Dettagli"));

        detailsPane = new JEditorPane();
        detailsPane.setContentType("text/html");
        detailsPane.setEditable(false);
        detailsPane.setText("");
        detailsPane.putClientProperty(JEditorPane.HONOR_DISPLAY_PROPERTIES, Boolean.TRUE);
        detailsPane.setFont(new Font("SansSerif", Font.PLAIN, 12));

        // Dimensioni “ampie”
        detailsPane.setPreferredSize(new Dimension(560, 520));
        detailsPane.setMinimumSize(new Dimension(560, 260));

        JScrollPane sp = new JScrollPane(detailsPane,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        sp.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));
        sp.getVerticalScrollBar().setUnitIncrement(16);

        detailsOuter.add(sp, BorderLayout.CENTER);
        add(detailsOuter, BorderLayout.CENTER);
    }

    // ===== API usata dal Controller =====

    public JButton getBtnSimulate() {
        return btnSimulate;
    }

    public void setArticles(List<ArticleRow> articles) {
        tableModel.setArticles(articles);
    }

    /** Compat: ritorna il primo selezionato (se ti serve in altri punti). */
    public ArticleRow getSelectedArticle() {
        List<SimRequest> req = getSimulationRequests();
        return req.isEmpty() ? null : req.get(0).article;
    }

    /** Lista richieste multi-articolo */
    public List<SimRequest> getSimulationRequests() {
        return tableModel.getSelectedRequests();
    }

 // ======================================================
 // COMPATIBILITÀ con vecchi controller (PremioCompController)
 // ======================================================

 /** Ritorna la percentuale della prima riga selezionata. */
 public double getPercent() {
     List<SimRequest> req = getSimulationRequests();
     if (req == null || req.isEmpty()) {
         throw new IllegalArgumentException("Seleziona almeno un articolo (colonna 'Sel').");
     }
     return req.get(0).percent;
 }

 /** Ritorna la leva (QUANTITY/PRICE) della prima riga selezionata. */
 public SimulationMode getMode() {
     List<SimRequest> req = getSimulationRequests();
     if (req == null || req.isEmpty()) {
         return SimulationMode.QUANTITY; // default sensato
     }
     return req.get(0).mode;
 }

    
    
    /** Compat: se ti serve ancora una “compensa globale” altrove, qui vale “true se almeno 1 spuntato” */
    public boolean isCompensateSelected() {
        for (SimRequest r : getSimulationRequests()) {
            if (r.compensate) return true;
        }
        return false;
    }

    public void setDetails(String html) {
        if (html == null) html = "";
        detailsPane.setText(html);
        detailsPane.setCaretPosition(0);
    }

    // ===== TableModel =====

    private static class ArticlesTableModel extends AbstractTableModel {

        private final String[] cols = new String[] {
                "Sel", "Cat", "Articolo", "Leva", "%", "Compensa"
        };

        private static class RowState {
            boolean selected = false;
            ArticleRow article;
            SimulationMode mode = SimulationMode.QUANTITY;
            double percent = 0.0;
            boolean compensate = false;
        }

        private final List<RowState> rows = new ArrayList<RowState>();

        public void setArticles(List<ArticleRow> articles) {
            rows.clear();
            if (articles != null) {
                for (ArticleRow a : articles) {
                    RowState r = new RowState();
                    r.article = a;
                    r.selected = false;
                    r.mode = SimulationMode.QUANTITY;
                    r.percent = 0.0;
                    r.compensate = false;
                    rows.add(r);
                }
            }
            fireTableDataChanged();
        }

        public void selectAll(boolean v) {
            for (RowState r : rows) r.selected = v;
            if (!rows.isEmpty()) fireTableRowsUpdated(0, rows.size() - 1);
        }

        public List<SimRequest> getSelectedRequests() {
            List<SimRequest> out = new ArrayList<SimRequest>();
            for (RowState r : rows) {
                if (!r.selected) continue;
                out.add(new SimRequest(r.article, r.mode, r.percent, r.compensate));
            }
            return out;
        }

        @Override public int getRowCount() { return rows.size(); }
        @Override public int getColumnCount() { return cols.length; }
        @Override public String getColumnName(int column) { return cols[column]; }

        @Override
        public Class<?> getColumnClass(int columnIndex) {
            if (columnIndex == 0) return Boolean.class;
            if (columnIndex == 4) return Double.class;
            if (columnIndex == 5) return Boolean.class;
            return String.class;
        }

        @Override
        public boolean isCellEditable(int rowIndex, int columnIndex) {
            // Cat/Articolo read-only
            return columnIndex != 1 && columnIndex != 2;
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            RowState r = rows.get(rowIndex);
            switch (columnIndex) {
                case 0: return r.selected;
                case 1: return (r.article.getCat() == null ? "" : r.article.getCat());
                case 2: return (r.article.getArticolo() == null ? "" : r.article.getArticolo());
                case 3: return (r.mode == SimulationMode.QUANTITY ? "QUANTITA" : "PREZZO");
                case 4: return r.percent;
                case 5: return r.compensate;
                default: return "";
            }
        }

        @Override
        public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
            RowState r = rows.get(rowIndex);
            try {
                switch (columnIndex) {
                    case 0:
                        r.selected = (aValue instanceof Boolean) && ((Boolean) aValue);
                        break;
                    case 3: {
                        String s = String.valueOf(aValue).toUpperCase().trim();
                        r.mode = s.contains("PREZZ") ? SimulationMode.PRICE : SimulationMode.QUANTITY;
                        break;
                    }
                    case 4: {
                        if (aValue instanceof Double) {
                            r.percent = ((Double) aValue);
                        } else {
                            String s = String.valueOf(aValue).trim().replace(",", ".");
                            r.percent = Double.parseDouble(s);
                        }
                        break;
                    }
                    case 5:
                        r.compensate = (aValue instanceof Boolean) && ((Boolean) aValue);
                        break;
                }
            } catch (Exception ignore) {
                // input non valido: non aggiorno
            }
            fireTableCellUpdated(rowIndex, columnIndex);
        }
    }
}
