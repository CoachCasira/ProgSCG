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

        JLabel hint = new JLabel("Seleziona uno o più articoli e imposta Leva/Percentuale/Compensa per ciascuno.");
        hint.setFont(new Font("SansSerif", Font.PLAIN, 12));
        hint.setBorder(BorderFactory.createEmptyBorder(8, 10, 0, 10));
        add(hint, BorderLayout.NORTH);

        // ====== TOP: tabella + pulsanti ======
        JPanel tableBlock = new JPanel(new BorderLayout(8, 8));
        tableBlock.setBorder(BorderFactory.createEmptyBorder(0, 10, 0, 10));

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
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        tableScroll.setBorder(BorderFactory.createTitledBorder("Articoli"));

        // ✅ su laptop vogliamo che la tabella NON mangi tutto lo spazio
        tableScroll.setPreferredSize(new Dimension(520, 260));

        JPanel actions = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        btnSelectAll = new JButton("Seleziona tutti");
        btnSelectNone = new JButton("Deseleziona tutti");
        btnSimulate = new JButton("Simula");

        actions.add(btnSelectAll);
        actions.add(btnSelectNone);
        actions.add(btnSimulate);

        tableBlock.add(tableScroll, BorderLayout.CENTER);
        tableBlock.add(actions, BorderLayout.SOUTH);

        btnSelectAll.addActionListener(e -> tableModel.selectAll(true));
        btnSelectNone.addActionListener(e -> tableModel.selectAll(false));

        // ====== BOTTOM: dettagli ======
        JPanel detailsOuter = new JPanel(new BorderLayout());
        detailsOuter.setBorder(BorderFactory.createTitledBorder("Dettagli"));

        detailsPane = new JEditorPane();
        detailsPane.setContentType("text/html");
        detailsPane.setEditable(false);
        detailsPane.setText("");
        detailsPane.putClientProperty(JEditorPane.HONOR_DISPLAY_PROPERTIES, Boolean.TRUE);
        detailsPane.setFont(new Font("SansSerif", Font.PLAIN, 12));

        // ✅ dettagli più grandi su laptop
        detailsPane.setPreferredSize(new Dimension(560, 420));
        detailsPane.setMinimumSize(new Dimension(560, 260));

        JScrollPane detailsScroll = new JScrollPane(detailsPane,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        detailsScroll.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));
        detailsScroll.getVerticalScrollBar().setUnitIncrement(16);

        detailsOuter.add(detailsScroll, BorderLayout.CENTER);

        // ====== Split verticale: sopra tabella, sotto dettagli ======
        JSplitPane vSplit = new JSplitPane(JSplitPane.VERTICAL_SPLIT, tableBlock, detailsOuter);
        vSplit.setResizeWeight(0.45);      // ✅ più spazio ai dettagli
        vSplit.setDividerSize(10);
        vSplit.setContinuousLayout(true);
        vSplit.setBorder(BorderFactory.createEmptyBorder(0, 0, 10, 0));

        // Divider iniziale (buono per laptop)
        vSplit.setDividerLocation(320);

        add(vSplit, BorderLayout.CENTER);
    }

    // ===== API usata dal Controller =====

    public JButton getBtnSimulate() { return btnSimulate; }

    public void setArticles(List<ArticleRow> articles) {
        tableModel.setArticles(articles);
    }

    /** Compat: ritorna il primo selezionato. */
    public ArticleRow getSelectedArticle() {
        List<SimRequest> req = getSimulationRequests();
        return req.isEmpty() ? null : req.get(0).article;
    }

    /** Lista richieste multi-articolo */
    public List<SimRequest> getSimulationRequests() {
        return tableModel.getSelectedRequests();
    }

    /** Compat: ritorna la "mode" del primo selezionato (o QUANTITY). */
    public SimulationMode getMode() {
        List<SimRequest> req = getSimulationRequests();
        return req.isEmpty() ? SimulationMode.QUANTITY : req.get(0).mode;
    }

    /** Compat: ritorna la percent del primo selezionato (o 0). */
    public double getPercent() {
        List<SimRequest> req = getSimulationRequests();
        return req.isEmpty() ? 0.0 : req.get(0).percent;
    }

    /** true se almeno una riga selezionata ha compensazione */
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

        private final String[] cols = {"Sel", "Cat", "Articolo", "Leva", "%", "Compensa"};

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
                    rows.add(r);
                }
            }
            fireTableDataChanged();
        }

        public void selectAll(boolean v) {
            for (RowState r : rows) r.selected = v;
            fireTableRowsUpdated(0, Math.max(0, rows.size() - 1));
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
                        r.selected = (aValue instanceof Boolean) ? (Boolean) aValue : false;
                        break;

                    case 3: {
                        String s = String.valueOf(aValue).toUpperCase().trim();
                        r.mode = s.contains("PREZZ") ? SimulationMode.PRICE : SimulationMode.QUANTITY;
                        break;
                    }

                    case 4: {
                        if (aValue instanceof Double) r.percent = (Double) aValue;
                        else r.percent = Double.parseDouble(String.valueOf(aValue).trim().replace(",", "."));
                        break;
                    }

                    case 5:
                        r.compensate = (aValue instanceof Boolean) ? (Boolean) aValue : false;
                        break;
                }
            } catch (Exception ignore) {
                // input non valido: non aggiorno
            }
            fireTableCellUpdated(rowIndex, columnIndex);
        }
    }
}