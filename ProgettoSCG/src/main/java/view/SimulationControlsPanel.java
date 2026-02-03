package view;

import model.ArticleRow;
import model.SimulationMode;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.util.List;

public class SimulationControlsPanel extends JPanel {

    private JComboBox<ArticleRow> cmbArticles;

    private JRadioButton rbQty;
    private JRadioButton rbPrice;

    private JTextField txtPercent;

    private JCheckBox chkCompensate;

    private JButton btnSimulate;

    // dettagli HTML
    private JEditorPane detailsPane;

    public SimulationControlsPanel() {
        super(new BorderLayout(10, 10));
        setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEtchedBorder(),
                "Controlli simulazione",
                TitledBorder.LEFT,
                TitledBorder.TOP
        ));

        JPanel form = new JPanel(new GridBagLayout());
        form.setBorder(BorderFactory.createEmptyBorder(8, 10, 8, 10));

        GridBagConstraints g = new GridBagConstraints();
        g.insets = new Insets(6, 6, 6, 6);
        g.anchor = GridBagConstraints.WEST;
        g.fill = GridBagConstraints.HORIZONTAL;

        // riga 0: articolo
        g.gridx = 0; g.gridy = 0; g.weightx = 0;
        form.add(new JLabel("Articolo:"), g);

        cmbArticles = new JComboBox<>();
        cmbArticles.setPreferredSize(new Dimension(320, 26));
        g.gridx = 1; g.gridy = 0; g.weightx = 1;
        form.add(cmbArticles, g);

        // riga 1: leva
        g.gridx = 0; g.gridy = 1; g.weightx = 0;
        form.add(new JLabel("Leva:"), g);

        JPanel modePanel = new JPanel();
        modePanel.setLayout(new BoxLayout(modePanel, BoxLayout.Y_AXIS));
        rbQty = new JRadioButton("Quantità (kg)", true);
        rbPrice = new JRadioButton("Prezzo");
        ButtonGroup bg = new ButtonGroup();
        bg.add(rbQty);
        bg.add(rbPrice);
        modePanel.add(rbQty);
        modePanel.add(rbPrice);

        g.gridx = 1; g.gridy = 1; g.weightx = 1;
        form.add(modePanel, g);

        // riga 2: percentuale
        g.gridx = 0; g.gridy = 2; g.weightx = 0;
        form.add(new JLabel("Percentuale (%) (es: 20 oppure -15):"), g);

        txtPercent = new JTextField("20");
        txtPercent.setColumns(10);
        g.gridx = 1; g.gridy = 2; g.weightx = 1;
        form.add(txtPercent, g);

        // riga 3: compensazione
        g.gridx = 0; g.gridy = 3; g.weightx = 0;
        form.add(new JLabel("Compensa:"), g);

        chkCompensate = new JCheckBox("Mantieni POS costante");
        g.gridx = 1; g.gridy = 3; g.weightx = 1;
        form.add(chkCompensate, g);

        // riga 4: bottone simula
        btnSimulate = new JButton("Simula");
        btnSimulate.setPreferredSize(new Dimension(140, 30));

        JPanel btnWrap = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        btnWrap.add(btnSimulate);

        g.gridx = 1; g.gridy = 4; g.weightx = 1;
        form.add(btnWrap, g);

        add(form, BorderLayout.NORTH);

        // =========================
        // DETTAGLI (PIÙ LARGHI + HTML WRAP)
        // =========================
        JPanel detailsOuter = new JPanel(new BorderLayout());
        detailsOuter.setBorder(BorderFactory.createTitledBorder("Dettagli"));

        detailsPane = new JEditorPane();
        detailsPane.setContentType("text/html");
        detailsPane.setEditable(false);
        detailsPane.setText("");
        detailsPane.putClientProperty(JEditorPane.HONOR_DISPLAY_PROPERTIES, Boolean.TRUE);
        detailsPane.setFont(new Font("SansSerif", Font.PLAIN, 12));

        // ✅ dimensioni: qui è il vero fix per la “finestra stretta”
        detailsPane.setPreferredSize(new Dimension(520, 520));
        detailsPane.setMinimumSize(new Dimension(520, 260));

        JScrollPane sp = new JScrollPane(detailsPane,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);

        // ✅ più “aria” e niente bordi pesanti
        sp.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));
        sp.getVerticalScrollBar().setUnitIncrement(16);

        detailsOuter.add(sp, BorderLayout.CENTER);

        // ✅ se c’è spazio, che se lo prenda tutto
        add(detailsOuter, BorderLayout.CENTER);
    }

    // =========================
    // API usata dal Controller
    // =========================

    public JButton getBtnSimulate() {
        return btnSimulate;
    }

    public void setArticles(List<ArticleRow> articles) {
        DefaultComboBoxModel<ArticleRow> m = new DefaultComboBoxModel<>();
        if (articles != null) {
            for (ArticleRow a : articles) m.addElement(a);
        }
        cmbArticles.setModel(m);
        if (m.getSize() > 0) cmbArticles.setSelectedIndex(0);
    }

    public ArticleRow getSelectedArticle() {
        Object o = cmbArticles.getSelectedItem();
        return (o instanceof ArticleRow) ? (ArticleRow) o : null;
    }

    public SimulationMode getMode() {
        return rbQty.isSelected() ? SimulationMode.QUANTITY : SimulationMode.PRICE;
    }

    public double getPercent() {
        String t = txtPercent.getText();
        if (t == null || t.trim().isEmpty()) throw new IllegalArgumentException("Inserisci una percentuale.");
        try {
            return Double.parseDouble(t.trim().replace(",", "."));
        } catch (Exception e) {
            throw new IllegalArgumentException("Percentuale non valida: " + t);
        }
    }

    public boolean isCompensateSelected() {
        return chkCompensate.isSelected();
    }

    public void setDetails(String html) {
        if (html == null) html = "";
        detailsPane.setText(html);
        detailsPane.setCaretPosition(0);
    }
}
