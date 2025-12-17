package view;

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;

import javax.swing.*;

import model.ArticleRow;
import model.SimulationMode;

/**
 * Controlli simulazione (versione "orale"):
 * - scegli articolo
 * - scegli leva (Quantità / Prezzo)
 * - inserisci UNA percentuale (es. 20 oppure -15)
 * - simula
 */
public class SimulationControlsPanel extends JPanel {

    private JComboBox<ArticleRow> cmbArticle;
    private JRadioButton rbQty;
    private JRadioButton rbPrice;

    private JTextField txtPercent;

    private JButton btnSimulate;

    private JTextArea txtDetails;

    public SimulationControlsPanel() {
        setLayout(new GridBagLayout());
        setBorder(BorderFactory.createTitledBorder("Controlli simulazione"));

        cmbArticle = new JComboBox<>();
        rbQty = new JRadioButton("Quantità (kg)", true);
        rbPrice = new JRadioButton("Prezzo", false);

        ButtonGroup bg = new ButtonGroup();
        bg.add(rbQty);
        bg.add(rbPrice);

        // ✅ un solo campo percentuale
        txtPercent = new JTextField("20", 8);

        btnSimulate = new JButton("Simula");

        txtDetails = new JTextArea(10, 24);
        txtDetails.setEditable(false);
        txtDetails.setLineWrap(true);
        txtDetails.setWrapStyleWord(true);
        txtDetails.setBorder(BorderFactory.createTitledBorder("Dettagli"));

        GridBagConstraints c = new GridBagConstraints();
        c.insets = new Insets(6, 8, 6, 8);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 0; c.gridy = 0;

        add(new JLabel("Articolo:"), c);
        c.gridy++;
        add(cmbArticle, c);

        c.gridy++;
        add(new JLabel("Leva:"), c);
        c.gridy++;
        add(rbQty, c);
        c.gridy++;
        add(rbPrice, c);

        c.gridy++;
        add(new JLabel("Percentuale (%) (es: 20 oppure -15):"), c);
        c.gridy++;
        add(txtPercent, c);

        c.gridy++;
        add(btnSimulate, c);

        c.gridy++;
        c.fill = GridBagConstraints.BOTH;
        c.weighty = 1.0;
        add(new JScrollPane(txtDetails), c);
    }

    public void setArticles(java.util.List<ArticleRow> rows) {
        DefaultComboBoxModel<ArticleRow> m = new DefaultComboBoxModel<>();
        for (ArticleRow r : rows) m.addElement(r);
        cmbArticle.setModel(m);
    }

    public ArticleRow getSelectedArticle() {
        return (ArticleRow) cmbArticle.getSelectedItem();
    }

    public SimulationMode getMode() {
        return rbQty.isSelected() ? SimulationMode.QUANTITY : SimulationMode.PRICE;
    }

    public double getPercent() {
        String raw = txtPercent.getText().trim();
        if (raw.isEmpty()) throw new IllegalArgumentException("Percentuale non valida: campo vuoto");

        raw = raw.replace("%", "").trim().replace(",", ".");

        try {
            return Double.parseDouble(raw);
        } catch (Exception e) {
            throw new IllegalArgumentException("Percentuale non valida: '" + txtPercent.getText() + "'");
        }
    }


    public JButton getBtnSimulate() { return btnSimulate; }

    public void setDetails(String text) { txtDetails.setText(text); }
}
