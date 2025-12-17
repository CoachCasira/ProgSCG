package view;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

import javax.swing.*;
import java.awt.*;

public class PremioCompFrame extends JFrame {

    private final SimulationControlsPanel controlsPanel;

    private final ChartPanel posChartPanel;
    private final ChartPanel premioChartPanel;
    private final JButton btnReset;


    private final JButton btnBack;

    public PremioCompFrame() {
        super("Simulazione con compensazione PREMIO (POS costante)");

        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        setSize(1200, 800);
        setLocationRelativeTo(null);

        // top bar: solo "Indietro"
        JPanel top = new JPanel(new FlowLayout(FlowLayout.RIGHT));

        btnReset = new JButton("Reset Excel");   // ✅
        btnBack = new JButton("Indietro");

        top.add(btnReset);                       // ✅
        top.add(btnBack);


        // ✅ nuovo pannello controlli (NON riusare quello della main)
        controlsPanel = new SimulationControlsPanel();

        // charts container (scrollabile)
        JPanel chartsContainer = new JPanel();
        chartsContainer.setLayout(new BoxLayout(chartsContainer, BoxLayout.Y_AXIS));

        posChartPanel = new ChartPanel(null, false);
        premioChartPanel = new ChartPanel(null, false);

        posChartPanel.setPreferredSize(new Dimension(1100, 300));
        premioChartPanel.setPreferredSize(new Dimension(1100, 300));

        chartsContainer.add(posChartPanel);
        chartsContainer.add(Box.createVerticalStrut(10));
        chartsContainer.add(premioChartPanel);

        JScrollPane scroll = new JScrollPane(chartsContainer);
        scroll.getVerticalScrollBar().setUnitIncrement(16);

        // layout principale: a sinistra controlli, a destra grafici
        setLayout(new BorderLayout());
        add(top, BorderLayout.NORTH);
        add(controlsPanel, BorderLayout.WEST);
        add(scroll, BorderLayout.CENTER);
    }

    public JButton getBtnReset() { return btnReset; }
    public JButton getBtnBack() { return btnBack; }

    public SimulationControlsPanel getControlsPanel() { return controlsPanel; }

    public void setPosChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            posChartPanel.setChart(chart);
            posChartPanel.revalidate();
            posChartPanel.repaint();
        });
    }

    public void setPremioChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            premioChartPanel.setChart(chart);
            premioChartPanel.revalidate();
            premioChartPanel.repaint();
        });
    }
}
