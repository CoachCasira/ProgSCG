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
        setSize(1300, 850);
        setLocationRelativeTo(null);

        // ===== Top bar =====
        JPanel top = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 6));
        btnReset = new JButton("Reset Excel");
        btnBack  = new JButton("Indietro");
        top.add(btnReset);
        top.add(btnBack);

        // ===== Left: controlli =====
        controlsPanel = new SimulationControlsPanel();
        controlsPanel.setPreferredSize(new Dimension(560, 0));
        controlsPanel.setMinimumSize(new Dimension(520, 0));

        // ===== Right: grafici (scrollabili) =====
        posChartPanel = new ChartPanel(null, false);
        premioChartPanel = new ChartPanel(null, false);

        setupChartPanel(posChartPanel);
        setupChartPanel(premioChartPanel);

        JPanel chartsContainer = new JPanel();
        chartsContainer.setLayout(new BoxLayout(chartsContainer, BoxLayout.Y_AXIS));
        chartsContainer.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        chartsContainer.add(wrapWithTitle("POS Totale", posChartPanel));
        chartsContainer.add(Box.createVerticalStrut(12));
        chartsContainer.add(wrapWithTitle("Leva + Premio", premioChartPanel));

        JScrollPane scroll = new JScrollPane(chartsContainer,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        scroll.getVerticalScrollBar().setUnitIncrement(16);
        scroll.setBorder(BorderFactory.createTitledBorder("Grafici"));

        // ===== Split pane per “allineamento” perfetto =====
        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, controlsPanel, scroll);
        split.setResizeWeight(0.42);
        split.setDividerSize(10);
        split.setDividerLocation(560);

        setLayout(new BorderLayout(10, 10));
        add(top, BorderLayout.NORTH);
        add(split, BorderLayout.CENTER);
    }

    private JPanel wrapWithTitle(String title, JComponent content) {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(BorderFactory.createTitledBorder(title));
        p.add(content, BorderLayout.CENTER);
        return p;
    }

    private void setupChartPanel(ChartPanel cp) {
        cp.setPreferredSize(new Dimension(1000, 340));

        // niente zoom / popup
        cp.setMouseWheelEnabled(false);
        cp.setMouseZoomable(false);
        cp.setDomainZoomable(false);
        cp.setRangeZoomable(false);
        cp.setPopupMenu(null);

        cp.setOpaque(true);
        cp.setBackground(Color.WHITE);
        cp.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));

        cp.setMinimumDrawWidth(0);
        cp.setMinimumDrawHeight(0);
    }

    public JButton getBtnReset() { return btnReset; }
    public JButton getBtnBack()  { return btnBack; }

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
