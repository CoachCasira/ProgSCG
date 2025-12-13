package view;

import javax.swing.*;
import java.awt.*;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

public class ChartsPanel extends JPanel {

    private final ChartPanel posChartPanel;
    private final ChartPanel compChartPanel;

    // ✅ 3 grafici CE separati
    private final ChartPanel ceBaseChartPanel;
    private final ChartPanel ceVarChartPanel;
    private final ChartPanel ceDeltaChartPanel;

    public ChartsPanel() {
        setLayout(new BorderLayout());

        JPanel inner = new JPanel();
        inner.setLayout(new GridLayout(5, 1, 0, 8));
        inner.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));

        posChartPanel = new ChartPanel(null, false);
        compChartPanel = new ChartPanel(null, false);

        ceBaseChartPanel  = new ChartPanel(null, false);
        ceVarChartPanel   = new ChartPanel(null, false);
        ceDeltaChartPanel = new ChartPanel(null, false);

        setupPanel(posChartPanel);
        setupPanel(compChartPanel);
        setupPanel(ceBaseChartPanel);
        setupPanel(ceVarChartPanel);
        setupPanel(ceDeltaChartPanel);

        inner.add(posChartPanel);
        inner.add(compChartPanel);
        inner.add(ceBaseChartPanel);
        inner.add(ceVarChartPanel);
        inner.add(ceDeltaChartPanel);

        JScrollPane sp = new JScrollPane(inner);
        sp.setBorder(null);
        sp.getVerticalScrollBar().setUnitIncrement(16);

        add(sp, BorderLayout.CENTER);
    }

    private void setupPanel(ChartPanel cp) {
        cp.setPreferredSize(new Dimension(1000, 260));
        cp.setMouseWheelEnabled(true);
        cp.setMinimumDrawWidth(0);
        cp.setMinimumDrawHeight(0);
    }

    public void setPosChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            posChartPanel.setChart(chart);
            posChartPanel.revalidate();
            posChartPanel.repaint();
        });
    }

    public void setCompChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            compChartPanel.setChart(chart);
            compChartPanel.revalidate();
            compChartPanel.repaint();
        });
    }

    public void setCeBaseChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ceBaseChartPanel.setChart(chart);
            ceBaseChartPanel.revalidate();
            ceBaseChartPanel.repaint();
        });
    }

    public void setCeVarChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ceVarChartPanel.setChart(chart);
            ceVarChartPanel.revalidate();
            ceVarChartPanel.repaint();
        });
    }

    public void setCeDeltaChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ceDeltaChartPanel.setChart(chart);
            ceDeltaChartPanel.revalidate();
            ceDeltaChartPanel.repaint();
        });
    }

    // (facoltativo: compatibilità se da qualche parte chiamavi ancora setCeChart)
    public void setCeChart(JFreeChart chart) {
        setCeVarChart(chart);
    }
}
