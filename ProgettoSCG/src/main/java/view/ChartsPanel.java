package view;

import javax.swing.*;
import java.awt.*;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

public class ChartsPanel extends JPanel {

    private final ChartPanel posChartPanel;
    private final ChartPanel compChartPanel;

    // CE separati
    private final ChartPanel ceBaseChartPanel;
    private final ChartPanel ceVarChartPanel;

    // CE Delta
    private final ChartPanel ceDeltaChartPanel;

    public ChartsPanel() {
        setLayout(new BorderLayout());

        JPanel content = new JPanel();
        content.setLayout(new GridLayout(5, 1, 0, 8));

        posChartPanel = new ChartPanel(null, false);
        compChartPanel = new ChartPanel(null, false);

        ceBaseChartPanel = new ChartPanel(null, false);
        ceVarChartPanel  = new ChartPanel(null, false);
        ceDeltaChartPanel = new ChartPanel(null, false);

        // Dimensioni coerenti con la tua UI
        Dimension pref = new Dimension(1000, 260);

        posChartPanel.setPreferredSize(pref);
        compChartPanel.setPreferredSize(pref);
        ceBaseChartPanel.setPreferredSize(pref);
        ceVarChartPanel.setPreferredSize(pref);
        ceDeltaChartPanel.setPreferredSize(pref);

        posChartPanel.setMouseWheelEnabled(true);
        compChartPanel.setMouseWheelEnabled(true);
        ceBaseChartPanel.setMouseWheelEnabled(true);
        ceVarChartPanel.setMouseWheelEnabled(true);
        ceDeltaChartPanel.setMouseWheelEnabled(true);

        // evita casi in cui ChartPanel decide di non disegnare perché “troppo piccolo”
        setMinDraw(posChartPanel);
        setMinDraw(compChartPanel);
        setMinDraw(ceBaseChartPanel);
        setMinDraw(ceVarChartPanel);
        setMinDraw(ceDeltaChartPanel);

        content.add(posChartPanel);
        content.add(compChartPanel);
        content.add(ceBaseChartPanel);
        content.add(ceVarChartPanel);
        content.add(ceDeltaChartPanel);

        JScrollPane scroll = new JScrollPane(content);
        scroll.setBorder(null);
        scroll.getVerticalScrollBar().setUnitIncrement(16);

        add(scroll, BorderLayout.CENTER);
    }

    private void setMinDraw(ChartPanel p) {
        p.setMinimumDrawWidth(0);
        p.setMinimumDrawHeight(0);
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
}
