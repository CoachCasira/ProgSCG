package view;

import javax.swing.*;
import java.awt.*;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

public class ChartsPanel extends JPanel {

    private final ChartPanel posChartPanel;
    private final ChartPanel compChartPanel;
    private final ChartPanel ceChartPanel;

    public ChartsPanel() {
        setLayout(new GridLayout(3, 1, 0, 8));

        posChartPanel = new ChartPanel(null, false);
        compChartPanel = new ChartPanel(null, false);
        ceChartPanel   = new ChartPanel(null, false);

        posChartPanel.setPreferredSize(new Dimension(1000, 260));
        compChartPanel.setPreferredSize(new Dimension(1000, 260));
        ceChartPanel.setPreferredSize(new Dimension(1000, 260));

        posChartPanel.setMouseWheelEnabled(true);
        compChartPanel.setMouseWheelEnabled(true);
        ceChartPanel.setMouseWheelEnabled(true);

        // evita casi in cui ChartPanel decide di non disegnare perché “troppo piccolo”
        posChartPanel.setMinimumDrawWidth(0);
        posChartPanel.setMinimumDrawHeight(0);
        compChartPanel.setMinimumDrawWidth(0);
        compChartPanel.setMinimumDrawHeight(0);
        ceChartPanel.setMinimumDrawWidth(0);
        ceChartPanel.setMinimumDrawHeight(0);

        add(posChartPanel);
        add(compChartPanel);
        add(ceChartPanel);
    }

    public void setPosChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            posChartPanel.setChart(chart);
            posChartPanel.revalidate();
            posChartPanel.repaint();
            this.revalidate();
            this.repaint();
        });
    }

    public void setCompChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            compChartPanel.setChart(chart);
            compChartPanel.revalidate();
            compChartPanel.repaint();
            this.revalidate();
            this.repaint();
        });
    }

    public void setCeChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ceChartPanel.setChart(chart);
            ceChartPanel.revalidate();
            ceChartPanel.repaint();
            this.revalidate();
            this.repaint();
        });
    }
}
