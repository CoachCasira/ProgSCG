package view;

import javax.swing.*;
import java.awt.*;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

public class ChartsPanel extends JPanel {

    private final ChartPanel posChartPanel;
    private final ChartPanel compChartPanel;

    public ChartsPanel() {
        setLayout(new GridLayout(2, 1, 0, 8));

        // costruttore safe (evita NPE in alcune versioni)
        posChartPanel = new ChartPanel(null, false);
        compChartPanel = new ChartPanel(null, false);

        posChartPanel.setPreferredSize(new Dimension(1000, 380));
        compChartPanel.setPreferredSize(new Dimension(1000, 380));

        posChartPanel.setMouseWheelEnabled(true);
        compChartPanel.setMouseWheelEnabled(true);

        add(posChartPanel);
        add(compChartPanel);
    }

    public void setPosChart(JFreeChart chart) {
        posChartPanel.setChart(chart);
        posChartPanel.setMouseWheelEnabled(true);
    }

    public void setCompChart(JFreeChart chart) {
        compChartPanel.setChart(chart);
        compChartPanel.setMouseWheelEnabled(true);
    }
}
