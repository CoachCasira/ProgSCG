package view;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

import javax.swing.*;
import java.awt.*;

public class CeBudgetFrame extends JFrame {

    private final JButton btnBack;
    private final ChartPanel chartPanel;

    public CeBudgetFrame() {
        super("CE Budget 2022 - Base (Excel)");

        setSize(1200, 700);
        setMinimumSize(new Dimension(1000, 650));
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        setLayout(new BorderLayout(12, 12));

        JPanel top = new JPanel(new BorderLayout());
        top.setBorder(BorderFactory.createEmptyBorder(10, 12, 0, 12));

        JLabel title = new JLabel("CE Budget 2022 – Valori base (Excel originale, senza modifiche)");
        title.setFont(new Font("SansSerif", Font.BOLD, 18));

        btnBack = new JButton("Indietro");

        top.add(title, BorderLayout.WEST);
        top.add(btnBack, BorderLayout.EAST);

        chartPanel = new ChartPanel(null, false);
        chartPanel.setMouseWheelEnabled(true);
        chartPanel.setMinimumDrawWidth(0);
        chartPanel.setMinimumDrawHeight(0);

        add(top, BorderLayout.NORTH);
        add(chartPanel, BorderLayout.CENTER);
    }

    public JButton getBtnBack() {
        return btnBack;
    }

    public void setChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            chartPanel.setChart(chart);
            chartPanel.revalidate();
            chartPanel.repaint();
        });
    }
}
