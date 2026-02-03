package view;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.Font;

import javax.swing.*;

public class MainFrame extends JFrame {

    private JButton btnLoadExcel;
    private JButton btnExit;
    private JButton btnOpenWorkingCopy;
    private JButton btnShowCeBudget;
    private JButton btnResetExcel;

    private JLabel lblStatus;
    private JLabel lblFileName;

    private JButton btnShowPremioComp;
    private SimulationControlsPanel controlsPanel;
    private ChartsPanel chartsPanel;

    public MainFrame() {
        super("Progetto SCG - Budget & Grafici");

        setSize(1200, 700);
        setMinimumSize(new Dimension(1000, 650));
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout(12, 12));

        // Header
        JPanel header = new JPanel(new BorderLayout());
        header.setBorder(BorderFactory.createEmptyBorder(14, 18, 6, 18));

        JLabel title = new JLabel("Budget vendite 2022 – Simulazioni & Grafici");
        title.setFont(new Font("SansSerif", Font.BOLD, 22));

        JLabel subtitle = new JLabel("Carica il file Excel fornito dal docente.");
        subtitle.setFont(new Font("SansSerif", Font.PLAIN, 13));

        header.add(title, BorderLayout.NORTH);
        header.add(subtitle, BorderLayout.CENTER);

        // Stato + bottoni file
        JPanel status = new JPanel(new BorderLayout(8, 8));
        status.setBorder(BorderFactory.createTitledBorder("Stato"));

        lblStatus = new JLabel();
        lblStatus.setFont(new Font("SansSerif", Font.BOLD, 14));
        lblFileName = new JLabel();
        lblFileName.setFont(new Font("SansSerif", Font.PLAIN, 13));

        JPanel statusLeft = new JPanel(new BorderLayout());
        statusLeft.add(lblStatus, BorderLayout.NORTH);
        statusLeft.add(lblFileName, BorderLayout.CENTER);

        JPanel statusRight = new JPanel();

        btnLoadExcel = new JButton("Carica Excel");
        btnExit = new JButton("Esci");

        btnOpenWorkingCopy = new JButton("Apri copia");
        btnOpenWorkingCopy.setEnabled(false);

        btnShowCeBudget = new JButton("CE Budget 2022");
        btnShowCeBudget.setEnabled(false);

        btnShowPremioComp = new JButton("Compensa Premio");
        btnShowPremioComp.setEnabled(false);

        btnResetExcel = new JButton("Reset Excel");
        btnResetExcel.setEnabled(false);

        statusRight.add(btnShowPremioComp);
        statusRight.add(btnLoadExcel);
        statusRight.add(btnOpenWorkingCopy);
        statusRight.add(btnResetExcel);
        statusRight.add(btnShowCeBudget);
        statusRight.add(btnExit); // ✅ UNA SOLA VOLTA

        status.add(statusLeft, BorderLayout.CENTER);
        status.add(statusRight, BorderLayout.EAST);

        JPanel top = new JPanel(new BorderLayout(10, 10));
        top.setBorder(BorderFactory.createEmptyBorder(0, 18, 0, 18));
        top.add(header, BorderLayout.NORTH);
        top.add(status, BorderLayout.CENTER);

        // Center: controlli + grafici
        controlsPanel = new SimulationControlsPanel();
        chartsPanel = new ChartsPanel();

        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, controlsPanel, chartsPanel);
        split.setDividerSize(10);

        // ✅ pannello sinistro più largo (controlli + dettagli)
        final int leftWidth = 520;
        controlsPanel.setPreferredSize(new Dimension(leftWidth, 0));
        controlsPanel.setMinimumSize(new Dimension(480, 0));

        // ✅ lascia più spazio ai grafici ma non schiaccia i dettagli
        split.setResizeWeight(0.40);

        // ✅ divider location fatto nel momento giusto
        SwingUtilities.invokeLater(() -> split.setDividerLocation(leftWidth));

        add(top, BorderLayout.NORTH);
        add(split, BorderLayout.CENTER);

        setExcelNotLoaded();
    }

    public JButton getBtnResetExcel() { return btnResetExcel; }
    public JButton getBtnShowPremioComp() { return btnShowPremioComp; }
    public JButton getBtnLoadExcel() { return btnLoadExcel; }
    public JButton getBtnExit() { return btnExit; }
    public JButton getBtnOpenWorkingCopy() { return btnOpenWorkingCopy; }
    public JButton getBtnShowCeBudget() { return btnShowCeBudget; }

    public SimulationControlsPanel getControlsPanel() { return controlsPanel; }
    public ChartsPanel getChartsPanel() { return chartsPanel; }

    public void setExcelLoaded(String fileName) {
        lblStatus.setText("Excel caricato");
        lblFileName.setText("File: " + fileName);
        btnOpenWorkingCopy.setEnabled(true);
        btnShowCeBudget.setEnabled(true);
        btnShowPremioComp.setEnabled(true);
        btnResetExcel.setEnabled(true);
    }

    public void setExcelNotLoaded() {
        lblStatus.setText("Excel non caricato");
        lblFileName.setText("Seleziona un file .xlsx per iniziare");
        btnOpenWorkingCopy.setEnabled(false);
        btnShowCeBudget.setEnabled(false);
        btnShowPremioComp.setEnabled(false);
        btnResetExcel.setEnabled(false);
    }
}