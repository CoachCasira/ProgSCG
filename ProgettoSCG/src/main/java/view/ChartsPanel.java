package view;

import javax.swing.*;
import java.awt.*;
import java.util.LinkedHashMap;
import java.util.Map;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

public class ChartsPanel extends JPanel {

    private final JTabbedPane tabs = new JTabbedPane();

    // Per ogni articoloKey manteniamo 2 tab: POS e VAR
    private final Map<String, PairTabs> byArticle = new LinkedHashMap<>();

    // fallback per API vecchie senza key
    private String lastActiveArticleKey = null;

    public ChartsPanel() {
        super(new BorderLayout());
        tabs.setBorder(BorderFactory.createTitledBorder("Grafici"));
        add(tabs, BorderLayout.CENTER);
    }

    // ============================================================
    // API "nuova" (articolo esplicito)
    // ============================================================

    public void setPosChart(String articleKey, JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            PairTabs p = getOrCreatePair(articleKey);
            p.posPanel.setChart(chart);
            p.posPanel.revalidate();
            p.posPanel.repaint();
            lastActiveArticleKey = normalizeKey(articleKey);
        });
    }

    public void setCompChart(String articleKey, JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            PairTabs p = getOrCreatePair(articleKey);
            p.varPanel.setChart(chart);
            p.varPanel.revalidate();
            p.varPanel.repaint();
            lastActiveArticleKey = normalizeKey(articleKey);
        });
    }

    /**
     * Seleziona il tab POS (quello “principale”) dell’articolo.
     * Se vuoi selezionare direttamente VAR, usa setActiveVariationTab(articleKey).
     */
    public void setActiveArticleTab(String articleKey) {
        SwingUtilities.invokeLater(() -> {
            PairTabs p = getOrCreatePair(articleKey);
            int idx = tabs.indexOfComponent(p.posRoot);
            if (idx >= 0) tabs.setSelectedIndex(idx);
            lastActiveArticleKey = normalizeKey(articleKey);
        });
    }

    public void setActiveVariationTab(String articleKey) {
        SwingUtilities.invokeLater(() -> {
            PairTabs p = getOrCreatePair(articleKey);
            int idx = tabs.indexOfComponent(p.varRoot);
            if (idx >= 0) tabs.setSelectedIndex(idx);
            lastActiveArticleKey = normalizeKey(articleKey);
        });
    }

    // ============================================================
    // API "vecchia" (compatibilità) usata dal MainController
    // ============================================================

    public void addOrReplaceArticleCharts(String labelTab, String subtitle, JFreeChart posChart, JFreeChart compChart) {
        // Crea/aggiorna entrambe le schede
        setPosChart(labelTab, posChart);
        setCompChart(labelTab, compChart);

        // Rimani sul tab POS (come comportamento “naturale”)
        setActiveArticleTab(labelTab);
    }

    public void clearArticleCharts() {
        clearAll();
    }

    // versioni senza key: applico all’ultima attiva / selezionata (POS di default)
    public void setPosChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            PairTabs p = getOrCreateActivePair();
            p.posPanel.setChart(chart);
            p.posPanel.revalidate();
            p.posPanel.repaint();
        });
    }

    public void setCompChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            PairTabs p = getOrCreateActivePair();
            p.varPanel.setChart(chart);
            p.varPanel.revalidate();
            p.varPanel.repaint();
        });
    }

    // ============================================================
    // CE charts: NON più mostrati qui. NO-OP per compatibilità.
    // ============================================================

    public void setCeBaseChart(JFreeChart chart) { /* NO-OP */ }
    public void setCeVarChart(JFreeChart chart)  { /* NO-OP */ }
    public void setCeDeltaChart(JFreeChart chart){ /* NO-OP */ }

    // ============================================================
    // Utility
    // ============================================================

    public void clearAll() {
        SwingUtilities.invokeLater(() -> {
            byArticle.clear();
            tabs.removeAll();
            lastActiveArticleKey = null;
            revalidate();
            repaint();
        });
    }

    private PairTabs getOrCreateActivePair() {
        // 1) tab selezionata -> capisco a quale articolo appartiene
        int sel = tabs.getSelectedIndex();
        if (sel >= 0) {
            Component comp = tabs.getComponentAt(sel);
            String key = findKeyByComponent(comp);
            if (key != null) {
                lastActiveArticleKey = key;
                return byArticle.get(key);
            }
        }

        // 2) ultima attiva
        if (lastActiveArticleKey != null && byArticle.containsKey(lastActiveArticleKey)) {
            return byArticle.get(lastActiveArticleKey);
        }

        // 3) fallback
        return getOrCreatePair("Risultati");
    }

    private String findKeyByComponent(Component comp) {
        for (Map.Entry<String, PairTabs> e : byArticle.entrySet()) {
            PairTabs p = e.getValue();
            if (p.posRoot == comp || p.varRoot == comp) return e.getKey();
        }
        return null;
    }

    private PairTabs getOrCreatePair(String articleKey) {
        String key = normalizeKey(articleKey);

        PairTabs existing = byArticle.get(key);
        if (existing != null) return existing;

        PairTabs p = new PairTabs();
        p.articleKey = key;

        // --- POS tab
        p.posPanel = new ChartPanel(null, false);
        setupChartPanel(p.posPanel);

        JPanel posWrap = wrapWithTitle("POS", p.posPanel);
        p.posRoot = buildScrollableRoot(posWrap);

        // --- VAR tab
        p.varPanel = new ChartPanel(null, false);
        setupChartPanel(p.varPanel);

        JPanel varWrap = wrapWithTitle("Variazione (Quantità / Prezzo)", p.varPanel);
        p.varRoot = buildScrollableRoot(varWrap);

        // label dei tab
        String labelPOS = key;                 // es: "MP||MP1"
        String labelVAR = key + " – VAR";      // es: "MP||MP1 – VAR"
        // se preferisci più esplicito:
        // String labelVAR = key + " – Q/P";

        byArticle.put(key, p);

        tabs.addTab(labelPOS, p.posRoot);
        tabs.addTab(labelVAR, p.varRoot);

        lastActiveArticleKey = key;

        return p;
    }

    private String normalizeKey(String articleKey) {
        if (articleKey == null || articleKey.trim().isEmpty()) return "Risultati";
        return articleKey.trim();
    }

    private JComponent buildScrollableRoot(JComponent inner) {
        JScrollPane sp = new JScrollPane(inner,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        sp.setBorder(null);
        sp.getVerticalScrollBar().setUnitIncrement(16);
        sp.getHorizontalScrollBar().setUnitIncrement(16);
        return sp;
    }

    private JPanel wrapWithTitle(String title, JComponent content) {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(BorderFactory.createTitledBorder(title));
        p.add(content, BorderLayout.CENTER);
        return p;
    }

    private void setupChartPanel(ChartPanel cp) {
        // più grande del viewport => compaiono le scroll se lo schermo è piccolo
        cp.setPreferredSize(new Dimension(1400, 650));

        // niente zoom / menu
        cp.setMouseWheelEnabled(false);
        cp.setMouseZoomable(false);
        cp.setDomainZoomable(false);
        cp.setRangeZoomable(false);
        cp.setPopupMenu(null);

        // look pulito
        cp.setOpaque(true);
        cp.setBackground(Color.WHITE);
        cp.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));

        // evita limiti strani in resize
        cp.setMinimumDrawWidth(0);
        cp.setMinimumDrawHeight(0);
    }

    private static class PairTabs {
        String articleKey;

        JComponent posRoot;
        JComponent varRoot;

        ChartPanel posPanel;
        ChartPanel varPanel;
    }
}
