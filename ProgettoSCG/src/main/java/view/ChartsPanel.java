package view;

import javax.swing.*;
import java.awt.*;
import java.util.LinkedHashMap;
import java.util.Map;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

public class ChartsPanel extends JPanel {

    /** Tab per articolo (MP1, MP2, ...). */
    private final JTabbedPane articleTabs = new JTabbedPane();

    /** Mantiene i pannelli per articolo. */
    private final Map<String, ArticleTab> byArticle = new LinkedHashMap<>();

    /** Ultimo articolo “attivo” (fallback per API vecchie). */
    private String lastActiveArticleKey = null;

    public ChartsPanel() {
        super(new BorderLayout());
        articleTabs.setBorder(BorderFactory.createTitledBorder("Grafici"));
        add(articleTabs, BorderLayout.CENTER);
    }

    // ============================================================
    // API "nuova" (articolo esplicito)
    // ============================================================

    public void setPosChart(String articleKey, JFreeChart chart) {
        final String key = normalizeKey(articleKey);
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateTab(key);
            tab.posChartPanel.setChart(chart);
            tab.posChartPanel.revalidate();
            tab.posChartPanel.repaint();
        });
    }

    public void setCompChart(String articleKey, JFreeChart chart) {
        final String key = normalizeKey(articleKey);
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateTab(key);
            tab.compChartPanel.setChart(chart);
            tab.compChartPanel.revalidate();
            tab.compChartPanel.repaint();
        });
    }

    public void setActiveArticleTab(String articleKey) {
        final String key = normalizeKey(articleKey);
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateTab(key);
            int idx = articleTabs.indexOfComponent(tab.root);
            if (idx >= 0) articleTabs.setSelectedIndex(idx);
            lastActiveArticleKey = key;
        });
    }

    // ============================================================
    // API "vecchia" (compatibilità)
    // ============================================================

    public void addOrReplaceArticleCharts(String labelTab, String subtitle, JFreeChart posChart, JFreeChart compChart) {
        setPosChart(labelTab, posChart);
        setCompChart(labelTab, compChart);
        setActiveArticleTab(labelTab);
    }

    public void clearArticleCharts() {
        clearAll();
    }

    public void setPosChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateActiveTab();
            tab.posChartPanel.setChart(chart);
            tab.posChartPanel.revalidate();
            tab.posChartPanel.repaint();
        });
    }

    public void setCompChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateActiveTab();
            tab.compChartPanel.setChart(chart);
            tab.compChartPanel.revalidate();
            tab.compChartPanel.repaint();
        });
    }

    // ============================================================
    // CE charts: NO-OP (non vuoi la tab CE qui)
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
            articleTabs.removeAll();
            lastActiveArticleKey = null;
            revalidate();
            repaint();
        });
    }

    private ArticleTab getOrCreateActiveTab() {
        int sel = articleTabs.getSelectedIndex();
        if (sel >= 0) {
            Component comp = articleTabs.getComponentAt(sel);
            String key = findKeyByComponent(comp);
            if (key != null) {
                lastActiveArticleKey = key;
                return byArticle.get(key);
            }
        }
        if (lastActiveArticleKey != null && byArticle.containsKey(lastActiveArticleKey)) {
            return byArticle.get(lastActiveArticleKey);
        }
        return getOrCreateTab("Risultati");
    }

    private String findKeyByComponent(Component comp) {
        for (Map.Entry<String, ArticleTab> e : byArticle.entrySet()) {
            if (e.getValue().root == comp) return e.getKey();
        }
        return null;
    }

    private ArticleTab getOrCreateTab(String articleKey) {
        String key = normalizeKey(articleKey);

        ArticleTab existing = byArticle.get(key);
        if (existing != null) {
            lastActiveArticleKey = key;
            return existing;
        }

        ArticleTab tab = new ArticleTab();
        tab.root = buildTabUI(tab);

        byArticle.put(key, tab);
        articleTabs.addTab(key, tab.root);

        lastActiveArticleKey = key;
        return tab;
    }

    private String normalizeKey(String k) {
        if (k == null) return "Risultati";
        String t = k.trim();
        return t.isEmpty() ? "Risultati" : t;
    }

    /**
     * UI tab: due grafici in verticale.
     * Ogni grafico è dentro uno JScrollPane con H-scroll AS_NEEDED
     * così su schermi piccoli puoi scorrere sx↔dx.
     */
    private JComponent buildTabUI(ArticleTab tab) {
        JPanel inner = new JPanel();
        inner.setLayout(new BoxLayout(inner, BoxLayout.Y_AXIS));
        inner.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        tab.posChartPanel = new ChartPanel(null, false);
        tab.compChartPanel = new ChartPanel(null, false);

        setupChartPanel(tab.posChartPanel);
        setupChartPanel(tab.compChartPanel);

        inner.add(wrapChartWithTitleAndHScroll("POS", tab.posChartPanel));
        inner.add(Box.createVerticalStrut(10));
        inner.add(wrapChartWithTitleAndHScroll("Variazione (Quantità / Prezzo)", tab.compChartPanel));

        JScrollPane sp = new JScrollPane(inner,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        sp.setBorder(null);
        sp.getVerticalScrollBar().setUnitIncrement(16);

        return sp;
    }

    /**
     * Wrapper con titolo + scroll ORIZZONTALE sul grafico.
     * Nota: per far comparire la scrollbar, il ChartPanel deve avere preferredWidth
     * maggiore della viewport (lo facciamo in setupChartPanel).
     */
    private JComponent wrapChartWithTitleAndHScroll(String title, ChartPanel chartPanel) {
        JPanel outer = new JPanel(new BorderLayout());
        outer.setBorder(BorderFactory.createTitledBorder(title));

        JScrollPane hScroll = new JScrollPane(chartPanel,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_NEVER,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);

        // scorrimento più “fluido”
        hScroll.getHorizontalScrollBar().setUnitIncrement(20);

        // evita bordo doppio
        hScroll.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));

        outer.add(hScroll, BorderLayout.CENTER);
        return outer;
    }

    private void setupChartPanel(ChartPanel cp) {
        // ✅ width più grande della viewport -> appare scroll orizzontale su laptop
        cp.setPreferredSize(new Dimension(1400, 320));

        // niente zoom / menu
        cp.setMouseWheelEnabled(false);
        cp.setMouseZoomable(false);
        cp.setDomainZoomable(false);
        cp.setRangeZoomable(false);
        cp.setPopupMenu(null);

        // look pulito
        cp.setOpaque(true);
        cp.setBackground(Color.WHITE);

        // evita limiti strani in resize
        cp.setMinimumDrawWidth(0);
        cp.setMinimumDrawHeight(0);
    }

    private static class ArticleTab {
        JComponent root;
        ChartPanel posChartPanel;
        ChartPanel compChartPanel;
    }
}