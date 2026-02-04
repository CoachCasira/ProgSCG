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
    // CE charts: NO-OP
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
     * Tab con 2 grafici verticali.
     * Ogni grafico: JScrollPane con H+V AS_NEEDED.
     */
    private JComponent buildTabUI(ArticleTab tab) {
        JPanel inner = new JPanel();
        inner.setLayout(new BoxLayout(inner, BoxLayout.Y_AXIS));
        inner.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        tab.posChartPanel = new ScrollableChartPanel(null);
        tab.compChartPanel = new ScrollableChartPanel(null);

        setupChartPanel(tab.posChartPanel);
        setupChartPanel(tab.compChartPanel);

        inner.add(wrapChartWithTitleAndScroll("POS", tab.posChartPanel));
        inner.add(Box.createVerticalStrut(10));
        inner.add(wrapChartWithTitleAndScroll("Variazione (Quantità / Prezzo)", tab.compChartPanel));

        JScrollPane sp = new JScrollPane(inner,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        sp.setBorder(null);
        sp.getVerticalScrollBar().setUnitIncrement(16);

        return sp;
    }

    /**
     * Wrapper: titolo + scroll X/Y sul grafico.
     */
    private JComponent wrapChartWithTitleAndScroll(String title, ChartPanel chartPanel) {
        JPanel outer = new JPanel(new BorderLayout());
        outer.setBorder(BorderFactory.createTitledBorder(title));

        JScrollPane scroll = new JScrollPane(chartPanel,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);

        // scorrimento più fluido
        scroll.getHorizontalScrollBar().setUnitIncrement(20);
        scroll.getVerticalScrollBar().setUnitIncrement(16);

        // padding interno pulito
        scroll.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));

        outer.add(scroll, BorderLayout.CENTER);
        return outer;
    }

    private void setupChartPanel(ChartPanel cp) {
        // ✅ più grande della viewport -> scrollbar sempre disponibili su laptop
        cp.setPreferredSize(new Dimension(1600, 520));
        cp.setMinimumSize(new Dimension(1200, 420));

        // niente zoom / menu
        cp.setMouseWheelEnabled(false);
        cp.setMouseZoomable(false);
        cp.setDomainZoomable(false);
        cp.setRangeZoomable(false);
        cp.setPopupMenu(null);

        cp.setOpaque(true);
        cp.setBackground(Color.WHITE);

        cp.setMinimumDrawWidth(0);
        cp.setMinimumDrawHeight(0);
    }

    /**
     * 🔥 Fix vero: ChartPanel “scrollabile” che NON si adatta alla viewport.
     * Se tracka la viewport -> le scrollbar non compaiono mai.
     */
    private static class ScrollableChartPanel extends ChartPanel implements Scrollable {
        public ScrollableChartPanel(JFreeChart chart) {
            super(chart, false);
        }

        @Override
        public Dimension getPreferredScrollableViewportSize() {
            return new Dimension(900, 320);
        }

        @Override
        public int getScrollableUnitIncrement(Rectangle visibleRect, int orientation, int direction) {
            return (orientation == SwingConstants.HORIZONTAL) ? 20 : 16;
        }

        @Override
        public int getScrollableBlockIncrement(Rectangle visibleRect, int orientation, int direction) {
            return (orientation == SwingConstants.HORIZONTAL) ? 200 : 120;
        }

        @Override
        public boolean getScrollableTracksViewportWidth() {
            return false; // ✅ fondamentale
        }

        @Override
        public boolean getScrollableTracksViewportHeight() {
            return false; // ✅ fondamentale
        }
    }

    private static class ArticleTab {
        JComponent root;
        ChartPanel posChartPanel;
        ChartPanel compChartPanel;
    }
}
