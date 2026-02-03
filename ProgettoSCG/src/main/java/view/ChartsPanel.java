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
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateTab(articleKey);
            tab.posChartPanel.setChart(chart);
            tab.posChartPanel.revalidate();
            tab.posChartPanel.repaint();
        });
    }

    public void setCompChart(String articleKey, JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateTab(articleKey);
            tab.compChartPanel.setChart(chart);
            tab.compChartPanel.revalidate();
            tab.compChartPanel.repaint();
        });
    }

    public void setActiveArticleTab(String articleKey) {
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateTab(articleKey);
            int idx = articleTabs.indexOfComponent(tab.root);
            if (idx >= 0) articleTabs.setSelectedIndex(idx);
            lastActiveArticleKey = articleKey;
        });
    }

    // ============================================================
    // API "vecchia" (compatibilità) usata dal tuo MainController
    // ============================================================

    /**
     * Vecchia API: aggiunge/aggiorna la tab articolo e setta i due grafici.
     * - labelTab: spesso è MP1, MP2, ecc. (noi lo usiamo come key)
     * - subtitle: ignorato (titoli li metti già nei chart)
     */
    public void addOrReplaceArticleCharts(String labelTab, String subtitle, JFreeChart posChart, JFreeChart compChart) {
        // subtitle non serve più
        setPosChart(labelTab, posChart);
        setCompChart(labelTab, compChart);
        setActiveArticleTab(labelTab);
    }

    /** Vecchia API: pulisce tutte le tab articolo. */
    public void clearArticleCharts() {
        clearAll();
    }

    /**
     * Vecchia API: setPosChart “globale” (senza key articolo).
     * La applichiamo alla tab selezionata / ultima attiva.
     */
    public void setPosChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateActiveTab();
            tab.posChartPanel.setChart(chart);
            tab.posChartPanel.revalidate();
            tab.posChartPanel.repaint();
        });
    }

    /** Vecchia API: setCompChart “globale” (senza key articolo). */
    public void setCompChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            ArticleTab tab = getOrCreateActiveTab();
            tab.compChartPanel.setChart(chart);
            tab.compChartPanel.revalidate();
            tab.compChartPanel.repaint();
        });
    }

    // ============================================================
    // CE charts: il MainController li chiama, ma tu NON vuoi più la tab CE qui.
    // Quindi li lasciamo come NO-OP per compilare senza cambiare logica.
    // ============================================================

    public void setCeBaseChart(JFreeChart chart) {
        // NO-OP: CE Budget lo apri con pulsante dedicato, non lo mostriamo nei "Grafici"
    }

    public void setCeVarChart(JFreeChart chart) {
        // NO-OP
    }

    public void setCeDeltaChart(JFreeChart chart) {
        // NO-OP
    }

    // ============================================================
    // Utility
    // ============================================================

    /** Rimuove tutte le tab (articoli). */
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
        // 1) tab selezionata
        int sel = articleTabs.getSelectedIndex();
        if (sel >= 0) {
            Component comp = articleTabs.getComponentAt(sel);
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
        return getOrCreateTab("Risultati");
    }

    private String findKeyByComponent(Component comp) {
        for (Map.Entry<String, ArticleTab> e : byArticle.entrySet()) {
            if (e.getValue().root == comp) return e.getKey();
        }
        return null;
    }

    private ArticleTab getOrCreateTab(String articleKey) {
        if (articleKey == null || articleKey.trim().isEmpty()) articleKey = "Risultati";
        articleKey = articleKey.trim();

        ArticleTab existing = byArticle.get(articleKey);
        if (existing != null) {
            lastActiveArticleKey = articleKey;
            return existing;
        }

        ArticleTab tab = new ArticleTab();
        tab.root = buildTabUI(tab);

        byArticle.put(articleKey, tab);
        articleTabs.addTab(articleKey, tab.root);

        lastActiveArticleKey = articleKey;
        return tab;
    }

    private JComponent buildTabUI(ArticleTab tab) {
        // contenuto “scrollabile”: due grafici in verticale
        JPanel inner = new JPanel();
        inner.setLayout(new GridLayout(2, 1, 0, 10));
        inner.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        tab.posChartPanel = new ChartPanel(null, false);
        tab.compChartPanel = new ChartPanel(null, false);

        setupChartPanel(tab.posChartPanel);
        setupChartPanel(tab.compChartPanel);

        JPanel posWrap = wrapWithTitle("POS", tab.posChartPanel);
        JPanel compWrap = wrapWithTitle("Variazione (Quantità / Prezzo)", tab.compChartPanel);

        inner.add(posWrap);
        inner.add(compWrap);

        JScrollPane sp = new JScrollPane(inner,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        sp.setBorder(null);
        sp.getVerticalScrollBar().setUnitIncrement(16);

        return sp;
    }

    private JPanel wrapWithTitle(String title, JComponent content) {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(BorderFactory.createTitledBorder(title));
        p.add(content, BorderLayout.CENTER);
        return p;
    }

    private void setupChartPanel(ChartPanel cp) {
        cp.setPreferredSize(new Dimension(1000, 320));

        // niente zoom
        cp.setMouseWheelEnabled(false);
        cp.setMouseZoomable(false);
        cp.setDomainZoomable(false);
        cp.setRangeZoomable(false);

        // niente menu popup
        cp.setPopupMenu(null);

        // look pulito
        cp.setOpaque(true);
        cp.setBackground(Color.WHITE);
        cp.setBorder(BorderFactory.createEmptyBorder(6, 6, 6, 6));

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
