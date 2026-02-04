package view;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;

import javax.swing.*;
import java.awt.*;
import java.util.LinkedHashMap;
import java.util.Map;

public class ChartsPanel extends JPanel {

    /** Tab per articolo (MP||MP1, MP||MP1 - VAR, ...). */
    private final JTabbedPane tabs = new JTabbedPane();

    /** Mantiene i pannelli per key tab. */
    private final Map<String, TabState> byKey = new LinkedHashMap<>();

    /** Ultima tab attiva (fallback per API vecchie senza key). */
    private String lastActiveKey = null;

    // --- tuning “zoom”/scroll: più piccoli = meno scroll su schermi normali
    private static final Dimension CHART_PREF = new Dimension(920, 520);

    public ChartsPanel() {
        super(new BorderLayout());
        tabs.setBorder(BorderFactory.createTitledBorder("Grafici"));
        add(tabs, BorderLayout.CENTER);
    }

    // ============================================================
    // API "nuova" (articolo esplicito)
    // ============================================================

    /** Imposta il grafico POS nella tab "<articleKey>" */
    public void setPosChart(String articleKey, JFreeChart chart) {
        final String key = normalizeKey(articleKey);
        SwingUtilities.invokeLater(() -> {
            TabState t = getOrCreateTab(key, "POS");
            t.chartPanel.setChart(chart);
            t.chartPanel.revalidate();
            t.chartPanel.repaint();
            setActiveTab(key);
        });
    }

    /**
     * Imposta il grafico VAR nella tab "<articleKey> - VAR"
     * (variazione quantità/prezzo)
     */
    public void setCompChart(String articleKey, JFreeChart chart) {
        final String base = normalizeKey(articleKey);
        final String key = base + " - VAR";
        SwingUtilities.invokeLater(() -> {
            TabState t = getOrCreateTab(key, "Variazione (Quantità / Prezzo)");
            t.chartPanel.setChart(chart);
            t.chartPanel.revalidate();
            t.chartPanel.repaint();
            // NON cambio tab automaticamente se vuoi restare su POS:
            // se invece vuoi andare sulla VAR appena simuli, scommenta:
            // setActiveTab(key);
        });
    }

    public void setActiveArticleTab(String articleKey) {
        final String key = normalizeKey(articleKey);
        SwingUtilities.invokeLater(() -> setActiveTab(key));
    }

    // ============================================================
    // API "vecchia" (compatibilità) usata dal MainController
    // ============================================================

    /**
     * Vecchia API: aggiunge/aggiorna la tab articolo e setta i due grafici.
     * - labelTab: es. "MP||MP1"
     * - subtitle: ignorato
     */
    public void addOrReplaceArticleCharts(String labelTab, String subtitle, JFreeChart posChart, JFreeChart compChart) {
        // POS su tab base
        setPosChart(labelTab, posChart);
        // VAR su tab separata "<base> - VAR"
        setCompChart(labelTab, compChart);
    }

    /** Vecchia API: pulisce tutte le tab articolo. */
    public void clearArticleCharts() {
        clearAll();
    }

    /** Vecchia API: setPosChart “globale” (senza key articolo). */
    public void setPosChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            TabState t = getOrCreateActiveTab();
            t.chartPanel.setChart(chart);
            t.chartPanel.revalidate();
            t.chartPanel.repaint();
        });
    }

    /** Vecchia API: setCompChart “globale” (senza key articolo). */
    public void setCompChart(JFreeChart chart) {
        SwingUtilities.invokeLater(() -> {
            TabState t = getOrCreateActiveTab();
            t.chartPanel.setChart(chart);
            t.chartPanel.revalidate();
            t.chartPanel.repaint();
        });
    }

    // ============================================================
    // CE charts: non li vuoi più qui -> NO-OP (ma compila col controller)
    // ============================================================

    public void setCeBaseChart(JFreeChart chart) { /* NO-OP */ }
    public void setCeVarChart(JFreeChart chart)  { /* NO-OP */ }
    public void setCeDeltaChart(JFreeChart chart){ /* NO-OP */ }

    // ============================================================
    // Utility
    // ============================================================

    public void clearAll() {
        SwingUtilities.invokeLater(() -> {
            byKey.clear();
            tabs.removeAll();
            lastActiveKey = null;
            revalidate();
            repaint();
        });
    }

    private void setActiveTab(String key) {
        TabState t = byKey.get(key);
        if (t == null) return;
        int idx = tabs.indexOfComponent(t.root);
        if (idx >= 0) tabs.setSelectedIndex(idx);
        lastActiveKey = key;
    }

    private TabState getOrCreateActiveTab() {
        // 1) tab selezionata
        int sel = tabs.getSelectedIndex();
        if (sel >= 0) {
            Component comp = tabs.getComponentAt(sel);
            for (Map.Entry<String, TabState> e : byKey.entrySet()) {
                if (e.getValue().root == comp) {
                    lastActiveKey = e.getKey();
                    return e.getValue();
                }
            }
        }
        // 2) ultima attiva
        if (lastActiveKey != null && byKey.containsKey(lastActiveKey)) {
            return byKey.get(lastActiveKey);
        }
        // 3) fallback
        return getOrCreateTab("Risultati", "POS");
    }

    private TabState getOrCreateTab(String key, String titleBorder) {
        key = normalizeKey(key);

        TabState existing = byKey.get(key);
        if (existing != null) {
            // ✅ difensivo: se per qualche motivo chartPanel è null, ricreo UI
            if (existing.chartPanel == null || existing.root == null) {
                TabState rebuilt = buildTab(key, titleBorder);
                byKey.put(key, rebuilt);

                int oldIdx = tabs.indexOfTab(key);
                if (oldIdx >= 0) tabs.setComponentAt(oldIdx, rebuilt.root);
                else tabs.addTab(key, rebuilt.root);

                lastActiveKey = key;
                return rebuilt;
            }
            lastActiveKey = key;
            return existing;
        }

        TabState created = buildTab(key, titleBorder);
        byKey.put(key, created);
        tabs.addTab(key, created.root);

        lastActiveKey = key;
        return created;
    }

    private TabState buildTab(String key, String titleBorder) {
        TabState t = new TabState();

        // ChartPanel
        ChartPanel cp = new ChartPanel(null, false);
        setupChartPanel(cp);
        t.chartPanel = cp;

        // ScrollPane che permette sia orizzontale sia verticale “as needed”
        JScrollPane sp = new JScrollPane(cp,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        sp.getVerticalScrollBar().setUnitIncrement(16);
        sp.getHorizontalScrollBar().setUnitIncrement(16);

        JPanel root = new JPanel(new BorderLayout());
        root.setBorder(BorderFactory.createTitledBorder(titleBorder));
        root.add(sp, BorderLayout.CENTER);

        t.root = root;
        return t;
    }

    private void setupChartPanel(ChartPanel cp) {
        // “meno zoomato”: dimensione preferita più contenuta (ma scroll su schermi piccoli)
        cp.setPreferredSize(CHART_PREF);

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

    private String normalizeKey(String k) {
        if (k == null) return "Risultati";
        k = k.trim();
        return k.isEmpty() ? "Risultati" : k;
    }

    private static class TabState {
        JComponent root;
        ChartPanel chartPanel;
    }
}
