package main;

import controller.MainController;
import model.AppModel;
import repository.ExcelRepository;
import view.MainFrame;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import javax.swing.SwingUtilities;

public class MainApp {

    private static final Logger log = LogManager.getLogger(MainApp.class);

    public static void main(String[] args) {
        log.info("Avvio applicazione...");

        SwingUtilities.invokeLater(() -> {
            AppModel model = new AppModel();
            MainFrame view = new MainFrame();
            ExcelRepository repo = new ExcelRepository();

            new MainController(model, view, repo);

            view.setVisible(true);
            log.info("GUI mostrata.");
        });
    }
}
