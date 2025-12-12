package repository;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Gestisce l'accesso al file Excel.
 * Scelta progettuale: NON tocchiamo mai l'originale. Creiamo sempre una working copy.
 *
 * Nota importante: quando salviamo un workbook, scriviamo su un file temporaneo e poi facciamo replace,
 * così non rischiamo mai di lasciare il file a 0 byte in caso di crash/errore.
 */
public class ExcelRepository {

    private static final Logger log = LogManager.getLogger(ExcelRepository.class);

    private File originalFile;
    private File workingCopyFile;

    public File createWorkingCopy(File original) throws IOException {
        this.originalFile = original;

        workingCopyFile = File.createTempFile("budget_work_", ".xlsx");

        // Copia robusta via stream (evita problemi su alcune macchine/lock)
        try (InputStream in = new BufferedInputStream(new FileInputStream(original));
             OutputStream out = new BufferedOutputStream(new FileOutputStream(workingCopyFile))) {

            byte[] buf = new byte[8192];
            int len;
            while ((len = in.read(buf)) > 0) {
                out.write(buf, 0, len);
            }
            out.flush();
        }

        log.info("Working copy creata: {} ({} bytes)",
                workingCopyFile.getAbsolutePath(), workingCopyFile.length());

        if (workingCopyFile.length() == 0) {
            throw new IOException("Working copy creata ma vuota (0 bytes).");
        }

        return workingCopyFile;
    }

    /**
     * Salvataggio sicuro: scrivo su file temp, poi rimpiazzo la working copy.
     * Così non rischio mai di corrompere il file.
     */
    public void safeSaveWorkbook(Workbook wb) throws IOException {
        if (workingCopyFile == null) {
            throw new IllegalStateException("Working copy non creata.");
        }

        File tmp = File.createTempFile("budget_save_", ".xlsx");
        log.debug("Salvataggio sicuro su temp: {}", tmp.getAbsolutePath());

        try (OutputStream out = new BufferedOutputStream(new FileOutputStream(tmp))) {
            wb.write(out);
            out.flush();
        }

        if (tmp.length() == 0) {
            throw new IOException("Salvataggio fallito: file temp vuoto (0 bytes).");
        }

        Files.move(tmp.toPath(), workingCopyFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        log.info("Working copy aggiornata (safe save). Size={} bytes", workingCopyFile.length());
    }

    public void cleanup() {
        if (workingCopyFile != null && workingCopyFile.exists()) {
            boolean ok = workingCopyFile.delete();
            if (ok) log.info("Working copy eliminata. Reset completato.");
            else log.warn("Impossibile eliminare working copy: {}", workingCopyFile.getAbsolutePath());
        }
    }

    public File getOriginalFile() { return originalFile; }
    public File getWorkingCopyFile() { return workingCopyFile; }
}
