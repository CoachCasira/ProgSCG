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
 * In più: manteniamo uno snapshot "base" (temp) preso al caricamento,
 * così il reset NON dipende dal file originale (che può essere lockato).
 */
public class ExcelRepository {

    private static final Logger log = LogManager.getLogger(ExcelRepository.class);

    private File originalFile;
    private File baseSnapshotFile;   // ✅ snapshot base (immutabile)
    private File workingCopyFile;    // working copy modificabile

    public File createWorkingCopy(File original) throws IOException {
        this.originalFile = original;

        // 1) creo snapshot base (una volta) copiando l'originale
        baseSnapshotFile = File.createTempFile("budget_base_", ".xlsx");
        copyFileRobust(originalFile, baseSnapshotFile);

        if (baseSnapshotFile.length() == 0) {
            throw new IOException("Snapshot base creato ma vuoto (0 bytes).");
        }

        // 2) creo working copy e la inizializzo dallo snapshot base
        workingCopyFile = File.createTempFile("budget_work_", ".xlsx");
        copyFileRobust(baseSnapshotFile, workingCopyFile);

        log.info("Snapshot base: {}", baseSnapshotFile.getAbsolutePath());
        log.info("Working copy creata: {} ({} bytes)",
                workingCopyFile.getAbsolutePath(), workingCopyFile.length());

        if (workingCopyFile.length() == 0) {
            throw new IOException("Working copy creata ma vuota (0 bytes).");
        }

        return workingCopyFile;
    }

    /**
     * ✅ RESET: ripristina la working copy allo snapshot base (stato iniziale sessione).
     * Nota: se la working copy è aperta in Excel, Windows potrebbe bloccare la scrittura.
     */
    public File resetWorkingCopyToBase() throws IOException {
        if (workingCopyFile == null) throw new IllegalStateException("Working copy non creata.");
        if (baseSnapshotFile == null) throw new IllegalStateException("Snapshot base non creato.");

        copyFileRobust(baseSnapshotFile, workingCopyFile);

        log.info("RESET completato: working copy ripristinata dallo snapshot base. Size={} bytes",
                workingCopyFile.length());

        if (workingCopyFile.length() == 0) {
            throw new IOException("Reset fallito: working copy vuota (0 bytes).");
        }

        return workingCopyFile;
    }

    /**
     * Salvataggio sicuro: scrivo su file temp, poi rimpiazzo la working copy.
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

    private void copyFileRobust(File src, File dst) throws IOException {
        try (InputStream in = new BufferedInputStream(new FileInputStream(src));
             OutputStream out = new BufferedOutputStream(new FileOutputStream(dst))) {

            byte[] buf = new byte[8192];
            int len;
            while ((len = in.read(buf)) > 0) {
                out.write(buf, 0, len);
            }
            out.flush();
        }
    }

    public void cleanup() {
        if (workingCopyFile != null && workingCopyFile.exists()) {
            boolean ok = workingCopyFile.delete();
            if (ok) log.info("Working copy eliminata.");
            else log.warn("Impossibile eliminare working copy: {}", workingCopyFile.getAbsolutePath());
        }
        if (baseSnapshotFile != null && baseSnapshotFile.exists()) {
            boolean ok = baseSnapshotFile.delete();
            if (ok) log.info("Snapshot base eliminato.");
            else log.warn("Impossibile eliminare snapshot base: {}", baseSnapshotFile.getAbsolutePath());
        }
    }

    public File getOriginalFile() { return originalFile; }
    public File getWorkingCopyFile() { return workingCopyFile; }
    public File getBaseSnapshotFile() { return baseSnapshotFile; }
}
