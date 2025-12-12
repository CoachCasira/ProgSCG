package model;

import java.io.File;

/**
 * Model dell'app: contiene lo stato condiviso.
 * Per ora: riferimenti ai file (originale e copia).
 * In seguito: lista articoli, parametri, valori calcolati, ecc.
 */
public class AppModel {

    private File originalExcel;
    private File workingExcelCopy;

    public File getOriginalExcel() { return originalExcel; }
    public void setOriginalExcel(File originalExcel) { this.originalExcel = originalExcel; }

    public File getWorkingExcelCopy() { return workingExcelCopy; }
    public void setWorkingExcelCopy(File workingExcelCopy) { this.workingExcelCopy = workingExcelCopy; }
}
