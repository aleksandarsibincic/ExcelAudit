package model;

import java.util.List;

public class Workbook {

    private String filename;
    private List<Spreadsheet> hasSheets;

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public List<Spreadsheet> getHasSheets() {
        return hasSheets;
    }

    public void setHasSheets(List<Spreadsheet> hasSheets) {
        this.hasSheets = hasSheets;
    }
}
