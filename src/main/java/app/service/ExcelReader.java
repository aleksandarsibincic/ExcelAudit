package app.service;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

    private Workbook workbook;

    public ExcelReader(File file) throws IOException {
        workbook = WorkbookFactory.create(file);
    }

    public Workbook getWorkbook(){
        return workbook;
    }

    public List<Sheet> getSheets(){
        List<Sheet> sheets = new ArrayList<>();
        workbook.forEach(sheet -> {
            sheets.add(sheet);
        });
        return sheets;
    }

    public List<Cell> getCells(Sheet sheet){
        List<Cell> cells = new ArrayList<>();
        for (Row row: sheet) {
            for(Cell cell: row) {
                cells.add(cell);
            }
        }
        return cells;
    }
}
