import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;

public class ExcelReader {


    public void readFile(String file) throws IOException {
        Workbook workbook = WorkbookFactory.create(new File(file));
        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        CellFormatter dataFormatter = new CellFormatter();

        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                dataFormatter.printCellValue(cell);
            }
            System.out.println();
        }
    }
}
