import org.apache.jena.ontology.OntModelSpec;
import org.apache.jena.rdf.model.Model;
import org.apache.jena.rdf.model.ModelFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class ExcelAudit {

    public static final String SAMPLE_XLSX_FILE_PATH = "./sample_file.xlsx";
    public static final String ONTOLOGY_PATH = "./excel_ontology.owl" ;

    public static void main(String[] args) throws IOException {


        ExcelReader excelReader = new ExcelReader(SAMPLE_XLSX_FILE_PATH);
        Workbook workbook = excelReader.getWorkbook();
        System.out.println("Workbook named " + new File(SAMPLE_XLSX_FILE_PATH).getName());

        List<Sheet> sheets = excelReader.getSheets();
        System.out.println("Workbook has " + sheets.size() + " Sheets : ");
        sheets.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        List<Cell> cells = excelReader.getCells(sheets.get(0));
        CellFormatter dataFormatter = new CellFormatter();
        cells.forEach(cell -> {
            System.out.print(cell.getAddress().toString()+" ");
            dataFormatter.printCellValue(cell);
            System.out.println();
        });

        Model model = ModelFactory.createOntologyModel(OntModelSpec.OWL_DL_MEM);
        File ontologyFile = new File(ONTOLOGY_PATH);
        model.read(ontologyFile.toURL().toString());
        model.write(System.out);

    }
}
