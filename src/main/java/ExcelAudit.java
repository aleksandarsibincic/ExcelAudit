import java.io.IOException;

public class ExcelAudit {

    public static final String SAMPLE_XLSX_FILE_PATH = "./sample_file.xlsx";

    public static void main(String[] args) throws IOException {

        ExcelReader excelReader = new ExcelReader();
        excelReader.readFile(SAMPLE_XLSX_FILE_PATH);

    }
}
