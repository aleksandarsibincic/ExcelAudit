package app.service;

import com.ontotext.trree.OwlimSchemaRepository;
import org.apache.jena.base.Sys;
import org.apache.jena.ontology.*;
import org.apache.jena.query.*;
import org.apache.jena.rdf.model.Model;
import org.apache.jena.rdf.model.ModelFactory;
import org.apache.jena.rdf.model.Property;
import org.apache.jena.tdb.TDB;
import org.apache.jena.tdb.TDBFactory;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.rdf4j.model.impl.TreeModel;
import org.eclipse.rdf4j.repository.RepositoryConnection;
import org.eclipse.rdf4j.repository.manager.LocalRepositoryManager;
import org.eclipse.rdf4j.repository.manager.RepositoryManager;
import org.eclipse.rdf4j.repository.sail.SailRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class ExcelAuditService {

//    public static final String SAMPLE_XLSX_FILE_PATH = "./sample_file.xlsx";
    public static final String ONTOLOGY_PATH = "./excel_ontology.owl";
    public static final String NAMESPACE = "http://www.semanticweb.org/ds/ontologies/2020/5/excel-audit#";


    public void processFile(byte[] file1) throws IOException, InvalidFormatException {
        /* Create temporary file from bytes */
        File file = new File("text.xlsx");
        OutputStream os = new FileOutputStream(file);
        os.write(file1);
        os.close();

        ExcelReader excelReader = new ExcelReader(file);
        List<Sheet> sheets = excelReader.getSheets();

        XSSFWorkbook wb = new XSSFWorkbook(file);
        XSSFEvaluationWorkbook xssfew = XSSFEvaluationWorkbook.create(wb);

        /* Create Ontology model */
        OntModel model = ModelFactory.createOntologyModel(OntModelSpec.OWL_DL_MEM);
        File ontologyFile = new File(ONTOLOGY_PATH);
        model.read(ontologyFile.toURL().toString());

        /* adding data from excel to ontology model */

        /*create workbook class*/
        OntClass workbookClass = model.getOntClass(NAMESPACE + "Workbook");
        /*create workbook individual*/
        Individual workbookIndividual = workbookClass.createIndividual(NAMESPACE + file.getName());
        /*add filename property to the workbook*/
        DatatypeProperty filename = model.getDatatypeProperty(NAMESPACE + "filename");
        workbookIndividual.addProperty(filename, file.getName());

        /*get spreadsheet class*/
        OntClass spreadsheetClass = model.getOntClass(NAMESPACE + "Spreadsheet");
        /* get hasSheet relation/object property */
        ObjectProperty hasSheets = model.getObjectProperty(NAMESPACE + "hasSheets");
        /* get cell class */
        OntClass cellClass = model.getOntClass(NAMESPACE + "Cell");
        /* get hasCell relation/object property */
        ObjectProperty hasCell = model.getObjectProperty(NAMESPACE + "hasCell");
        /*get cell value property*/
        DatatypeProperty value = model.getDatatypeProperty(NAMESPACE + "value");
        /*get cell identifier*/
        DatatypeProperty identifier = model.getDatatypeProperty(NAMESPACE + "identifier");
        /*get cell relations*/
        ObjectProperty in = model.getObjectProperty(NAMESPACE + "in");
        sheets.forEach(sheet -> {
            /* create spreadsheet individual*/
            Individual spreadsheetIndividual = spreadsheetClass.createIndividual(NAMESPACE + sheet.getSheetName());
            /* add relations between workbook and sheet*/
            model.add(workbookIndividual, hasSheets, spreadsheetIndividual);
            /* get all cells from sheet */
            List<Cell> sheetCells = excelReader.getCells(sheet);
            sheetCells.forEach(cell -> {
                if (cell.getCellTypeEnum().equals(CellType.NUMERIC) && !DateUtil.isCellDateFormatted(cell)) {
                    /* create cell individual*/
                    Individual cellIndividual = cellClass.createIndividual(NAMESPACE + cell.getAddress().toString());
                    cellIndividual.addProperty(identifier, cell.getAddress().toString());
                    /* add relations between sheet and cells*/
                    model.add(spreadsheetIndividual, hasCell, cellIndividual);
                    //add numerical value to cell
                    double cellValue = cell.getNumericCellValue();
                    cellIndividual.addProperty(value, String.valueOf(cellValue));
                } else if (cell.getCellTypeEnum().equals(CellType.STRING)) {
                    /* create cell individual*/
                    Individual cellIndividual = cellClass.createIndividual(NAMESPACE + cell.getAddress().toString());
                    cellIndividual.addProperty(identifier, cell.getAddress().toString());
                    /* add relations between sheet and cells*/
                    model.add(spreadsheetIndividual, hasCell, cellIndividual);
                    //add numerical value to cell
                    String cellValue = cell.getStringCellValue();
                    cellIndividual.addProperty(value, cellValue);
                }
            });
            sheetCells.forEach(cell -> {
                if (cell.getCellTypeEnum().equals(CellType.FORMULA)) {
                    Individual cellIndividual = cellClass.createIndividual(NAMESPACE + cell.getAddress().toString());
                    /* add relations between sheet and cells*/
                    model.add(spreadsheetIndividual, hasCell, cellIndividual);
                    Ptg[] ptg = FormulaParser.parse(cell.getCellFormula(), xssfew, FormulaType.NAMEDRANGE, sheets.indexOf(sheet), cell.getRowIndex());
                    for (int i = 0; i < ptg.length; i++) {
                        byte cla = ptg[i].getPtgClass();
                        if (cla == Ptg.CLASS_ARRAY) {
                            String formula = ptg[i].toFormulaString();
                            String rangeFormulaPattern = "[A-Z]*\\$\\d*\\:[A-Z]*\\$\\d*";
                            if (formula.matches(rangeFormulaPattern)) {
                                String column = "";
                                String startRow = "";
                                String endRow = formula.substring(formula.lastIndexOf("$") + 1);
                                Pattern columnPattern = Pattern.compile("(.*?)\\$");
                                Matcher columnMatcher = columnPattern.matcher(formula);
                                if (columnMatcher.find()) {
                                    column = columnMatcher.group().replaceAll("\\$", "");
                                }
                                Pattern startRowPattern = Pattern.compile("\\$(.*?)\\:");
                                Matcher startRowMatcher = startRowPattern.matcher(formula);
                                if (startRowMatcher.find()) {
                                    startRow = startRowMatcher.group().replaceAll("\\$", "").replaceAll("\\:", "");
                                }
                                for (int index = Integer.valueOf(startRow); index <= Integer.valueOf(endRow); index++) {
                                    Individual individualCell = model.getIndividual(NAMESPACE + column + index);
                                    if (individualCell != null) {
                                        model.add(individualCell, in, cellIndividual);
                                    }

                                }
                            }
                        }
                        if (cla == Ptg.CLASS_VALUE) {
                            String index = ptg[i].toFormulaString();
                            Individual individualCell = model.getIndividual(NAMESPACE + index);
                            if (individualCell != null) {
                                model.add(individualCell, in, cellIndividual);
                            }
                        }
                    }
                    cellIndividual.addProperty(value, String.valueOf(cell.getNumericCellValue()));
                }
            });
        });
        model.write(System.out);

        // store Ontology model to in memory TDB, which is not persisted over the sessions

        Dataset dataset = TDBFactory.createDataset();
        dataset.begin(ReadWrite.WRITE);
        dataset.addNamedModel(NAMESPACE, model);
        dataset.commit();

        System.out.println("Application is ready");
        Scanner scanner = new Scanner(System.in);

        // extract saved ontology model from TDB and execute simple query

        dataset.begin(ReadWrite.READ);
        Model queryModel = dataset.getNamedModel(NAMESPACE);

        System.out.println("Choose an option a, b or c:");
        System.out.println("a - Find dependent cells");
        System.out.println("b - Find cells with certain value");
        System.out.println("c - Upload a new file");

        while (scanner.hasNextLine()) {
            String queryString = "";
            String line = scanner.nextLine();
            if (line.toLowerCase().equals("a")) {
                System.out.println("Type cell (e.g. B3) to find dependent cells:");
                if (scanner.hasNextLine()) {
                    line = scanner.nextLine();
                    queryString =
                            "PREFIX rdf: <http://www.semanticweb.org/ds/ontologies/2020/5/excel-audit#> " +
                                    "PREFIX a2: <http://www.semanticweb.org/ds/ontologies/2020/5/excel-audit#" + line + "> " +
                                    "SELECT * WHERE { a2: rdf:in ?Cell . }";
                }
            } else if (line.toLowerCase().equals("b")) {
                System.out.println("Type float value (e.g. 4.0) to find cells containing that value:");
                if (scanner.hasNextLine()) {
                    line = scanner.nextLine();
                    queryString =
                            "PREFIX rdf: <http://www.semanticweb.org/ds/ontologies/2020/5/excel-audit#> " +
                                    "SELECT * WHERE { ?Cell rdf:value \"" + line + "\" . }";
                }
            } else if (line.toLowerCase().equals("c")) {
                break;
            }
            if (queryString.equals("")) {
                System.out.println("Non-existing option has been chosen");
            } else {
                Query query = QueryFactory.create(queryString);
                QueryExecution queryExecution = QueryExecutionFactory.create(query, queryModel);

                try {
                    ResultSet resultSet = queryExecution.execSelect();
                    while (resultSet.hasNext()) {
                        QuerySolution solution = resultSet.nextSolution();
                        System.out.println(solution.get("Cell").toString().split("#")[1]);
                    }
                } catch (Exception e) {
                    System.out.println("There was a problem executing query! Maybe you entered an invalid input.");
                }
            }

            System.out.println("Choose an option a or b:");
            System.out.println("a - Find dependent cells");
            System.out.println("b - Find cells with certain value");
            System.out.println("c - exit application");
        }

        dataset.commit();
        file.delete();
        System.out.println("Upload xlsx file, you want to process, via post method http://localhost:8080/upload");
    }
}
