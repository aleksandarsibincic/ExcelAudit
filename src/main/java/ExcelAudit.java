import com.ontotext.trree.OwlimSchemaRepository;
import org.apache.jena.base.Sys;
import org.apache.jena.ontology.*;
import org.apache.jena.query.*;
import org.apache.jena.rdf.model.Model;
import org.apache.jena.rdf.model.ModelFactory;
import org.apache.jena.rdf.model.Property;
import org.apache.jena.tdb.TDB;
import org.apache.jena.tdb.TDBFactory;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.usermodel.*;
import org.eclipse.rdf4j.model.impl.TreeModel;
import org.eclipse.rdf4j.repository.RepositoryConnection;
import org.eclipse.rdf4j.repository.manager.LocalRepositoryManager;
import org.eclipse.rdf4j.repository.manager.RepositoryManager;
import org.eclipse.rdf4j.repository.sail.SailRepository;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ExcelAudit {

    public static final String SAMPLE_XLSX_FILE_PATH = "./sample_file.xlsx";
    public static final String ONTOLOGY_PATH = "./excel_ontology.owl";
    public static final String NAMESPACE = "http://www.semanticweb.org/ds/ontologies/2020/5/excel-audit#";

    public static void main(String[] args) throws IOException {

        /*
        prvi dio je vise vjezbanje da li radi excel
        ovo ce kasnije sve brisati
         */

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
            System.out.print(cell.getAddress().toString() + " ");
            dataFormatter.printCellValue(cell);
            System.out.println();
        });

        /* Create Ontology model */
        OntModel model = ModelFactory.createOntologyModel(OntModelSpec.OWL_DL_MEM);
        File ontologyFile = new File(ONTOLOGY_PATH);
        model.read(ontologyFile.toURL().toString());

        /* adding data from excel to ontology model */

        /*create workbook class*/
        OntClass workbookClass = model.getOntClass(NAMESPACE + "Workbook");
        /*create workbook individual*/
        Individual workbookIndividual = workbookClass.createIndividual(NAMESPACE + new File(SAMPLE_XLSX_FILE_PATH).getName());
        /*add filename property to the workbook*/
        DatatypeProperty filename = model.getDatatypeProperty(NAMESPACE + "filename");
        workbookIndividual.addProperty(filename, new File(SAMPLE_XLSX_FILE_PATH).getName());

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
        sheets.forEach(sheet -> {
            /* create spreadsheet individual*/
            Individual spreadsheetIndividual = spreadsheetClass.createIndividual(NAMESPACE + sheet.getSheetName());
            /* add relations between workbook and sheet*/
            model.add(workbookIndividual, hasSheets, spreadsheetIndividual);
            /* get all cells from sheet */
            List<Cell> sheetCells = excelReader.getCells(sheet);
            sheetCells.forEach(cell -> {
                /* create cell individual*/
                Individual cellIndividual = cellClass.createIndividual(NAMESPACE + cell.getAddress().toString());
                /* add relations between sheet and cells*/
                model.add(spreadsheetIndividual, hasCell, cellIndividual);
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        if (!DateUtil.isCellDateFormatted(cell)) {
                            double cellValue = cell.getNumericCellValue();
                            cellIndividual.addProperty(value, String.valueOf(cellValue));
                        }
                        break;
                    case FORMULA:
                        String formula = cell.getCellFormula();
                        /* maybe to use org.apache.poi.ss.formula.FormulaParser for this */
                    default:
                        //do nothing
                }
            });
        });
        model.write(System.out);


        OwlimSchemaRepository schema = new OwlimSchemaRepository();

        // set the data folder where GraphDB will persist its data
        schema.setDataDir(new File("./local-storage"));

        // configure GraphDB with some parameters
        Map<String, String> parameters = new HashMap<>();
        parameters.put("storage-folder", "./");
        parameters.put("repository-type", "file-repository");
        parameters.put("ruleset", "rdfs");
        schema.setParameters(parameters);

        // store Ontology model to in memory TDB, which is not persisted over the sessions

        Dataset dataset = TDBFactory.createDataset();
        dataset.begin(ReadWrite.WRITE);
        dataset.addNamedModel(NAMESPACE, model);
        dataset.commit();

        // extract saved ontology model from TDB and execute simple query

        dataset.begin(ReadWrite.READ);
        Model queryModel = dataset.getNamedModel(NAMESPACE);

        String queryString = "PREFIX rdf: <http://www.semanticweb.org/ds/ontologies/2020/5/excel-audit#> " +
                "SELECT * WHERE { ?Cell rdf:value \"4.0\" . }";
        Query query = QueryFactory.create(queryString);
        QueryExecution queryExecution = QueryExecutionFactory.create(query, queryModel);

        ResultSet resultSet = queryExecution.execSelect();
        try {
            while (resultSet.hasNext()) {
                QuerySolution solution = resultSet.nextSolution();
                System.out.println(solution.get("Cell"));
            }
        } finally {
            queryExecution.close();
        }
        dataset.commit();

    }
}
