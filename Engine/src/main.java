import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

public class main {

	private static final String casListFile = "../auxiliar/casList.xlsx";
	private static final String listaFile = "../auxiliar/lista.xlsx";
	private static final boolean hasHeader = true;
	private static final String outputPath = "../Output";
	
	public static void main(String[] args) throws IOException {		
		
		try {
			processInputChemges();
			Output.print("Processo concluído com sucesso !", true);
		}
		catch(Exception e) {
			Output.print(e.getMessage());
		}
		
		try {
			Desktop.getDesktop().open(new File(outputPath));
		} catch (Exception e) {
			Output.print("Erro a abrir o diretório com os ficheiros finais");
		}
		
		System.exit(0);
	}
	
	public static void processInputChemges() {
		
		// read config file
		Config config = new Config("../auxiliar/config.xlsx");
		Company company = new Company(config.getField("Empresa:"),
				config.getField("Pessoa de contacto:"),
				config.getField("Cargo:"),
				config.getField("Rua e Nº:"),
				config.getField("Código Postal:"),
				config.getField("Localidade:"),
				config.getField("Pais:"),
				config.getField("Telefone:"),
				config.getField("E-mail:"));
		
		// get CAS
		CasList cas = new CasList(casListFile);
		
		// get input file
		Workbook inputFile = ExcelManagement.getInputFile();

		// get lista
		Workbook listaWorkbook = null;
		try {
			listaWorkbook = new XSSFWorkbook(new FileInputStream(new File(listaFile)));
		} catch (FileNotFoundException e) {
			Output.print(e.getMessage());
		} catch (IOException e) {
			Output.print(e.getMessage());
		}
		Sheet lista = listaWorkbook.getSheetAt(0);
		
		// get input file objects (Products)
		List<Product> products = new LinkedList<Product>();
		Sheet sheet = inputFile.getSheet("Portugal");
		Row header = sheet.getRow(0);
		
		int firstPosition = 20;
		int maxIngredientsPerProduct = 40;
		int maxLenIngredients = firstPosition + 4 * maxIngredientsPerProduct;
		Cell cell0, cell1, cell2,cell3;
		for(Row row: sheet) {
			if(!row.equals(header) && row.getCell(0)!= null) {
				if(row.getCell(0).getRichStringCellValue().getString().startsWith("XXP")) {
					
					// create product
					Product product = new Product(row.getCell(0).getRichStringCellValue().getString(),
							row.getCell(1).getRichStringCellValue().getString(),
							row.getCell(2).getRichStringCellValue().getString());
					
					// add ingredients
					for(int i = firstPosition; i < maxLenIngredients; i = i+4) {
						cell0 = row.getCell(i);
						cell1 = row.getCell(i+1);
						cell2 = row.getCell(i+2);
						cell3 = row.getCell(i+3);
						if(cell0 == null && cell1 == null && cell2 == null  && cell3 == null ) {
							i = maxLenIngredients;
							continue;
						}
						
						Ingredient ingredient = new Ingredient();
						if(cell0 != null) {
							ingredient.setCasNumber(cell0.getRichStringCellValue().getString());
						}
						if(cell1 != null) {
							ingredient.setProductCode(cell1.getRichStringCellValue().getString());
						}
						if(cell2 != null) {
							ingredient.setDescription(cell2.getRichStringCellValue().getString());
						}
						if(cell3 != null) {
							ingredient.setPercent(cell3.getRichStringCellValue().getString());
						}
						product.addIngredient(ingredient);
					}
					
					// add product to products list
					products.add(product);
				}
			}
		}
		
		// produce output file
		Workbook workbook;
		ExcelManagement excel;
		
		for(Product product: products) {
			workbook = ExcelManagement.getOutputDocument(product, company, cas, lista, hasHeader);
			excel = new ExcelManagement(workbook);
			excel.save(product.getName() + ".xlsx");
		}
		
		excel = new ExcelManagement(ExcelManagement.updateCasList(cas, products, "casList"));
		excel.save(casListFile);
		
	}
	
}
