import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

public class main {

	private static final String casListFile = "../auxiliar/casList.xlsx";
	private static final String listaFile = "../auxiliar/lista.xlsx";
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
	
	public static void processInputChemges() throws Exception{
		
		// read config file
		Config config;
		Company company;
		boolean hasHeader = true;
		String chemgesFile;
		int IngredientFirstPosition, maxIngredientsPerProduct, productCodeIndex, preparationNumberIndex, descriptionIndex, UFIcodeCalculatedIndex, UFIcodeNotifiedIndex, EuPCSIndex;
		try {
			config = new Config("../auxiliar/config.xlsx");
			company = new Company(config.getField("Empresa"),
					config.getField("Pessoa de contacto"),
					config.getField("Cargo"),
					config.getField("Rua e Nº"),
					config.getField("Código Postal"),
					config.getField("Localidade"),
					config.getField("Pais"),
					config.getField("Telefone"),
					config.getField("E-mail"));
			
			if(config.getField("Ter header").equalsIgnoreCase("false")) hasHeader = false;
			IngredientFirstPosition = Integer.parseInt(config.getField("Ingredients"));
			maxIngredientsPerProduct = Integer.parseInt(config.getField("Número máximo de ingredientes"));			
			productCodeIndex = Integer.parseInt(config.getField("Product code"));
			preparationNumberIndex = Integer.parseInt(config.getField("Preparation number"));
			descriptionIndex = Integer.parseInt(config.getField("Description 1"));
			UFIcodeCalculatedIndex = Integer.parseInt(config.getField("UFI code calculated"));
			UFIcodeNotifiedIndex = Integer.parseInt(config.getField("UFI code notified"));
			EuPCSIndex = Integer.parseInt(config.getField("EuPCS"));
			chemgesFile = config.getField("Nome ficha");
		}
		catch(Exception e) {
			throw new Exception("Erro ao ler o ficheiro de configuracao: " + e.getMessage());
		}
		int maxLenIngredients = maxIngredientsPerProduct*4;
		
		// get CAS
		// CasList cas = new CasList(casListFile);
		CasList cas = null;
		
		// get input file
		Workbook inputFile = ExcelManagement.getInputFile();

		// get lista
		Workbook listaWorkbook = null;
		Sheet lista;
		try {
			listaWorkbook = new XSSFWorkbook(new FileInputStream(new File(listaFile)));
			lista = listaWorkbook.getSheetAt(0);
			listaWorkbook.close();
		}
		catch(Exception e) {
			throw new Exception("Erro ao ler a lista: " + e.getMessage());
		}
		
		// get input file objects (Products)
		List<Product> products = new LinkedList<Product>();
		
		Sheet sheet;
		Row header;
		try{
			sheet= inputFile.getSheet(chemgesFile);
			header = sheet.getRow(0);
		}
		catch(Exception e) {
			throw new Exception("Erro ao ler o ficheiro Chemges: " + e.getMessage());
		}
		
		Cell cell0, cell1, cell2,cell3;
		int rowIndex = 0;
		for(Row row: sheet) {
			rowIndex++;
			if(!row.equals(header) && row.getCell(productCodeIndex)!= null) {
				
				if(row.getCell(productCodeIndex).getRichStringCellValue().getString().startsWith("XXP")) {
					
					// create product
					Product product;
					try {
						product = new Product(
								row.getCell(productCodeIndex).getRichStringCellValue().getString(),
								row.getCell(preparationNumberIndex).getRichStringCellValue().getString(),
								row.getCell(descriptionIndex).getRichStringCellValue().getString(),
								row.getCell(UFIcodeCalculatedIndex).getRichStringCellValue().getString(),
								row.getCell(UFIcodeNotifiedIndex).getRichStringCellValue().getString(),
								row.getCell(EuPCSIndex).getRichStringCellValue().getString());
					}
					catch(Exception e) {
						throw new Exception("Erro ao ler informacao do produto da linha " + rowIndex + ": " + e.getMessage());
					}
					
					// add ingredients
					Ingredient ingredient;
					try {
						for(int i = 0; i<maxLenIngredients; i=i+4) {
							cell0 = row.getCell(IngredientFirstPosition+i);
							cell1 = row.getCell(IngredientFirstPosition+i+1);
							cell2 = row.getCell(IngredientFirstPosition+i+2);
							cell3 = row.getCell(IngredientFirstPosition+i+3);
							if(cell0 == null && cell1 == null && cell2 == null  && cell3 == null ) continue;
							
							ingredient = new Ingredient();
							if(cell0 != null) ingredient.setCasNumber(cell0.getRichStringCellValue().getString());
							if(cell1 != null) ingredient.setProductCode(cell1.getRichStringCellValue().getString());
							if(cell2 != null) ingredient.setDescription(cell2.getRichStringCellValue().getString());
							if(cell3 != null) ingredient.setPercent(cell3.getRichStringCellValue().getString());
							product.addIngredient(ingredient);
						}
					}
					catch(Exception e) {
						throw new Exception("Erro ao ler os ingredientes do produto da linha " + rowIndex + ": " + e.getMessage());
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
			try {
				workbook = ExcelManagement.getOutputDocument(product, company, cas, lista, hasHeader);
				excel = new ExcelManagement(workbook);
				excel.save(product.getName() + ".xlsx");
			}
			catch(Exception e) {
				throw new Exception("Erro ao escrever a ficha do produto \"" + product.getName() + "\": " + e.getMessage());
			}
		}
		
		// update cas list
		//excel = new ExcelManagement(ExcelManagement.updateCasList(cas, products, "casList"));
		//excel.save(casListFile);
		
	}
}
