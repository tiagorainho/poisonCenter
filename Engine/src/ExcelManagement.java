import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JFileChooser;


public class ExcelManagement {
	private Workbook workbook;
	
	public ExcelManagement(Workbook workbook) {
		this.workbook = workbook;
	}
	
	public void save(String fileName) {
		String outputFileName = "../Output";
		File file = new File(outputFileName);
		if(!file.exists()) {
			file.mkdir();
		}
		try {
			FileOutputStream output = new FileOutputStream(outputFileName + "/" + fileName);
			workbook.write(output);
			output.close();
		}
		catch(Exception e) {
			Output.print(e.getMessage());
		}
	}
	
	public static Workbook getWorkbook(String fileName) {
		try {
			return new XSSFWorkbook(new FileInputStream(new File(fileName)));
		} catch (FileNotFoundException e) {
			Output.print(e.getMessage());
		} catch (IOException e) {
			Output.print(e.getMessage());
		}
		return null;
	}
	
	public static Workbook getInputFile() {
		JFileChooser fileChooser = new JFileChooser();
		if(fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
			try {
				return new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
			}
			catch(FileNotFoundException e) {
				Output.print(e.getMessage());
			}
			catch(IOException e2) {
				Output.print(e2.getMessage());
			}
		}
		return null;		
	}
	
	public static Workbook updateCasList(CasList casList, List<Product> products, String casListSheetName) {
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet(casListSheetName);
		List<Cas> lst = casList.getAllCas();
		Row row;
		Cas cas;
		Cell cell;
		int startingPoint = 2;
		CellStyle headerStyle = Styles.HeaderBoldGray(wb);
		CellStyle warning = Styles.warning(wb);
		
		
		sheet.setColumnWidth(0, 5000);
		sheet.setColumnWidth(1, 5000);
		sheet.setColumnWidth(2, 20000);
		sheet.setColumnWidth(3, 20000);
		
		// add header
		row = sheet.createRow(1);
		row.setHeightInPoints(20);
		
		cell = row.createCell(0);
		cell.setCellValue("CAS-number");
		cell.setCellStyle(headerStyle);
		
		cell = row.createCell(1);
		cell.setCellValue("Product code");
		cell.setCellStyle(headerStyle);
		
		cell = row.createCell(2);
		cell.setCellValue("Description EN");
		cell.setCellStyle(headerStyle);
		
		cell = row.createCell(3);
		cell.setCellValue("Description PT");
		cell.setCellStyle(headerStyle);
		
		// add old cas
		for(int i = 0; i < lst.size(); i++) {
			row = sheet.createRow(i + startingPoint);
			cas = lst.get(i);
			if(cas.getCasNumber() != null) {
				row.createCell(0).setCellValue(cas.getCasNumber());
			}
			if(cas.getProductCode() != null) {
				row.createCell(1).setCellValue(cas.getProductCode());
			}
			if(cas.getDescriptionEN() != null) {
				row.createCell(2).setCellValue(cas.getDescriptionEN());
			}
			if(cas.getDescriptionPT().length() > 0) {
				row.createCell(3).setCellValue(cas.getDescriptionPT());
			}
			else {
				row.createCell(3).setCellStyle(warning);
			}
		}
		
		// add new cas
		int count = startingPoint + lst.size();
		String traduction;
		for(Product product: products) {
			for(Ingredient ingredient: product.getIngredients()) {
				if(!casList.contains(ingredient.getCasNumber())) {
					traduction = casList.getTraduction(ingredient.getCasNumber());
					if(traduction == null || traduction.length() == 0) {
						
						casList.add(new Cas(ingredient.getCasNumber(), ingredient.getProductCode(), ingredient.getName(), ""));
						row = sheet.createRow(count++);
						
						cell = row.createCell(0);
						cell.setCellValue(ingredient.getCasNumber());
						cell.setCellStyle(warning);
						
						cell = row.createCell(1);
						cell.setCellValue(ingredient.getProductCode());
						cell.setCellStyle(warning);
						
						cell = row.createCell(2);
						cell.setCellValue(ingredient.getName());
						cell.setCellStyle(warning);
						
						row.createCell(3).setCellStyle(warning);
					}
				}
			}
		}
		
		return wb;
	}
	
	public static Workbook getOutputDocument(Product product, Company company, CasList cas, Sheet lista, boolean hasHeader) {
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet(product.getName());
		sheet.setColumnWidth(0, 8000);
		sheet.setColumnWidth(1, 15000);
		sheet.setColumnWidth(2, 1000);
		sheet.setColumnWidth(3, 21000);
		sheet.setColumnWidth(4, 7000);
		sheet.setColumnWidth(5, 3000);
		
		int startingPoint = 3;
		int rowsLen = 60;
		
		// load most used styles
		CellStyle markerStyle = Styles.marker(wb);
		CellStyle tableHeader = Styles.TableHeader(wb);
		CellStyle warning = Styles.warning(wb);
		CellStyle underline = Styles.Underlined(wb);
		
		// ficha header
		Cell fichaHeader = sheet.createRow(0).createCell(0);
		fichaHeader.setCellValue("Ficha de Informação de Produto_CIAV");
		fichaHeader.getRow().setHeightInPoints(20);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));
		fichaHeader.setCellStyle(Styles.HeaderBoldBlack(wb));
		
		
		// productName header
		Cell productNameHeader = sheet.createRow(1).createCell(0);
		productNameHeader.getRow().setHeightInPoints(25);
		productNameHeader.setCellValue(product.getName());
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 5));
		productNameHeader.setCellStyle(Styles.HeaderBoldGray(wb));
		
		// create row list
		List<Row> rows = new LinkedList<Row>();
		if(hasHeader) {
			for(int i = 0; i < rowsLen; i++) {
				rows.add(sheet.createRow(i + startingPoint));
			}
		}
		else {
			for(int i = 0; i < rowsLen; i++) {
				rows.add(sheet.createRow(i));
			}
		}
		
		
		// populate markers and field cells
		List<Cell> markers = new LinkedList<Cell>();
		markers.add(rows.get(0).createCell(0));
		markers.add(rows.get(4).createCell(0));
		markers.add(rows.get(13).createCell(0));
		markers.add(rows.get(18).createCell(0));
		markers.add(rows.get(29).createCell(0));
		markers.add(rows.get(32).createCell(0));
		markers.add(rows.get(0).createCell(3));
		
		String[] markersNames = {"1 – Nome comercial do produto",
		                         "2 – Descrição do Produto",
		                         "3 – Embalagem",
		                         "4 – Dados da empresa responsável pela colocação do produto no mercado",
		                         "5 – Data de início da comercialização",
		                         "6 – Referência da fórmula na empresa",
		                         "7 – Composição",
		                         "8 – Dados do fabricante"};
		
		
		// populate markers
		Cell cell;
		for(int i=0;i<markers.size(); i++) {
			cell = markers.get(i);
			cell.setCellValue(markersNames[i]);
			cell.setCellStyle(markerStyle);
			cell.getRow().createCell(cell.getColumnIndex() + 1).setCellStyle(markerStyle);
		}
		cell = markers.get(6);
		cell.setCellValue(markersNames[6]);
		cell.setCellStyle(markerStyle);
		cell.getRow().createCell(cell.getColumnIndex() + 1).setCellStyle(markerStyle);
		cell.getRow().createCell(cell.getColumnIndex() + 2).setCellStyle(markerStyle);
		
		String fields[] = {"Empresa:",
				"Pessoa de contacto:",
				"Cargo:", "Rua e Nº:",
				"Código Postal:",
				"Localidade:",
				"Pais:",
				"Telefone:",
				"E-mail:"};
		
		// populate fields
		rows.get(1).createCell(0).setCellValue(product.getName());
		rows.get(2).createCell(0).setCellValue("Ficha:");
		rows.get(5).createCell(0).setCellValue("Uso Recomendado:");
		rows.get(6).createCell(0).setCellValue("Utilizador:");
		rows.get(7).createCell(0).setCellValue("Cor:");
		rows.get(8).createCell(0).setCellValue("pH:");
		rows.get(9).createCell(0).setCellValue("Apresentação:");
		rows.get(14).createCell(0).setCellValue("Tipo");
		rows.get(15).createCell(0).setCellValue("Tamanho:");
		rows.get(19).createCell(0).setCellValue(fields[0]);
		rows.get(20).createCell(0).setCellValue(fields[1]);
		rows.get(21).createCell(0).setCellValue(fields[2]);
		rows.get(22).createCell(0).setCellValue(fields[3]);
		rows.get(23).createCell(0).setCellValue(fields[4]);
		rows.get(24).createCell(0).setCellValue(fields[5]);
		rows.get(25).createCell(0).setCellValue(fields[6]);
		rows.get(26).createCell(0).setCellValue(fields[7]);
		rows.get(27).createCell(0).setCellValue(fields[8]);
		rows.get(30).createCell(0).setCellStyle(warning);

		// populate autocompletes
		rows.get(2).createCell(1).setCellValue("Nova");
		cell = rows.get(19).createCell(1);
		cell.setCellValue(company.getName());
		cell.setCellStyle(underline);
		rows.get(20).createCell(1).setCellValue(company.getPerson());
		rows.get(21).createCell(1).setCellValue(company.getPosition());
		rows.get(22).createCell(1).setCellValue(company.getStreet());
		rows.get(23).createCell(1).setCellValue(company.getPostalCode());
		rows.get(24).createCell(1).setCellValue(company.getLocality());
		rows.get(25).createCell(1).setCellValue(company.getCountry());
		rows.get(26).createCell(1).setCellValue(company.getPhoneNumber());
		rows.get(27).createCell(1).setCellValue(company.getEmail());
		rows.get(33).createCell(0).setCellValue(product.getReference());
		
		
		// populate table header (7 option)
		String[] tableHeaderTitles = {"Designação", "Nº CAS", "%"};
		List<Cell> tableHeaderCells = new LinkedList<Cell>();
		tableHeaderCells.add(rows.get(1).createCell(3));
		tableHeaderCells.add(rows.get(1).createCell(4));
		tableHeaderCells.add(rows.get(1).createCell(5));
		for(int i=0;i<tableHeaderCells.size();i++) {
			tableHeaderCells.get(i).setCellValue(tableHeaderTitles[i]);
			tableHeaderCells.get(i).setCellStyle(tableHeader);
		}
		
		// populate table with ingredients
		List<Ingredient> ingredients = product.getIngredients();
		String casNumber;
		String name;
		for(int i = 0; i < ingredients.size(); i++) {
			
			// preparing ingredient name
			name = cas.getTraduction(ingredients.get(i).getCasNumber());
			if(name == null || name.length() == 0) {
				cell = rows.get(i+2).createCell(3);
				cell.setCellValue(ingredients.get(i).getName());
				cell.setCellStyle(warning);
			}
			else {
				if(name.length() > 1) {
					name = String.valueOf(name.charAt(0)) + name.substring(1).toLowerCase();
				}
				rows.get(i+2).createCell(3).setCellValue(name);	
			}
			
			// preparing CAS Number
			casNumber = ingredients.get(i).getCasNumber();
			if(casNumber == null || casNumber.startsWith("(")) {
				rows.get(i+2).createCell(4).setCellValue("-");
			}
			else {
				rows.get(i+2).createCell(4).setCellValue(casNumber);
			}
			
			// preparing percentage
			rows.get(i+2).createCell(5).setCellValue(ingredients.get(i).getPercent());
		}
		
		// fit company info
		int height = 33;
		int height7 = ingredients.size() + 1;
		int height8 = fields.length + 1;
		if(height > height7 + height8) {
			
			// typical format
			cell = rows.get(height - height8 + 1).createCell(3);
			cell.setCellValue(markersNames[7]);
			cell.setCellStyle(markerStyle);
			cell.getRow().createCell(cell.getColumnIndex() + 1).setCellStyle(markerStyle);
			cell.getRow().createCell(cell.getColumnIndex() + 2).setCellStyle(markerStyle);
			for(int i = 0; i < fields.length ; i++) {
				rows.get(height - height8 + 2 + i).createCell(3).setCellValue(fields[i]);
			}
		}
		else {
			
			// atypical format
			int begins8 = startingPoint + height7 - 1;
			cell = rows.get(begins8).createCell(3);
			cell.setCellValue(markersNames[7]);
			cell.setCellStyle(markerStyle);
			cell.getRow().createCell(cell.getColumnIndex() + 1).setCellStyle(markerStyle);
			cell.getRow().createCell(cell.getColumnIndex() + 2).setCellStyle(markerStyle);
			for(int i = 0; i < fields.length ; i++) {
				rows.get(begins8 + 1 + i).createCell(3).setCellValue(fields[i]);
			}
			height = begins8 + height8 - 1;
		}
		cell = rows.get(height - height8 + 2).createCell(4);
		cell.setCellValue("");
		cell.setCellStyle(underline);
		
		// get date
		SimpleDateFormat formatter= new SimpleDateFormat("dd.MM.yyyy");
		Date date = new Date(System.currentTimeMillis());
		
		// populate dates
		rows.get(height + 2).createCell(0).setCellValue("Data de informação do CIAV:");
		rows.get(height + 2).createCell(1).setCellStyle(warning);
		rows.get(height + 4).createCell(0).setCellValue("Data de atualização:");
		rows.get(height + 4).createCell(1).setCellValue(formatter.format(date));
		
		// add lista
		Sheet listaSheet = wb.createSheet("lista");
		PoiCopySheet.copySheets(listaSheet, lista);
		
		// preparing border style specs
		BorderStyle borderStyle = BorderStyle.THICK;
		short borderColor = IndexedColors.BLACK.getIndex();
		
		// preparing border style positions
		CellStyle DefaultBorderStyleBottom = wb.createCellStyle();
		DefaultBorderStyleBottom.setBorderBottom(borderStyle);
		DefaultBorderStyleBottom.setBottomBorderColor(borderColor);
		
		CellStyle DefaultBorderStyleRight = wb.createCellStyle();
		DefaultBorderStyleRight.setBorderRight(borderStyle);
		DefaultBorderStyleRight.setRightBorderColor(borderColor);
		
		CellStyle DefaultBorderStyleLeft = wb.createCellStyle();
		DefaultBorderStyleLeft.setBorderLeft(borderStyle);
		DefaultBorderStyleLeft.setLeftBorderColor(borderColor);
		
		CellStyle DefaultBorderStyleTop = wb.createCellStyle();
		DefaultBorderStyleTop.setBorderTop(borderStyle);
		DefaultBorderStyleTop.setTopBorderColor(borderColor);
		
		// drawing border bottom
		for(int i = 0; i <= 5; i++) {
			cell = rows.get(height).getCell(i);
			if(cell == null) {
				cell = rows.get(height).createCell(i);
			}
			CellStyle newStyle = wb.createCellStyle();
			newStyle.cloneStyleFrom(cell.getCellStyle());
			newStyle.setBorderBottom(borderStyle);
			newStyle.setBottomBorderColor(borderColor);
			cell.setCellStyle(newStyle);
		}
		
		// drawing border left and right
		for(int i = 0; i <= height; i++) {
			
			// right side
			cell = rows.get(i).getCell(5);
			if(cell == null) {
				cell = rows.get(i).createCell(5);
			}
			CellStyle newStyleRight = wb.createCellStyle();
			newStyleRight.cloneStyleFrom(cell.getCellStyle());
			newStyleRight.setBorderRight(borderStyle);
			newStyleRight.setRightBorderColor(borderColor);
			cell.setCellStyle(newStyleRight);
			
			// left side
			cell = rows.get(i).getCell(0);
			if(cell == null) {
				cell = rows.get(i).createCell(0);
			}
			CellStyle newStyleLeft = wb.createCellStyle();
			newStyleLeft.cloneStyleFrom(cell.getCellStyle());
			newStyleLeft.setBorderLeft(borderStyle);
			newStyleLeft.setLeftBorderColor(borderColor);
			cell.setCellStyle(newStyleLeft);
		}
		
		
		// listing header rows
		List<Row> headerRows = new LinkedList<Row>();
		Row row;
		for(int i = 0; i < startingPoint; i++) {
			row = sheet.getRow(i);
			if(row == null) {
				row = sheet.createRow(i);
			}
			headerRows.add(row);
		}
		
		if(hasHeader) {
			
			// drawing header border
			for(int i = 0; i < headerRows.size(); i++) {
				
				// right side
				cell = headerRows.get(i).getCell(5);
				if(cell == null) {
					cell = headerRows.get(i).createCell(5);
				}
				CellStyle styleRight = wb.createCellStyle();
				styleRight.cloneStyleFrom(cell.getCellStyle());
				styleRight.setBorderRight(borderStyle);
				styleRight.setRightBorderColor(borderColor);
				cell.setCellStyle(styleRight);
				
				// left side
				cell = headerRows.get(i).getCell(0);
				if(cell == null) {
					cell = headerRows.get(i).createCell(0);
				}
				CellStyle styleLeft = wb.createCellStyle();
				styleLeft.cloneStyleFrom(cell.getCellStyle());
				styleLeft.setBorderLeft(borderStyle);
				styleLeft.setLeftBorderColor(borderColor);
				cell.setCellStyle(styleLeft);
			}
		}
		
		// drawing top
		for(int i = 0; i <= 5; i++) {
			cell = headerRows.get(0).getCell(i);
			if(cell == null) {
				cell = headerRows.get(0).createCell(i);
			}
			CellStyle style = wb.createCellStyle();
			style.cloneStyleFrom(cell.getCellStyle());
			style.setBorderTop(borderStyle);
			style.setTopBorderColor(borderColor);
			cell.setCellStyle(style);
		}
         
		return wb;
	}	
}
