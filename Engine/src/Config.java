import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellType;

public class Config {
	private Workbook workbook;
	
	public Config(String fileName) {
		XSSFWorkbook wb = null;
		try {
			wb = new XSSFWorkbook(new FileInputStream(new File(fileName)));
		} catch (FileNotFoundException e) {
			Output.print(e.getMessage());
		} catch (IOException e) {
			Output.print(e.getMessage());
		}
		this.workbook = wb;
	}
	
	public String getField(String field) {
		Sheet sheet = workbook.getSheetAt(0);
		for(Row row: sheet) {
			if(row.getCell(0).getRichStringCellValue().toString().trim().replace(":", "").equalsIgnoreCase(field.trim())) {
				if(row.getCell(1).getCellType() == CellType.STRING) return row.getCell(1).getRichStringCellValue().toString();
				else if(row.getCell(1).getCellType() == CellType.NUMERIC) return String.valueOf((int)Math.round(row.getCell(1).getNumericCellValue()));
			}
		}
		throw new NullPointerException("campo \"" + field + "\" nao encontrado");
	}
	
}
