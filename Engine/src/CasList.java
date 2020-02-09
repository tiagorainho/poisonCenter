import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CasList {
	private List<Cas> lst;
	
	public CasList() {
		this.lst = new LinkedList<Cas>();
	}
	
	public CasList(String fileName) {
		this.lst = new LinkedList<Cas>();
		XSSFWorkbook wb = null;
		try {
			wb = new XSSFWorkbook(new FileInputStream(new File(fileName)));
		} catch (FileNotFoundException e) {
			Output.print(e.getMessage());
		} catch (IOException e) {
			Output.print(e.getMessage());
		}
		Sheet sheet = wb.getSheet("casList");
		Row row;
		for(int i = 2;i <= sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			try {
				lst.add(new Cas(row.getCell(0).getRichStringCellValue().getString(),
						row.getCell(1).getRichStringCellValue().getString(),
						row.getCell(2).getRichStringCellValue().getString(),
						row.getCell(3).getRichStringCellValue().getString()));
			}
			catch(NullPointerException e) {
				
			}
			
		}
	}
	
	public String getTraduction(String casNumber) {
		for(Cas c: this.lst) {
			if(c.getCasNumber().equals(casNumber)) {
				return c.getDescriptionPT();
			}
		}
		return null;
	}
	
	public void add(Cas cas) {
		if(!this.contains(cas.getCasNumber())) {
			this.lst.add(cas);
		}
	}
	
	public boolean contains(String casNumber) {
		for(Cas c: this.lst) {
			if(casNumber.equals(c.getCasNumber())) {
				return true;
			}
		}
		return false;
	}
	
	public List<Cas> getAllCas() {
		return this.lst;
	}
	
	public int length() {
		return this.lst.size();
	}

}
