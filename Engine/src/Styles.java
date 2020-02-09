import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class Styles {
	
	public static CellStyle warning(Workbook wb) {
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return style;
	}
	
	public static CellStyle marker(Workbook wb) {
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return style;
	}
	
	public static CellStyle Underlined(Workbook wb) {
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setUnderline(Font.U_SINGLE);
		style.setFont(font);
		return style;
	}
	
	public static CellStyle HeaderBoldBlack(Workbook wb) {
		CellStyle styleBoldBlack = wb.createCellStyle();
		Font boldBlack = wb.createFont();
		boldBlack.setBold(true);
		styleBoldBlack.setFont(boldBlack);
		styleBoldBlack.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		styleBoldBlack.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		boldBlack.setColor(IndexedColors.WHITE.getIndex());
		styleBoldBlack.setAlignment(HorizontalAlignment.CENTER);
		styleBoldBlack.setVerticalAlignment(VerticalAlignment.CENTER);
		boldBlack.setFontHeightInPoints((short) 16);
		return styleBoldBlack;
	}
	
	public static CellStyle HeaderBoldGray(Workbook wb) {
		CellStyle styleBoldGray = wb.createCellStyle();
		Font boldGray = wb.createFont();
		boldGray.setBold(true);
		styleBoldGray.setFont(boldGray);
		styleBoldGray.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		styleBoldGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleBoldGray.setAlignment(HorizontalAlignment.CENTER);
		styleBoldGray.setVerticalAlignment(VerticalAlignment.CENTER);
		boldGray.setFontHeightInPoints((short) 16);
		return styleBoldGray;
	}
	
	public static CellStyle TableHeader(Workbook wb) {
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return style;
	}
	
	
}
