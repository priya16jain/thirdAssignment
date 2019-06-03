package assignment3automatiom.ExcelReadtest;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteTranslatedLanguage {
	private static String[] columns = { "first", "second"};
	private static List<Lang> TestContent = new ArrayList<Lang>();

	public static void writeLang(String tolang,String fromLang) throws IOException, 
	InvalidFormatException {
		TestContent.add(new Lang(tolang,fromLang));

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("contacts");

		Font headerFont = workbook.createFont();
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Create Other rows and cells with contacts data
		int rowNum = 1;

		for ( Lang content : TestContent) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(content.first);
			row.createCell(1).setCellValue(content.second);
		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("contacts.xlsx");
		workbook.write(fileOut);
		fileOut.close();
	}
}
