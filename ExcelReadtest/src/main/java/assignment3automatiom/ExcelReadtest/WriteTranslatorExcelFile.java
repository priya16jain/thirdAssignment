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

public class WriteTranslatorExcelFile {

	private static String[] columns = { "language1", "language2"};
	private static List<lang> TestContent = new ArrayList<lang>();

	public static void main(String[] args) throws IOException, 
	InvalidFormatException {
		TestContent.add(new lang("Apple",""));
		TestContent.add(new lang("Banana",""));
		TestContent.add(new lang("Orange",""));
		TestContent.add(new lang("Mango",""));

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

		for ( lang content : TestContent) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(content.language1);
			row.createCell(1).setCellValue(content.language2);
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
