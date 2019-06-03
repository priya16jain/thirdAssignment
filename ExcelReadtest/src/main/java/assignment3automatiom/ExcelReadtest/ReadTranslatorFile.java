package assignment3automatiom.ExcelReadtest;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadTranslatorFile {

	public static List<String> extractAsList(){
		List<String> list = new ArrayList<String>();
		int maxDataCount =0;
		try{
			File src = new File(("contacts.xlsx"));
			FileInputStream fis = new FileInputStream (src);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet contacts = wb.getSheetAt(0);
			Iterator<Row> rowIterator = contacts.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				//Skip the first row beacause it will be header
				if (row.getRowNum() == 0) {
					maxDataCount = row.getLastCellNum();
					continue;
				}
				for(int cn=1; cn<maxDataCount; cn++) {
					Cell cell= row.getSheet().getRow(cn).getCell(0);

					switch (cell.getCellType()) {

					case Cell.CELL_TYPE_STRING:
						list.add(cell.getStringCellValue());
						break;
					}

				}
				fis.close();
			}
		}catch (Exception e) {  e.printStackTrace();}
		return list;

	}

}