package assignment3automatiom.ExcelReadtest;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadTranslatorFile {
	

	File src = new File(("contacts.xlsx"));

	FileInputStream fis = new FileInputStream (src);

	XSSFWorkbook wb = new XSSFWorkbook(fis);

	XSSFSheet contacts = wb.getSheetAt(0);

	Iterator<Row> iterator = contacts.iterator();
    
    while (iterator.hasNext()) {
        Row nextRow = iterator.next();
        Iterator<Cell> cellIterator = nextRow.cellIterator();
         
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
             
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    System.out.print(cell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    System.out.print(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    System.out.print(cell.getNumericCellValue());
                    break;
            }
            System.out.print(" - ");
        }
        System.out.println();
    }
     
    workbook.close();
    inputStream.close();
}

}

