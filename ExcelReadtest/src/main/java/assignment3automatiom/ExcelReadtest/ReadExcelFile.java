package assignment3automatiom.ExcelReadtest;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {

	public static void main(String[] args) throws Exception{

		File src = new File(("getnewexcel.xlsx"));

		FileInputStream fis = new FileInputStream (src);

		XSSFWorkbook wb = new XSSFWorkbook(fis);

		XSSFSheet sheet = wb.getSheetAt(0);

		System.out.println(sheet.getRow(0).getCell(0).getStringCellValue());
		String firstname = sheet.getRow(1).getCell(0).getStringCellValue();
		System.out.println("Value of firstname  : "+firstname );
		System.out.println(sheet.getRow(0).getCell(1).getStringCellValue());
		String lastname = sheet.getRow(1).getCell(1).getStringCellValue();
		System.out.println(" Value of lastname  : "+ lastname);

		String email= firstname.concat(".").concat(lastname).concat("@gmail.com"); 
		System.out.println(" Value of email  : "+  email);
		
		XSSFRow row = sheet.createRow((short)1);
		row.createCell(3).setCellValue(email);
		System.out.println(sheet.getRow(1).getCell(3).getStringCellValue());
		fis.close();
	}

}



