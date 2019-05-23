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
		System.out.println(sheet.getRow(0).getCell(1).getStringCellValue());
		System.out.println(sheet.getRow(0).getCell(2).getStringCellValue());

		for(int i =1; ;i++) {

			String firstname = sheet.getRow(i).getCell(0).getStringCellValue();
			System.out.println("Value of firstname  : "+firstname );
			
			String lastname = sheet.getRow(i).getCell(1).getStringCellValue();
			System.out.println(" Value of lastname  : "+ lastname);

			String email= firstname.concat(".").concat(lastname).concat("@gmail.com"); 
			System.out.println(" Value of email  : "+  email);
			
			XSSFRow rowhead = sheet.createRow((short)1);
			rowhead.createCell(2).setCellValue(email);
            sheet.getRow(i).getCell(2).setCellValue(email);
			
			fis.close();
		}

		
	}

}


