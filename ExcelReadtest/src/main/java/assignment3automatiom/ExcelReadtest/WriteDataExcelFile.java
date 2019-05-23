package assignment3automatiom.ExcelReadtest;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataExcelFile {

	public static void main(String[]args) {
		try {

			String Firstname;
			String Lastname ;
			int count = 1;

			XSSFWorkbook workbook = new XSSFWorkbook();
			FileOutputStream file = new FileOutputStream(new File("getnewexcel.xlsx"));
			workbook.write(file);
			file.close();
			XSSFSheet sheet = workbook.createSheet("FirstSheet"); 
			XSSFRow rowhead = sheet.createRow((short)0);
			rowhead.createCell(0).setCellValue("FirstName");
			rowhead.createCell(1).setCellValue("LastName");
			rowhead.createCell(2).setCellValue("Email"); 

			Scanner in = new Scanner(System.in);

			while(count >= 1){
				System.out.println("Enter a Firstname ");
				Firstname = in.nextLine();
				System.out.println("You entered " + Firstname);

				System.out.println("Enter a Lastname");
				Lastname = in.nextLine();
				System.out.println("You entered " + Lastname);

				if(Firstname == null || Firstname.isEmpty()){
					System.out.println("Name is empty "); 
					break;
				}

				int rowCount = sheet.getLastRowNum();

				Row newRow = sheet.createRow(rowCount+1);

				newRow.createCell(0).setCellValue(Firstname);
				newRow.createCell(1).setCellValue(Lastname);
			}	
			
            FileOutputStream fileOut = new FileOutputStream("getnewexcel.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			System.out.println("Your excel file has been generated!");
	}
			catch ( Exception ex ) {
			System.out.println(ex);
		}
	}
}