# Automation(3rd Assignment)

## Task Overview
Create a script for getting user data (First Name and Last Name ) through user input.
Add this data in the excel sheet column 'First Name' and 'Last Name'.
Read data from the excel sheet.
Perform Operation(Create Email Id with the help of 'First Name' and 'Last Name' column value). 
Add the output result in the Email ID column.


## Project Setup
- **Automation Code** 
   GIT repository URL:https://github.com/priya16jain/thirdAssignment.git
   
- **Steps**
   - Created a script for getting user data (First Name and Last Name ) through Scanner Class.
   - Created a new scanner with the specified String Object
   - Declared the object and initialize with predefined standard input object(System.in).
   - We have used .nextLine() to read line of text.
   - Created a workbook(XSSF) FOR CREATING .xlsx file .
   - Used FileOutputStream for writing data to the file.
   - Created a sheet (XSSFSheet) in excel file.
   - Added Rows and values in the rows  using .createCell(0).setCellValue.
   - Used FileOutputStream for writing data to the cell.
   - Used FileInputStream to Read data from the excel sheet.
   - Used sheet.getRow().getCell().getStringCellValue() for getting the values present in cell.
   - Created Email Id with the help of 'First Name' and 'Last Name' column value using .concate        
   - function and pass the resultant to the 3 cell of the Sheet.
   - Get its Value using .getRow().getCell().getStringCellValue().
   
   






