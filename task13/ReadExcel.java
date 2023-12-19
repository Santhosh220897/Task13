package task13;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) {

		// object is created for class ReadExcel

		ReadExcel obj = new ReadExcel();

		String value;

		// method1 & method2 are surrounded with try & catch method

		try {
			// calling method1

			obj.readingExcel();

			// calling method 2
			value = obj.readingExcel("Sheet1", 1, 2);

			// to get the program is completed.

			System.out.println("Completed");
		} catch (Exception e) {

			e.printStackTrace();
		}

	}

	// Method 1 with whole excel sheet without arguments

	public void readingExcel() throws Exception {

		String filePath = "C:\\Users\\santh\\Desktop\\ReadExcel.xlsx";// file path

		// Dataformatter class is imported for reading the file in excel

		DataFormatter format = new DataFormatter();

		String result = null;

		// FileInout Stream is imported for opening the excel file & read

		FileInputStream inStream = new FileInputStream(filePath);

		XSSFWorkbook book = new XSSFWorkbook(inStream);

		// Nested for loop is used for iteration

		for (int i = 0; i <= 4; i++) {

			for (int j = 0; j <= 4; j++) {

				// to get the data from the sheet1

				XSSFCell cell = book.getSheet("Sheet1").getRow(i).getCell(j);

				result = format.formatCellValue(cell);

				System.out.print(result + " ");
			}

			System.out.println(" ");

		}

		book.close();// close the workbook
		inStream.close();

	}

	// method 2 with same method name and with arguments
	// method is used for reading specific row and cell

	public String readingExcel(String sheet, int row, int column) throws Exception {

		String result = null;

		String filePath = "C:\\Users\\santh\\Desktop\\ReadExcel.xlsx";// file path of excel

		DataFormatter format = new DataFormatter();

		FileInputStream inStream = new FileInputStream(filePath);

		XSSFWorkbook book = new XSSFWorkbook(inStream);

		XSSFCell cell = book.getSheet(sheet).getRow(row).getCell(column);// to get the specific row & cell

		result = format.formatCellValue(cell);

		System.out.println(cell);

		book.close();// close the workbook
		inStream.close();

		return (result);
	}

}

//output

/*Name     Age Email    
 John Doe  30 john@test.com    
 Jane Doe  28 john@test.com    
 Bob Smith 35 jacky@example.com    
 Swapnil   37 joe@example.com    
 john@test.com
         Completed */


