package task13;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) {

		// Object creation for the class

		WriteExcel obj = new WriteExcel();

		try {

			// calling the createexcel method and sorround with try catch
			obj.createExcel();

			System.out.println("Completed");

		} catch (Exception e) {

			e.printStackTrace();
		}

	}

	// method for creating the excel sheet and write a data

	public void createExcel() throws Exception {

		String filePath = "C:\\Users\\santh\\Desktop\\WriteExcel.xlsx";// file path to create a excel sheet

		File file = new File(filePath);// Create new excel sheet

		FileOutputStream outStream = new FileOutputStream(file);// class to write in excel sheet

		XSSFWorkbook book = new XSSFWorkbook();

		// sheet named as sheet1
		XSSFSheet sheet = book.createSheet("Sheet1");

		// writing the data for R0

		sheet.createRow(0).createCell(0).setCellValue("Name");// RO,C0
		sheet.getRow(0).createCell(1).setCellValue("Age");// RO,C1
		sheet.getRow(0).createCell(2).setCellValue("Email");// RO,C2

		sheet.createRow(1).createCell(0).setCellValue("John Doe");// R1,C0
		sheet.getRow(1).createCell(1).setCellValue("30");// R1,C1
		sheet.getRow(1).createCell(2).setCellValue("john@test.com");// R1,C2

		sheet.createRow(2).createCell(0).setCellValue("Jane Doe");// R2,C0
		sheet.getRow(2).createCell(1).setCellValue("28");// R2,C1
		sheet.getRow(2).createCell(2).setCellValue("john@test.com");// R2,C2

		sheet.createRow(3).createCell(0).setCellValue("Bob Smith");// R3,C0
		sheet.getRow(3).createCell(1).setCellValue("35");// R3,C1
		sheet.getRow(3).createCell(2).setCellValue("jacky@example.com");// R3,C2

		sheet.createRow(4).createCell(0).setCellValue("Swapnil ");// R4,C0
		sheet.getRow(4).createCell(1).setCellValue("37");// R4,C1
		sheet.getRow(4).createCell(2).setCellValue("joe@example.com");// R4,C2

		// Calling outstream to write the data
		book.write(outStream);

		// close the excel
		book.close();
		outStream.close();

	}
}

//Output

/*Name	    Age	Email
John Doe	30	john@test.com
Jane Doe	28	john@test.com
Bob Smith	35	jacky@example.com
Swapnil 	37	joe@example.com   */

