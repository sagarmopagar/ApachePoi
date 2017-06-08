package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		// Create an object of FileInputStream class to read excel file
		int rowno;
		File file = new File("C:\\BAPIS\\Book1.xlsx");
		FileInputStream fs = new FileInputStream(file);
		Workbook guru99Workbook = null;
		guru99Workbook = new XSSFWorkbook(fs);
		Sheet guru99Sheet = guru99Workbook.getSheet("Sheet1");

		int rowCount = guru99Sheet.getLastRowNum() - guru99Sheet.getFirstRowNum();
		for (int i = 0; i < rowCount + 1; i++) {
			// Loop over all the rows
			Row row = guru99Sheet.getRow(i);
			// Check if the first cell contain a value, if yes, That means it is
			// the new testcase name
			// if (row.getCell(0).toString() == "") {
			rowno = i + 1;

			if (row.getCell(0) == null) {

				System.out.println("First cell in Row " + rowno + " Is blank");
			} else {
				System.out.println("First cell in Row " + rowno + " is " + row.getCell(0).toString());
			}

		}

	}
}
