package org.excelread;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

public class ExcelRead {

	// .xls 97 - 2003 (HSSF)
	// .xlsx 2003 till date (XSSF)

	@Ignore
	@Test
	public void readExcel() throws IOException {
		File f = new File(
				System.getProperty("user.dir") + "/src/test/resources/Student Details - Project Class Dec 5.xlsx");
		FileInputStream input = new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Student Details");
		int totalRows = sheet.getPhysicalNumberOfRows(); // 12

		for (int i = 0; i < totalRows; i++) {
			XSSFRow row = sheet.getRow(i);
			int totalCells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < totalCells; j++) {
				XSSFCell cell = row.getCell(j);
				if (cell.getCellType() == CellType.NUMERIC) {
					double numericCellValue = cell.getNumericCellValue();
					System.out.println(numericCellValue + " ");
				} else {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue + " ");
				}
			}
			System.out.println("");
		}
		workbook.close();
	}
	
	@Test
	public void writeExcel() throws IOException
	{
		File f = new File(
				System.getProperty("user.dir") + "/src/test/resources/Student Details - Project Class Dec 5.xlsx");
		FileInputStream input = new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Student Details");
		int totalRows = sheet.getPhysicalNumberOfRows(); // 12
		XSSFRow row = sheet.getRow(11);
		// row.getCell(1).setCellValue("Karthik");
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("test");
		FileOutputStream out = new FileOutputStream(f);
		workbook.write(out);
		workbook.close();
		out.close();
	}
}
