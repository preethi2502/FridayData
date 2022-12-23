package com.test.in;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataRead {

	public static void main(String[] args) throws IOException {

		File f = new File("C:\\Users\\ADMIN\\eclipse-workspace\\new\\ZZDataDriven\\morningdataz.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet sheetAt = wb.getSheetAt(0);
//ALL DATA from the EXCEL 
		int rowsize = sheetAt.getPhysicalNumberOfRows();

		for (int i = 0; i < rowsize; i++) {

			Row row = sheetAt.getRow(i);
			int cellsize = row.getPhysicalNumberOfCells();

			for (int j = 0; j < cellsize; j++) {
				Cell cell = row.getCell(j);

				CellType cellType = cell.getCellType();

				if (cellType.equals(cellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue + "  : STRING");
				}

				else if (cellType.equals(cellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int data = (int) numericCellValue;

					System.out.println(data + "  : NUMERIC");

				}

			}

		}
	}
}