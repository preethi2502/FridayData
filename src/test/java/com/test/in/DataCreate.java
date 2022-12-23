package com.test.in;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataCreate {

	public static void main(String[] args) throws Exception {
		File f = new File("C:\\Users\\ADMIN\\eclipse-workspace\\new\\ZZDataDriven\\morningdataz.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		wb.createSheet("Nov").createRow(0).createCell(0).setCellValue("CarName");

		wb.getSheet("Nov").getRow(0).createCell(1).setCellValue("Number");

		wb.getSheet("Nov").createRow(1).createCell(0).setCellValue("Benz");
		wb.getSheet("Nov").getRow(1).createCell(1).setCellValue("987/8545");

		wb.getSheet("Nov").createRow(2).createCell(0).setCellValue("BMW");
		wb.getSheet("Nov").getRow(2).createCell(1).setCellValue("456/545");

		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		wb.close();

		System.out.println("write data completed");

	}

}
