package com.read.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Data {

	public static void main(String[] args) throws Throwable {

		// to read a file

		File f = new File("C:\\Users\\Windows\\eclipse-workspace\\Maven_Appache\\Maven project.xlsx");

		// to read file data values

		FileInputStream fis = new FileInputStream(f);

		// to read a excel sheet

		Workbook wb = new XSSFWorkbook(fis);

		// sheet//rows//column//data
		// to read data from sheet

		Sheet sheetAt = wb.getSheetAt(0);
		int row_size = sheetAt.getPhysicalNumberOfRows();

		// get data using loops

		for (int i = 0; i < row_size; i++) {
			Row row = sheetAt.getRow(i);

			int cell_size = row.getPhysicalNumberOfCells();

//			for (int j = 0; j < cell_size; j++) {

				Cell cell = row.getCell(1);

				CellType cellType = cell.getCellType();

				if (cellType.equals(cellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);

				}

				else if (cellType.equals(cellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int Value = (int) numericCellValue;
					System.out.println(Value);
				}
				
					// to get the particular cell values(column)

				for (int k = 5; k < cell_size; k++) {

					Cell cell1 = row.getCell(0);

					CellType cellType1 = cell1.getCellType();

					if (cellType1.equals(cellType1.STRING)) {
						String stringCellValue = cell1.getStringCellValue();
						System.out.println(stringCellValue);			
					
					
					
				}

				
			}

		}

	}

}
