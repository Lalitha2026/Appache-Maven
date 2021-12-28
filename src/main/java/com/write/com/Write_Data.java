package com.write.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data {

	public static void main(String[] args) throws Throwable {

		System.out.println("Creting a excel data in the name of DATA");
		
		File f = new File("C:\\Users\\Windows\\eclipse-workspace\\Maven_Appache\\Maven project.xlsx");
		
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fis);
		
// Create a sheet
		
		Sheet createSheet = wb.createSheet("Login Credentials");
		
		//create row
		
		Row createRow = createSheet.createRow(0);
		
		//create cell
		
		Cell createCell = createRow.createCell(0);
		
		//set the Values
		
		createCell.setCellValue("User Data");
		
		//set the value in the second cell
		
		wb.getSheet("Login Credentials").getRow(0).createCell(1).setCellValue("User Password");
		wb.getSheet("Login Credentials").createRow(1).createCell(0).setCellValue("Legha");
		wb.getSheet("Login Credentials").getRow(1).createCell(1).setCellValue("Kl@202620");
		
//		wb.getSheet("Login Credentials").createRow(1).createCell(0).setCellValue("Kamal");
//		wb.getSheet("Login Credentials").getRow(1).createCell(1).setCellValue("Kannanlek2026");
		
		FileOutputStream fos = new FileOutputStream(f);
		
			//write
		
		wb.write(fos);
		
		//close
		
		wb.close();
		
		//sys out println
		
		System.out.println("DATA Sheet Created Successfully");
		
		
		

	}

}
