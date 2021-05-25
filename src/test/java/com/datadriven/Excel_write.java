package com.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_write {
	
	public static void write() throws Throwable {
		
		File f= new File("C:\\Users\\elcot\\Desktop\\write.xlsx");
		
		FileInputStream fis=new FileInputStream(f);
		
		Workbook wb=new XSSFWorkbook(fis);
		
		Sheet createSheet = wb.createSheet("write_details");
		
		Row createRow = createSheet.createRow(0);
		
		Cell createCell = createRow.createCell(0);
		
		createCell.setCellValue("username");
		
		wb.getSheet("write_details").getRow(0).createCell(1).setCellValue("password");
		
		FileOutputStream fos=new FileOutputStream(f);
		
		wb.write(fos);
		
		wb.close();
		
		System.out.println("data created");
		
	}
public static void main(String[] args) throws Throwable {
	
	write();
	
}
}
