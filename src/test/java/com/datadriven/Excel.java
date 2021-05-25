package com.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.All;


public class Excel {
	
	public static void particular_data() throws Throwable {
		
		File f= new File("C:\\Users\\elcot\\eclipse-workspace\\datadriven\\userdetails\\base excel.xlsx");
	  
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb= new XSSFWorkbook(fis);
		
		Sheet sheetAt = wb.getSheetAt(0);
		
		Row row = sheetAt.getRow(3);
		
		Cell cell = row.getCell(0);
		
		CellType cellType = cell.getCellType();
		
		if(cellType.equals(cellType.STRING))
		{
			String Value = cell.getStringCellValue();
			System.out.println(Value);
		}
		
		if(cellType.equals(cellType.NUMERIC))
		{
			double numericCellValue = cell.getNumericCellValue();
			int nv=(int) numericCellValue;
			System.out.println(nv);
		}
}
	
	public static void alldata() throws Throwable {
		File f= new File("C:\\Users\\elcot\\eclipse-workspace\\datadriven\\userdetails\\base excel.xlsx");
		  
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb= new XSSFWorkbook(fis);
		
		Sheet sheetAt = wb.getSheetAt(0);
		
		int NumberOfRows = sheetAt.getPhysicalNumberOfRows();
		
		for (int i = 0; i <NumberOfRows; i++) {
			
			Row row = sheetAt.getRow(i);
			
			int NumberOfCells = row.getPhysicalNumberOfCells();
			
			for (int j = 0; j < NumberOfCells; j++) {
				Cell cell = row.getCell(j);
				
				CellType cellType = cell.getCellType();
				
				if(cellType.equals(cellType.STRING))
				{
					String Value = cell.getStringCellValue();
					System.out.println(Value);
				}
				
				if(cellType.equals(cellType.NUMERIC))
				{
					double numericCellValue = cell.getNumericCellValue();
					int nv=(int) numericCellValue;
					System.out.println(nv);
				}
				
			}
			
		}		
	}
	
	public static void allrowdata() throws Throwable {
		File f= new File("C:\\Users\\elcot\\eclipse-workspace\\datadriven\\userdetails\\base excel.xlsx");
		  
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb= new XSSFWorkbook(fis);
		
		Sheet sheetAt = wb.getSheetAt(0);
		
		Row row = sheetAt.getRow(3);
		
		int NumberOfCells = row.getPhysicalNumberOfCells();
		
		
		for (int i = 0; i < NumberOfCells; i++) {
			
			Cell cell = row.getCell(i);
			CellType cellType = cell.getCellType();
			
			if(cellType.equals(cellType.STRING))
			{
				String Value = cell.getStringCellValue();
				System.out.println(Value);
			}
			
			if(cellType.equals(cellType.NUMERIC))
			{
				double numericCellValue = cell.getNumericCellValue();
				int nv=(int) numericCellValue;
				System.out.println(nv);
			}

			
		}
	
	}
	public static void columndata() throws Throwable {
		File f= new File("C:\\Users\\elcot\\eclipse-workspace\\datadriven\\userdetails\\base excel.xlsx");
		  
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb= new XSSFWorkbook(fis);
		
		Sheet sheetAt = wb.getSheetAt(0);
		
		int NumberOfRows = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < NumberOfRows; i++) {
			
			Row row = sheetAt.getRow(i);
			Cell cell = row.getCell(1);
            CellType cellType = cell.getCellType();
			
			if(cellType.equals(cellType.STRING))
			{
				String Value = cell.getStringCellValue();
				System.out.println(Value);
			}
			
			if(cellType.equals(cellType.NUMERIC))
			{
				double numericCellValue = cell.getNumericCellValue();
				int nv=(int) numericCellValue;
				System.out.println(nv);
			}

		}	
		
	}
	public static void main(String[] args) throws Throwable {
		System.out.println("***Particular data***");
		particular_data();
		System.out.println("***All data***");
		alldata();
		System.out.println("***row data***");
		allrowdata();
		System.out.println("***column data***");
		columndata();
	}
}
