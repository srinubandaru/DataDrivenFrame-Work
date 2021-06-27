package com.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class dataFromExcel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException 
	{
		FileInputStream fis=new FileInputStream("dataFiles/Book1.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		
		Sheet s=wb.getSheet("Sheet1");
		
		int rc=s.getLastRowNum();
		
		for (int i = 0; i <= rc; i++) {
			
			Row r=s.getRow(i);
			
			int cc=r.getLastCellNum();
			
			for (int j = 0; j < cc; j++) {
				
				Cell c=r.getCell(j);
				
				String data=c.getStringCellValue();
				
				System.out.print( data +" ");
				
				
				
			}
			System.out.println();
			
			
		}
		
		
		
	}

}
