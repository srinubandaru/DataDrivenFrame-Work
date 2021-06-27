package com.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SetDataExcel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		

		FileInputStream fis=new FileInputStream("dataFiles/Book1.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		
		Sheet s=wb.getSheet("Sheet1");
		Row r;
		Cell c;
		int rc=s.getLastRowNum();
		
		System.out.println(" Writing data into Excel File");
		System.out.println("Row Count : "+rc);
		
		for (int i = 1; i <= rc; i++) {
			
			 r=s.getRow(i);
			
			int cc=r.getLastCellNum();
			
			
			
				
			     c=r.getCell(2);
				c.setCellValue("Pass");
				
				
				
			
				FileOutputStream fo=new FileOutputStream("dataFiles/Book1.xlsx");
				wb.write(fo);
				
			
		}
		
		
		System.out.println(" Getting data into Excel File");
		
    for (int i = 0; i <= rc; i++) {
			
			 r=s.getRow(i);
			
			int cc=r.getLastCellNum();
			
			for (int j = 0; j < cc; j++) {
				
				c=r.getCell(j);
				
				String data=c.getStringCellValue();
				
				System.out.print( data +" ");
				
				
				
			}
			System.out.println();
			
			
		}
		
		
	
		wb.close();

	}

}
