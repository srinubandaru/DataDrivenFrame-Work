package com.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class AddColorToText {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		

		FileInputStream fis=new FileInputStream("dataFiles/Book1.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		
		Sheet s=wb.getSheet("Sheet1");
		Row r;
		Cell c;
		int rc=s.getLastRowNum();
		
		System.out.println(" Writing data into Excel File");
		System.out.println("Row Count : "+rc);
		
		CellStyle style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
	    //style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    Font font = wb.createFont();
            font.setColor(IndexedColors.GREEN.getIndex());
            style.setFont(font);
		
		for (int i = 1; i <= rc; i++) {
			
			 r=s.getRow(i);
			
			int cc=r.getLastCellNum();
			
			
			
				
			     c=r.createCell(2);
				c.setCellValue("Pass");
				
				
				c.setCellStyle(style);
				
			
				FileOutputStream fo=new FileOutputStream("dataFiles/Book1.xlsx");
				wb.write(fo);
				
			
		}
		
		
		
		
	
		wb.close();

	}

}
