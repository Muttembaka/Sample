package com.ReadExcelFile;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateNewExcelSheet {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		File loc=new File("D:\\Selenium\\ExcelDemo\\target\\NewExcel.xlsx");
		FileOutputStream fo=new FileOutputStream(loc);
        Workbook w=new XSSFWorkbook();
        
        Sheet s=w.createSheet("Suman");
        //Sheet s=w.getSheet("Sheet1");
        
        /*Row r=s.getRow(1);
        Cell c=r.getCell(0);
        c.setCellValue("ramana");
        
        
*/     
        
        
        Row row1 = s.createRow((short) 0);
        
        
        
     // inserting first row cell value
      
     row1.createCell(0).setCellValue("Serial Number"); 
    
        w.write(fo);
        fo.close();

	}

}
