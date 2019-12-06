package com.ReadExcelFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndUpdate {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		File loc=new File("D:\\Selenium\\ExcelDemo\\target\\Book1.xlsx");
		FileInputStream fi=new FileInputStream(loc);
        Workbook w=new XSSFWorkbook(fi);
        Sheet s=w.getSheet("Sheet1");
        Row r=s.getRow(1);
        Cell c=r.getCell(0);
        c.setCellValue("ramana");
        
        fi.close();
        FileOutputStream output_file =new FileOutputStream(loc);  //Open FileOutputStream to write updates
        
        w.write(output_file); //write changes
           
        output_file.close();  
        
       // output_file.close();
        System.out.println("updated");

	}

}
