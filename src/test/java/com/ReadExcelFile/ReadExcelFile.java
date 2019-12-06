package com.ReadExcelFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		File loc=new File("D:\\Selenium\\ExcelDemo\\target\\Book1.xlsx");
		FileInputStream fi=new FileInputStream(loc);
        Workbook w=new XSSFWorkbook(fi);
        Sheet s=w.getSheet("Sheet1");
        
        System.out.println(s.getPhysicalNumberOfRows());
        //s.getPhysicalNumberOfRows();
        Row r=s.getRow(0);
        Cell cell=r.getCell(0);
        System.out.println(cell);
        
        for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
        	
        	Row r1=s.getRow(i);
        	
        	for(int j=0;j<r1.getPhysicalNumberOfCells();j++) {
        		Cell c=r1.getCell(j);
        		//System.out.println(c);
        		
        		int val=c.getCellType();
        		if(val==1) {
        			String text=c.getStringCellValue();
        			System.out.println(text);
        		}
        		else {
        			if(DateUtil.isCellDateFormatted(c)) {
        				SimpleDateFormat sim =new SimpleDateFormat("dd-MMM-yy");
        				String dateSim=sim.format(c.getDateCellValue());
        				System.out.println(dateSim);
        			}
        			else {
        				double d=c.getNumericCellValue();
        				long l=(long)d;
        				System.out.println(l);
        				String StringL=String.valueOf(l);
        				System.out.println(StringL);
        				System.out.println(l+7845112221564478787l);
        			}
        		}
        	}
        }
		
	}

}
