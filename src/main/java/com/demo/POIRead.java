package com.demo;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;

public class POIRead {

    /*读取Excel,桌面的*/
    public void getExcel(String excelName,String sheetName) throws Exception{
        FileInputStream in = new FileInputStream("C:/Users/Administrator/Desktop/"+excelName);
        XSSFWorkbook xwb=new XSSFWorkbook(in);
        XSSFSheet sheet = xwb.getSheet(sheetName);
        for (int i=0;i<4;i++){
            XSSFRow row = sheet.getRow(i);
            for (int j=0;j<4;j++){
                XSSFCell cell = row.getCell(j);
                String stringCellValue = cell.getStringCellValue();
                System.out.print(stringCellValue+"  ");
            }
            System.out.println();
        }
    }

    @Test
    public void get() throws Exception {
        String excelName="Sep.xlsx";
        String sheetName="POI创建";
        getExcel(excelName,sheetName);
    }
}
