package com.demo;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

public class POI {


    /*创建Excel*/
   public void CreateExcel(String sheetName,String excelName) throws Exception{
        XSSFWorkbook xwb=new XSSFWorkbook();
        XSSFSheet sheet = xwb.createSheet(sheetName);
        /*for循环每行填充数据*/
        int z=1;
        for(int i=0;i<4;i++){
            XSSFRow row = sheet.createRow(i);
            for(int j=0;j<4;j++){
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(z+"");
                z++;
            }
        }
        FileOutputStream out =new FileOutputStream("C:/Users/Administrator/Desktop/"+excelName);
        xwb.write(out);
        xwb.close();
    }


    public  void  create() throws Exception {
       String sheetName="POI创建";
       String excelName="Sep.xlsx";
       CreateExcel(sheetName,excelName);
    }
}
