package com.chi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author chi  2018-05-15 15:51
 **/
public class XlsMain {

    private Workbook wb;
    private FileInputStream fis;
    public Workbook createWorkbook(String filePath) throws IOException{
        if(filePath != null){
            fis = new FileInputStream(filePath);
            if(filePath.endsWith(".xls")){
                //2003版本的excel，用.xls结尾
                wb = new HSSFWorkbook(fis);//得到工作簿
            }else if(filePath.endsWith(".xlsx")){
                //2007版本的excel，用.xlsx结尾
                wb = new XSSFWorkbook(fis);//得到工作簿
            }else{
            }
            return wb;
        }
        return null;
    }

    public static void main(String[] args) {

        XlsMain ddl = new XlsMain();
        Workbook wb = null;
        try {
            wb = ddl.createWorkbook("D:\\sdlfj\\学科词汇.xls");
        } catch (IOException e) {
            e.printStackTrace();
        }
        ddl.doSomething(wb);
    }


    public void doSomething(Workbook wb){
        if(wb != null){
            Sheet sheet;
            Row row;
            //一个工作簿可能不止一个sheet表格
            for(int i = 0; i < wb.getNumberOfSheets(); i ++){
                sheet = wb.getSheetAt(i);
                //循环遍历每个sheet表的没行数据
                for(int j = 1; j < sheet.getPhysicalNumberOfRows(); j ++){
                    row = sheet.getRow(j);
                    //用迭代遍历，因为我看见它有一个iterator（）方法
                    try {
                        for(Iterator<Cell> cell = row.iterator(); cell.hasNext() ;){
                            System.out.print( cell.next().toString() + "  ");
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    System.out.println();
                }
            }
        }
    }



}
