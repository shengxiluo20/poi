package com.chi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

/**
 * @author chi  2018-05-15 15:51
 **/
public class XlsMain {

    private Workbook wb;
    private FileInputStream fis;

    public Workbook createWorkbook(String filePath) throws IOException {
        if (filePath != null) {
            fis = new FileInputStream(filePath);
            if (filePath.endsWith(".xls")) {
                //2003版本的excel，用.xls结尾
                wb = new HSSFWorkbook(fis);//得到工作簿
            } else if (filePath.endsWith(".xlsx")) {
                //2007版本的excel，用.xlsx结尾
                wb = new XSSFWorkbook(fis);//得到工作簿
            } else {
            }
            return wb;
        }
        return null;
    }
/*
    public static void main(String[] args) {

        XlsMain ddl = new XlsMain();
        Workbook wb = null;
        try {
            wb = ddl.createWorkbook("D:\\sdlfj\\学科词汇.xls");
        } catch (IOException e) {
            e.printStackTrace();
        }
        ddl.doSomething(wb);
    }*/

    public Map<String, Integer> doSomething(Workbook wb) {
        System.out.println("读取excel的数据:");
        TreeMap<String, Integer> treeMap = new TreeMap<String, Integer>();
        if (wb != null) {
            Sheet sheet;
            Row row;
            //一个工作簿可能不止一个sheet表格
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                sheet = wb.getSheetAt(i);
                //循环遍历每个sheet表的没行数据
                for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {
                    row = sheet.getRow(j);
                    //用迭代遍历，因为我看见它有一个iterator（）方法
                    try {
                        for (Iterator<Cell> cell = row.iterator(); cell.hasNext(); ) {
                            String s = cell.next().toString().trim().toLowerCase();
                            treeMap.put(s,0);
                            System.out.print(s + "  ");

                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    System.out.println();
                }
            }
        }
        System.out.println();
        System.out.println("excel中共读出"+ treeMap.size() +"个单词");
        System.out.println("============================");
        return treeMap;
    }

}
