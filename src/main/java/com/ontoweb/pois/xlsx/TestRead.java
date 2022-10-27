package com.ontoweb.pois.xlsx;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class TestRead {

    public static String getValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        // 把数字当成String来读，避免出现1读成z`1.0的情况

//        cell.setCellType(CellType.STRING);
        if (cell.getCellType() == CellType.BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == CellType.NUMERIC) {
            cellValue = String.valueOf(cell.getNumericCellValue());
        } else if (cell.getCellType() == CellType.STRING) {
            cellValue = String.valueOf(cell.getStringCellValue());
        } else if (cell.getCellType() == CellType.FORMULA) {
            cellValue = String.valueOf(cell.getCellFormula());
        } else if (cell.getCellType() == CellType.BLANK) {
            cellValue = " ";
        } else if (cell.getCellType() == CellType.ERROR) {
            cellValue = "非法字符";
        } else {
            cellValue = "未知类型";
        }
        return cellValue;
    }

    public static List<List<String>> test(File file) {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        List<String> rowList=null;//一行数据(原始未处理过)
        List<List<String>> rowLists=new ArrayList<>();//一张表的全部数据(原始未处理过)
        //=====获取行列数
        Sheet sheet = workbook.getSheetAt(0);
        int rowSize = sheet.getPhysicalNumberOfRows();
        Row row;
        Cell cell;
        int cellSize = sheet.getRow(0).getPhysicalNumberOfCells();
        //=====设置文档里面的每个值为String
        for (int i = 0; i < rowSize; i++) {
            rowList = new ArrayList<>();
            row = sheet.getRow(i);
            if (i==0) {
                for (int j = 0; j < cellSize; j++) {
                    cell = row.getCell(j);
                    CellType cellType = cell.getCellType();
                    if (!cellType.name().equals("STRING"))
                        cell.setCellType(CellType.STRING);//设置每个值为String
                    String celV = cell.getStringCellValue();
                    rowList.add(celV+"");
                }
            }else {
                for (int j = 0; j < cellSize; j++) {
                    cell = row.getCell(j);
                    CellType cellType = cell.getCellType();
                    if (!cellType.name().equals("STRING"))
                        cell.setCellType(CellType.STRING);//设置每个值为String
                    String celV = getValue(cell);
//                    String celV = cell.getStringCellValue();
                    rowList.add(celV);
                }
            }
            rowLists.add(rowList);
        }
        return rowLists;
    }

    public static void main(String[] args) {
        System.out.println(test(new File("D:\\项目文件\\疾控\\智能报告系统\\全国居民健康素养监测调查问卷.xlsx")));
    }
}
