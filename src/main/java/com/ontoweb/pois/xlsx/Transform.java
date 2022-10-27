package com.ontoweb.pois.xlsx;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Transform {
    static final List<String> predicts = Arrays.asList("ET020","ET060");
    static DecimalFormat df = new DecimalFormat("#.##");
    public static void transformData(File file, int sheetNum) throws IOException {
        String fileType = file.getPath().substring(file.getPath().lastIndexOf(".") + 1);
        Workbook workbook = null;
        if ("xls".equals(fileType)) {
            workbook = new HSSFWorkbook(Files.newInputStream(file.toPath()));
        }else if ("xlsx".equals(fileType)){
            workbook = new XSSFWorkbook(Files.newInputStream(file.toPath()));
        }
        if (workbook == null) throw new RuntimeException("工作簿创建失败");
        Sheet sheet = workbook.getSheetAt(sheetNum);
        int rowSize = sheet.getPhysicalNumberOfRows();
        List<List<String>> resList = new ArrayList<>();
        Row row;
        Cell cell;
        int cellSize = sheet.getRow(0).getPhysicalNumberOfCells();
        List<String> rowList;
        String strValue;
        for (int i = 0; i < rowSize; i++) {
            rowList = new ArrayList<>();
            row = sheet.getRow(i);
            if (i==0) {
                for (int j = 2; j < cellSize; j++) {
                    cell = row.getCell(j);
                    if (!predicts.contains(cell.getStringCellValue())) {
                        cell.setCellValue(String.format("A%03d", j + 1));
                    }else {
                        cell.setCellValue("pre col");
                    }
                }
            }else {
                for (int j = 2; j < cellSize; j++) {
                    cell = row.getCell(j);
                    cell.setCellType(CellType.STRING);
                    String celV = cell.getStringCellValue();
                    if (isNumberic(celV)) {
                        cell.setCellValue(df.format(Double.parseDouble(celV)));
                    }else if (celV.length() >= 4) {
                        cell.setCellValue(celV.substring(0, 4));
                    }
                }
            }
        }
        FileOutputStream outFile = new FileOutputStream("src/files/out.xlsx");
        workbook.write(outFile);
        workbook.close();
        outFile.close();
    }

    public static boolean isNumberic(String obj) {
        try{
            Integer.parseInt(obj.trim());
            return true;
        }catch(NumberFormatException e)  {
            System.out.println("异常：\"" + obj + "\"不是数字/整数...");
            return false;
        }
    }

    public static void main(String[] args) {
        try {
            transformData(new File("src/files/海工测试数据-1.xlsx"), 0);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
//        double d =11.103;
//        DecimalFormat df = new DecimalFormat("#.##");
//        String s=df.format(d);//1.0
////        if(s.indexOf(".") > 0){
////            s = s.replaceAll("0+?$", "");//去掉多余的0
////            s = s.replaceAll("[.]$", "");//如最后一位是.则去掉
////        }
//        System.out.println(s);//1.5
    }
}
