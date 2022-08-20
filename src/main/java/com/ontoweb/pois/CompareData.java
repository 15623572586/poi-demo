package com.ontoweb.pois;

import com.ontoweb.pois.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class CompareData {
    public static List<List<String>> readData(File file, int sheetNum) throws IOException {
        String fileType = file.getPath().substring(file.getPath().lastIndexOf(".") + 1);
        Workbook workbook = null;
        if ("xls".equals(fileType)) {
            workbook = new HSSFWorkbook(new FileInputStream(file));
        }else if ("xlsx".equals(fileType)){
            workbook = new XSSFWorkbook(new FileInputStream(file));
        }
        if (workbook == null) throw new RuntimeException("工作簿创建失败");
        Sheet sheet = workbook.getSheetAt(sheetNum);
        int rowSize = sheet.getPhysicalNumberOfRows();
        List<List<String>> resList = new ArrayList<List<String>>();
        Row row;
        Cell cell;
        int cellSize = 'S' - 'A';
        List<String> rowList;
        String strValue;
        for (int i = 2; i < rowSize; i++) {
            rowList = new ArrayList<String>();
            row = sheet.getRow(i);
            for (int j = 0; j < cellSize; j++) {
                cell = row.getCell(j);
                try {
                    strValue = cell.getStringCellValue();
                } catch (Exception e) {
                    strValue = "";
                    e.printStackTrace();
                }
                rowList.add(strValue);
            }
            resList.add(rowList);
        }
        return resList;
    }
    public static void merge(List<List<String>> EQMSData, File file, File outFile, int sheetNum) throws IOException {
        String fileType = file.getPath().substring(file.getPath().lastIndexOf(".") + 1);
        Workbook workbook = null;
        if ("xls".equals(fileType)) {
            workbook = new HSSFWorkbook(new FileInputStream(file));
        }else if ("xlsx".equals(fileType)){
            workbook = new XSSFWorkbook(new FileInputStream(file));
        }
        if (workbook == null) throw new RuntimeException("工作簿创建失败");
        Sheet sheet = workbook.getSheetAt(sheetNum);
        int rowSize = sheet.getPhysicalNumberOfRows();
        Row row;
        Cell cell;
        String strValue = "";
        String oldCode = "";
        String newCode;
        boolean flag = false;
        int startRow = 2;
        int endRow;
        for (int i = 2; i <= rowSize; i++) {
            if (i < rowSize) {
                row = sheet.getRow(i);
                cell = row.getCell('D' - 'A');
                try {
                    cell.setCellType(CellType.STRING);
                    strValue = cell.getStringCellValue();
                } catch (Exception e) {
                    strValue = "";
                    e.printStackTrace();
                }
            }
            if(StringUtils.isEmpty(oldCode)) {
                try {
                    oldCode = strValue.substring(0, 9);
                } catch (Exception e) {
                    oldCode = strValue;
                    e.printStackTrace();
                }
                startRow = i;
            }else {
                try {
                    newCode = strValue.substring(0, 9);
                } catch (Exception e) {
                    newCode = strValue;
                    e.printStackTrace();
                }
                if (!newCode.equals(oldCode) || i == rowSize) { // i == rowSize考虑遍历结束的情况
                    endRow = i - 1;
                    int height = endRow - startRow + 1;
                    // 熊EQMS数据中获取同9码的数据 List
                    List<List<String>> subEqmsData = getSameCode9Data(EQMSData, oldCode);
                    int size = subEqmsData.size();
                    if (size > height) { // 同九码EQMS导出数据比左边测点多
                        System.out.println(oldCode + " : " + (size - height));
                        // 从i行开始后面所有行的数据向下移动（size-height）
                        if (i <= sheet.getLastRowNum()) {
                            sheet.shiftRows(i, sheet.getLastRowNum(), size - height);
                            rowSize += (size - height);
                            i += (size - height);
                            endRow = i - 1;
                        }else
                            System.out.println(i);
                    }

                    List<String> rowData = null;
                    int idx = 0; // subEqmsData的索引
                    for (int j = startRow; j <= endRow && idx < size; j++) { //在右边填充
                        Row backRow = sheet.getRow(j);
                        if (backRow == null) backRow = sheet.createRow(j);
                        rowData = subEqmsData.get(idx++);
                        int c =0;
                        for (int k = 'P' - 'A'; k < 'Z' - 'A'; k++) { // 填充一行
                            cell = backRow.createCell(k);
                            cell.setCellValue(rowData.get(c++));
                        }
                    }
                    if (flag) {
                        // 设置颜色
                        flag= false;
                        setCellStyle(workbook, sheet, startRow, endRow, 0, 'Y' - 'A', IndexedColors.AQUA.getIndex());
                    }else {
                        flag = true;
                        setCellStyle(workbook, sheet, startRow, endRow, 0, 'Y' - 'A', IndexedColors.YELLOW1.getIndex());
                    }
                    startRow = i;
                    oldCode = newCode;
                }
            }
        }
        setAutoWith(sheet);
        FileOutputStream os = new FileOutputStream(outFile);
        workbook.write(os);
        os.flush();
        os.close();
    }

    private static List<List<String>> getSameCode9Data(List<List<String>> eqmsData, String oldCode) {
        List<List<String>> subEqmsData = new ArrayList<List<String>>();
        String code9 = "";
        List<String> rowData;
        for (List<String> data: eqmsData) {
            rowData = new ArrayList<String>();
            code9 = data.get('M' - 'A');
            if (StringUtils.isNotEmpty(code9)) {
                code9 = code9.trim().substring(0, 9);
                if (code9.equals(oldCode)) {
                    rowData.add(data.get('B' - 'A'));
                    rowData.add(data.get('C' - 'A'));
                    rowData.add(data.get('D' - 'A'));
                    rowData.add(data.get('E' - 'A'));
                    rowData.add(data.get('F' - 'A'));
                    rowData.add(data.get('G' - 'A'));
                    rowData.add(data.get('M' - 'A'));
                    rowData.add(data.get('N' - 'A'));
                    rowData.add(data.get('Q' - 'A'));
                    rowData.add(data.get('R' - 'A'));
                    subEqmsData.add(rowData);
                }
            }

        }
        return subEqmsData;
    }
    // 设置单元格样式
    public static void setCellStyle(Workbook workbook, Sheet sheet, int startRow, int endRow, int startCol, int endCol, short colorIndex) {
        CellStyle style = workbook.createCellStyle();
        for (int j = startRow; j <= endRow; j++) {
            //style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            for (int k = startCol; k <= endCol; k++) {
                Cell cell1 = sheet.getRow(j).getCell(k);
                if (cell1 == null) cell1 = sheet.getRow(j).createCell(k);
                style.setFillForegroundColor(colorIndex);// 设置背景色
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
//                sheet.autoSizeColumn(k);
                cell1.setCellStyle(style);
            }
        }
    }
    // 设置列宽自适应
    public static void setAutoWith(Sheet sheet) {
        int colSize = sheet.getRow(0).getPhysicalNumberOfCells();
        for (int i = 0; i <= colSize; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    public static void main(String[] args) throws IOException {

//        File eqmsFile = new File("D:\\tmp\EQMS比较\\设备9码清单-EQMS导出\\能环设备9码清单.xlsx"); //能环
//        File pointFile = new File("D:\\tmp\EQMS比较\\EMS.xlsx"); //能环
//        File outFile = new File("D:\\tmp\EQMS比较\\out\\能环.xlsx"); //能环

//        File eqmsFile = new File("D:\\tmp\EQMS比较\\设备9码清单-EQMS导出\\冷轧设备9码清单.xlsx"); //冷轧
//        File pointFile = new File("D:\\tmp\EQMS比较\\冷轧测点302+PI.xlsx"); //冷轧
//        File outFile = new File("D:\\tmp\EQMS比较\\out\\冷轧.xlsx"); //冷轧

        File eqmsFile = new File("D:\\tmp\\EQMS比较\\设备9码清单-EQMS导出\\炼铁设备9码清单.xlsx"); //炼铁
        File pointFile = new File("D:\\tmp\\EQMS比较\\炼铁项目导入汇总220412.xlsx"); //炼铁点表
        File outFile = new File("D:\\tmp\\EQMS比较\\out\\炼铁-1.xlsx"); //炼铁
        List<List<String>> eqmsData = readData(eqmsFile, 1);

        merge(eqmsData, pointFile, outFile, 1);
    }
}
