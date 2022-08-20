package com.ontoweb.pois;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 按照目标模板转储数据
 */
public class DataDump {

    /**
     * 按sheet读数据
     */
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
        int cellSize = 14;
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
    public static void writeData(List<List<String>> data, File file, int sheetNum, File outFile) throws IOException {
        String fileType = file.getPath().substring(file.getPath().lastIndexOf(".") + 1);
        Workbook workbook = null;
        if ("xls".equals(fileType)) {
            workbook = new HSSFWorkbook(new FileInputStream(file));
        }else if ("xlsx".equals(fileType)){
            workbook = new XSSFWorkbook(new FileInputStream(file));
        }
        if (workbook == null) throw new RuntimeException("工作簿创建失败");
        Sheet sheet = workbook.getSheetAt(sheetNum);
        int cellSize = sheet.getRow(0).getPhysicalNumberOfCells();
        Row row;
        Cell cell;
        List<String> dataRow;
        for (int i = 0; i < data.size(); i++) {
            dataRow = data.get(i);
            row = sheet.createRow(i + 2);
            for (int j = 0; j < cellSize; j++) {
                cell = row.createCell(j);
                switch (j) {
                    case 'E' - 'A': cell.setCellValue("BGBWAKD0"); break;
                    case 'I' - 'A': cell.setCellValue(dataRow.get('N' - 'A')); break;
                    case 'J' - 'A': cell.setCellValue(dataRow.get('F' - 'A')); break;
                    case 'L' - 'A': cell.setCellValue(dataRow.get('D' - 'A')); break;
                    case 'M' - 'A': cell.setCellValue("001"); break;
                    case 'N' - 'A': cell.setCellValue(dataRow.get('E' - 'A')); break;
                    case 'Q' - 'A': {
                        if (dataRow.get('F' - 'A').contains("震动") || dataRow.get('F' - 'A').contains("振动"))
                            cell.setCellValue("N");
                        else
                            cell.setCellValue("G");
                    } break;
                    case 'R' - 'A': cell.setCellValue(dataRow.get('F' - 'A')); break;
                    case 'S' - 'A': cell.setCellValue(dataRow.get('H' - 'A')); break;
                    case 'T' - 'A': cell.setCellValue(dataRow.get('I' - 'A')); break;
                    case 'U' - 'A': cell.setCellValue("在线采集"); break;
                    case 'V' - 'A': cell.setCellValue("PDA"); break;
                    case 'W' - 'A': cell.setCellValue("A"); break;
                }
            }
        }
        OutputStream os = new FileOutputStream(outFile);
        workbook.write(os);
        os.flush();
        os.close();
    }
    public static void dumpData() throws IOException {
//        File sourceFile = new File("F:\\tmp\\EMS(已配置数据的3292）设备导入.xlsx"); // 冷环
//        File outFile = new File("F:\\tmp\\output\\EMS-3292.xlsx");
//        File sourceFile = new File("F:\\tmp\\炼铁6#高炉导入表.xlsx"); // 炼铁
//        File outFile = new File("F:\\tmp\\output\\炼铁6#高炉.xlsx");
//        File sourceFile = new File("F:\\tmp\\炼铁C3导入表.xlsx"); // 炼铁
//        File outFile = new File("F:\\tmp\\output\\炼铁C3.xlsx");
//        File sourceFile = new File("F:\\tmp\\炼铁新2烧数据导入.xlsx"); // 炼铁
//        File outFile = new File("F:\\tmp\\output\\炼铁新2烧.xlsx");
        File sourceFile = new File("F:\\tmp\\冷轧测点302+PI.xlsx"); //冷轧
        File outFile = new File("F:\\tmp\\output\\冷轧302+PI.xlsx");
        File templateFile = new File("F:\\tmp\\9码对照表模板.xlsx");
        if (outFile.exists()) outFile.delete();
        List<List<String>> data = readData(sourceFile, 1);
        writeData(data, templateFile, 1, outFile);



    }

    public static void main(String[] args) {
        try {
            dumpData();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
