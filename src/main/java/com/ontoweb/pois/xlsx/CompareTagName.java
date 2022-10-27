package com.ontoweb.pois.xlsx;

import com.ontoweb.pois.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class CompareTagName {
    public static HashMap<String, List<String>> readData(File file, int sheetNum, int targetCol) throws IOException {
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
        List<String> resList = new ArrayList<>();
        Row row;
        Cell cell;
        String strValue;
        List<String> repeatedTagNames = new ArrayList<>();
        for (int i = 2; i < rowSize; i++) {
            row = sheet.getRow(i);
            cell = row.getCell(targetCol);
            try {
                strValue = cell.getStringCellValue();
            } catch (Exception e) {
                strValue = "";
                e.printStackTrace();
            }
            if (StringUtils.isEmpty(strValue)) continue;
            if (resList.contains(strValue)) repeatedTagNames.add(strValue);
            else resList.add(strValue);
        }
        HashMap<String, List<String>> resMap = new HashMap<>();
        resMap.put("result", resList);
        resMap.put("repeated", repeatedTagNames);
        return resMap;
    }



    public static void main(String[] args) throws IOException {
        HashMap<String, List<String>> resMap1 = readData(new File("D:\\tmp\\宝信整理的点表\\冷轧\\表1—冷轧9码对照表模板 （杨福祥）20220712.xlsx"), 1, 'J' - 'A');
        List<String> list1 = resMap1.get("result");
        HashMap<String, List<String>> resMap2 = readData(new File("D:\\tmp\\宝信整理的点表\\冷轧\\二冷轧计控设备（20220714版） - 副本.xlsx"), 0, 'J' - 'A');
        List<String> list2 = resMap2.get("result");
        for (String tagName : list1) {
            if (!list2.contains(tagName)) System.out.println(tagName);
        }

        HashMap<String, List<String>> resMap = readData(new File("D:\\tmp\\宝信整理的点表\\冷轧\\二冷轧计控设备（20220714版）还有两个重复测点点.xlsx"), 0, 'J' - 'A');
        List<String> allTagNames = resMap.get("result");
        List<String> repeatedTagNames = resMap.get("repeated");
        HashMap<String, List<String>> ourTagNameMap = readData(new File("D:\\tmp\\冷轧测点302+PI.xlsx"), 1, 'N' - 'A');
        List<String> ourTagNames = ourTagNameMap.get("result");
        List<String> commonList = new ArrayList<>();
        List<String> diffList = new ArrayList<>();
        for (String tagName:allTagNames) {
            if (ourTagNames.contains(tagName)) {
                commonList.add(tagName);
            }else {
                diffList.add(tagName);
            }
        }
        File file = new File("D:\\tmp\\宝信整理的点表\\冷轧\\与我们系统的冷轧测点对比结果\\比较结果.xlsx");
        XSSFWorkbook xwb = new XSSFWorkbook();
        Sheet sheet = xwb.createSheet("比较结果");
        Row headRow = sheet.createRow(0);
        headRow.createCell(0).setCellValue("系统的TagName(" + ourTagNames.size() + ")");
        headRow.createCell(1).setCellValue("相同的TagName(" + commonList.size() + ")");
        headRow.createCell(2).setCellValue("不同的TagName(" + diffList.size() + ")");
        headRow.createCell(3).setCellValue("重复的TagName(" + repeatedTagNames.size() + ")");

        for (int i = 0; i < ourTagNames.size(); i++) {
            sheet.createRow(i+1).createCell(0).setCellValue(ourTagNames.get(i));
        }
        for (int i = 0; i < commonList.size(); i++) {
            sheet.getRow(i+1).createCell(1).setCellValue(commonList.get(i));
        }
        for (int i = 0; i < diffList.size(); i++) {
            sheet.getRow(i+1).createCell(2).setCellValue(diffList.get(i));
        }
        for (int i = 0; i < repeatedTagNames.size(); i++) {
            sheet.getRow(i+1).createCell(3).setCellValue(repeatedTagNames.get(i));
        }
        CompareData.setAutoWith(sheet);
        xwb.write(Files.newOutputStream(file.toPath()));
        xwb.close();
    }
}
