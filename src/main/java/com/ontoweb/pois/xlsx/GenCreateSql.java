package com.ontoweb.pois.xlsx;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class GenCreateSql {
    public static void main(String[] args) {
        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(new FileInputStream("D:\\项目文件\\疾控\\智能报告系统\\问卷数据库字段.xlsx"));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Row row1 = sheet.getRow(1);
        int num = row.getLastCellNum();
        StringBuilder sb = new StringBuilder();
        sb.append("CREATE TABLE cdc_questionnaire(\n" +
                "`id` varchar(36) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NOT NULL COMMENT '主键',\n" +
//                "PRIMARY KEY (`id`) USING BTREE," +
                "`create_time` datetime COMMENT '创建时间',\n" +
                "`update_time` datetime COMMENT '更新时间',\n" +
                "`create_by` varchar(50) COMMENT '创建人',\n" +
                "`update_by` varchar(50) COMMENT '更新人',\n");
        for (int i = 0; i < num; i++) {
            Cell cell = row.getCell(i);
            Cell cell1 = row1.getCell(i);
            String col = cell.getStringCellValue();
            String col1 = cell1.getStringCellValue();
            sb.append("`").append(col).append("` VARCHAR(100) COMMENT '"+col1+"'");
            if(i < num-1) sb.append(",\n");
        }
        sb.append(")");
        System.out.println(sb);
    }
}
