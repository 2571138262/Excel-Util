package com.baixiaowen.demoexcelutil.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

public class ExcelUtil {
    
    private static ExcelUtil instance = new ExcelUtil();
    
    private ExcelUtil(){}
    
    public static ExcelUtil getInstance() {
        return instance;
    }
    
    public static HSSFWorkbook getHSSFWorkbook(String sheetName, String[] title, String[][] values, HSSFWorkbook wb){
        if (wb == null){
            wb = new HSSFWorkbook();
        }

        HSSFSheet sheet = wb.createSheet(sheetName);
        HSSFRow row = sheet.createRow(0);
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.forInt(2));
        HSSFCell cell = null;
        
        int i;
        for (i = 0; i < title.length; ++i) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        for (i = 0; i < values.length; ++i) {
            row = sheet.createRow(i + 1);

            for (int j = 0; j < values[i].length; j++) {
                row.createCell(j).setCellValue(values[i][j]);
            }
        }
        return wb;
    }
    
}
