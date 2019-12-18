package com.baixiaowen.demoexcelutil.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;

@Service
public class TestService {

    public String excelExport(HttpServletResponse response) {
        String fileName = "测试Excel.xls";
        String sheetName = "测试sheet";

        String[] title = new String[]{"标题1", "标题2", "标题3", "标题4", "标题5", "标题6", "标题7"};

        String[][] values = {
                {"1行1列", "1行2列", "1行3列", "1行4列", "1行5列", "1行6列", "1行7列"},
                {"2行1列", "2行2列", "2行3列", "2行4列", "2行5列", "2行6列", "2行7列"},
                {"3行1列", "3行2列", "3行3列", "3行4列", "3行5列", "3行6列", "3行7列"},
                {"4行1列", "4行2列", "4行3列", "4行4列", "4行5列", "4行6列", "4行7列"},
                {"5行1列", "5行2列", "5行3列", "5行4列", "5行5列", "5行6列", "5行7列"},
                {"6行1列", "6行2列", "6行3列", "6行4列", "6行5列", "6行6列", "6行7列"},
                {"7行1列", "7行2列", "7行3列", "7行4列", "7行5列", "7行6列", "7行7列"},
                {"8行1列", "8行2列", "8行3列", "8行4列", "8行5列", "8行6列", "8行7列"},
                {"9行1列", "9行2列", "9行3列", "9行4列", "9行5列", "9行6列", "9行7列"},
                {"10行1列", "10行2列", "10行3列", "10行4列", "10行5列", "10行6列", "10行7列"}
        };

        HSSFWorkbook wb = ExcelUtil.getHSSFWorkbook(sheetName, title, values, null);

        OutputStream os = null;
        try {
            this.setResponseHeader(response, fileName);
            os = response.getOutputStream();
            wb.write(os);
            os.flush();
            os.close();
            return "下载成功";
        } catch (Exception e) {
            System.out.println(e.getMessage());
            return "下载失败";
        } finally {
            try {
                os.flush();
                os.close();
            } catch (IOException e) {
                System.out.println(e.getMessage());
            }
        }
    }

    private void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            fileName = new String(fileName.getBytes(), "ISO8859-1");
        } catch (UnsupportedEncodingException e) {
            System.out.println(e.getMessage());
        }
        response.setContentType("application/vnd.ms-excel;charset=OSI8859-1");
        response.setHeader("Content-Disposition", "attachment;filename" + fileName);
        response.addHeader("Pargam", "no-cache");
        response.addHeader("Cache-Control", "no-cache");
    }

}
