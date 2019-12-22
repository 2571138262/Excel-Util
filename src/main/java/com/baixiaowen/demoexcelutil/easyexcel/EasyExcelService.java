package com.baixiaowen.demoexcelutil.easyexcel;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

@Service
public class EasyExcelService {
    
    public String writeExcel(HttpServletResponse response) throws IOException {
        // 文件名
        String fileName = "白晓文.xlsx";

        OutputStream out = response.getOutputStream(); 
        
        ExcelWriter writer = EasyExcelFactory.getWriter(out);

        // 写仅有一个 Sheet 的 Excel 文件， 此场景较为通用
        Sheet sheet1 = new Sheet(1, 0, WriteModel.class) ;
        
        // 第一个 sheet 名称
        sheet1.setSheetName("第一个sheet");
        
        // 写数据到 Write 上下文
        // 入参1：创建要写入的模型数据
        // 入参2：要写入的目标 sheet
        writer.write(createModelList(), sheet1);
        
        // 将上下文中的最终 outputStream 写入到指定文件中
        writer.finish();
        
        // 响应
        this.setResponseHeader(response, fileName);
        
        // 关闭流
        out.close();
        
        return "成功";
    }

    // 创建数据模型， 正常业务这块是从数据库中查询出来的
    private List<WriteModel> createModelList() {
        List<WriteModel> writeModels = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            WriteModel writeModel = WriteModel.builder()
                    .name("小白学Java" + i).password("123456").age(i+1).build();
            writeModels.add(writeModel);
        }
        return writeModels;
    }

    /**
     * 设置响应
     * @param response
     * @param fileName
     */
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
