package com.baixiaowen.demoexcelutil.easyexcel;
import	java.text.SimpleDateFormat;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.baixiaowen.demoexcelutil.poiusermodel.TestService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RestController
public class EasyTestController {
    
    @Autowired
    private EasyExcelService easyExcelService;
    
    
    public String test(HttpServletResponse response){
        try {
            return easyExcelService.writeExcel(response);
        } catch (IOException e) {
            e.printStackTrace();
            return "失败";
        }
    }

    @GetMapping("/cooperation")
    public void cooperation(HttpServletRequest request, HttpServletResponse response){
        ServletOutputStream out = null;
        try {
            out = response.getOutputStream();
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
            String fileName = new String(("UserInfo " + new SimpleDateFormat("yyyy-MM-dd").format(new Date()))
                    .getBytes(), "UTF-8");
            Sheet sheet1 = new Sheet(1, 0);
            sheet1.setSheetName("第一个sheet");
            writer.write0(getListString(), sheet1);
            writer.finish();
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("UTF-8");
            response.setHeader("Content-disposition", "attachment;filename="+fileName+".xlsx");
            out.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    // 创建数据模型， 正常业务这块是从数据库中查询出来的
    private List<WriteModel> getListString() {
        List<WriteModel> writeModels = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            WriteModel writeModel = WriteModel.builder()
                    .name("小白学Java" + i).password("123456").age(i+1).build();
            writeModels.add(writeModel);
        }
        return writeModels;
    }
    
}
