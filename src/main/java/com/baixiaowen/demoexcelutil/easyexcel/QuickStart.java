package com.baixiaowen.demoexcelutil.easyexcel;
import com.alibaba.excel.EasyExcelFactory;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;

import java.io.FileOutputStream;
import	java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class QuickStart {

    public static void main(String[] args) throws Exception {
        writeExcel1();
    }
    
    public static void writeExcel1() throws Exception{
        
        // 文件输出位置
        OutputStream out = new FileOutputStream("快速开始.xlsx");

        ExcelWriter writer = EasyExcelFactory.getWriter(out);
        
        // 写仅有一个 Sheet 的 Excel 文件， 此场景较为通用
        Sheet sheet1 = new Sheet(1, 0, WriteModel.class);
        
        // 第一个 sheet 名称
        sheet1.setSheetName("第一个sheet");
        
        // 写数据到 Writer 上下文
        // 入参1：创建要写入的模型数据
        // 入参2：要写入的目标 sheet
        writer.write(createModelList(), sheet1);
        
        // 将上下文中的最终 outputStream 写入到指定文件中
        writer.finish();
        
        // 关闭流
        out.close();
    }

    // 创建数据模型， 正常业务这块是从数据库中查询出来的
    private static List<WriteModel> createModelList() {
        List<WriteModel> writeModels = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            WriteModel writeModel = WriteModel.builder()
                    .name("小白学Java" + i).password("123456").age(i+1).build();
            writeModels.add(writeModel);
        }
        return writeModels;
    }
    
}
