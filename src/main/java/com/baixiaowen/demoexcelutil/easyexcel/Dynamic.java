package com.baixiaowen.demoexcelutil.easyexcel;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 动态生成Excel内容
 */
public class Dynamic {

    /**
     * 
     * @throws Exception
     */
    @Test
    public void writeExcel2() throws Exception {
        // 文件输出位置
        OutputStream out = new FileOutputStream("动态生成Excel内容.xlsx");

        ExcelWriter writer = EasyExcelFactory.getWriter(out);
        
        // 动态添加表头，使用一些表头动态变化的场景
        Sheet sheet1 = new Sheet(1, 0);
        
        sheet1.setSheetName("第一个sheet");
        
        // 创建一个表格， 用于 Sheet 中使用
        Table table1 = new Table(1);
        
        // 无注解的模式，动态添加表头
        table1.setHead(DataUtil.createTestListStringHead());
        // 写数据
        writer.write1(createDynamicModelList(), sheet1, table1);
        
        // 将上下文中的最终 outputStream 写入到指定文件中
        writer.finish();
        
        // 关闭流
        out.close();
    }

    /**
     * 无注解实体类
     * 创建动态数据，注意这里的数据类型是 Object:
     * @return
     */
    private List<List<Object>> createDynamicModelList() {
        // 所有行数据
        List<List<Object>> rows = new ArrayList<>();

        for (int i = 0; i < 100; i++) {
            // 一行数据
            List<Object> row = new ArrayList<>();
            row.add("字符串" + i);
            row.add(Long.valueOf(18554000000L + i));
            row.add(Integer.valueOf(12300 + i));
            row.add("犬小哈");
            row.add("微信公众号");
            rows.add(row);
        }
        return rows;
    }


}
