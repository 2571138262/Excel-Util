package com.baixiaowen.demoexcelutil.easyexcel;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RestController
public class WebExcel {


    @GetMapping("/a.htm")
    public void cooperation(HttpServletRequest request, HttpServletResponse response) throws Exception {
        ServletOutputStream out = response.getOutputStream();
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
        String fileName = new String(("UserInfo " + new SimpleDateFormat("yyyy-MM-dd").format(new Date())).getBytes(), "UTF-8");
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
        response.setContentType("multipart/form-data");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        out.flush();
    }

    /**
     * 无注解实体类
     * 创建动态数据，注意这里的数据类型是 Object:
     *
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
