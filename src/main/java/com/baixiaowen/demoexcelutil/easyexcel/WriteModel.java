package com.baixiaowen.demoexcelutil.easyexcel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
//@ContentRowHeight(34)
//@HeadRowHeight(34)
public class WriteModel extends BaseRowModel{

//    @ColumnWidth(13)
    @ExcelProperty(value = "姓名", index = 0)
    private String name;

//    @ColumnWidth(13)
    @ExcelProperty(value = "密码", index = 1)
    private String password;

//    @ColumnWidth(13)
    @ExcelProperty(value = "年龄", index = 2)
    private Integer age;
    
}
