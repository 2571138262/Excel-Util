# EasyExcel
原文链接 ： https://mp.weixin.qq.com/s/SXxXgrNZbbY1R-VV6jQ1IA
### 概念：
##### 什么是 EasyExcel? 见名知意，就是让你操作 Excel 异常的酸爽。 EasyExcel GitHub 官方 （代码托管）
### EasyExcel 解决了什么
##### 1、传统 Excel 框架，如 Apache poi、jxl 都存在内存溢出的问题；
##### 2、传统 excel 开源框架使用复杂、繁琐;
##### 3、EasyExcel 底层还是使用了 poi, 但是做了很多优化，如修复了并发情况下的一些 bug, 具体修复细节，可阅读官方文档https://github.com/alibaba/easyexcel；

### 快速上手
##### 1、添加依赖
    <!-- https://mvnrepository.com/artifact/com.alibaba/easyexcel -->
        <dependency>
            <groupId>com.alibaba</groupId>
            <artifactId>easyexcel</artifactId>
            <version>2.1.4</version>
        </dependency>
##### 2、七行代码搞定 Excel 生成
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
##### 上面这段示例代码中，有两个点很重要，小哈已经重点标注标：
###### ①：WriteModel 这个对象就是要写入 Excel 的数据模型对象，等等，你这好像不行吧？表头 head，以及每个单元格内的数据顺序都没指定，能达到想要的效果么？别急，后面会讨论这块！
###### ②：创建需要写入的数据集，当然了，正常业务中，这块都是从数据库中查询出来的。
##### PS: 如果说写入的数据量很大，需要做分片查询再写入的处理，否则可能会 OOM（Out of Memory）.
##### WriteModel
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

##### ExayExcel 提供注解的方式, 来方便的定义 Excel 需要的数据模型：
###### ①：首先，定义的写入模型必须要继承自 BaseRowModel.java;
###### ②：通过 @ExcelProperty 注解来指定每个字段的列名称，以及下标位置；
##### 同时，上面定义的 createModelList() 方法也很简单，通过循环，创建一个写入模型的 List 集合：

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
        
##### 这里可能会出现错误
    Exception in thread "main" com.alibaba.excel.exception.ExcelGenerateException: Can not close IO.
##### 参考 https://www.cnblogs.com/zhangzhonghui/articles/12020630.html
### 特殊场景支持
##### 1、动态生成 Excel 内容
上面的例子是基于注解的，也就是说表头 head, 以及内容都是写死的，换句话说，我定义好了一个数据模型，那么，生成的 Excel 文件也就是只能遵循这种模型来了，但是，实际业务中可能会存在动态变化的需求，要怎么做呢？
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
##### 2、动态生成 Excel 内容


    
