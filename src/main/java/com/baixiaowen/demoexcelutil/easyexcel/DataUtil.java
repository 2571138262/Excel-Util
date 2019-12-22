package com.baixiaowen.demoexcelutil.easyexcel;

import java.util.ArrayList;
import java.util.List;

/**
 * 无注解模式，动态添加表头，也可自由组合复杂表头，代码如下：
 */
public class DataUtil {
    
    public static List<List<String>> createTestListStringHead(){
        // 模型上没有注解，表头数据动态传入
        List<List<String>> head = new ArrayList<>();
        List<String> headCoulumn1 = new ArrayList<>();
        List<String> headCoulumn2 = new ArrayList<>();
        List<String> headCoulumn3 = new ArrayList<>();
        List<String> headCoulumn4 = new ArrayList<>();
        List<String> headCoulumn5 = new ArrayList<>();
        
        headCoulumn1.add("第一列");headCoulumn1.add("第一列");headCoulumn1.add("第一列");
        headCoulumn2.add("第一列");headCoulumn2.add("第一列");headCoulumn2.add("第一列");
        
        headCoulumn3.add("第二列");headCoulumn3.add("第二列");headCoulumn3.add("第二列");
        headCoulumn4.add("第三列");headCoulumn4.add("第三列2");headCoulumn4.add("第三列2");
        headCoulumn5.add("第一列");headCoulumn5.add("第3列");headCoulumn5.add("第4列");
        
        head.add(headCoulumn1);
        head.add(headCoulumn2);
        head.add(headCoulumn3);
        head.add(headCoulumn4);
        head.add(headCoulumn5);
        return head;
    }
    
}
