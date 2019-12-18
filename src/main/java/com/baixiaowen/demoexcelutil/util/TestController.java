package com.baixiaowen.demoexcelutil.util;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;

@RestController
public class TestController {
    
    @Autowired
    private TestService testService;
    
    @GetMapping("/test")
    public String test(HttpServletResponse response){
        return testService.excelExport(response);
    }
    
}
