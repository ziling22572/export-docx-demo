package org.example.exportdocxdemo.controller;


import org.example.exportdocxdemo.service.WordTemplateService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.*;

/**
 * @author 自己的姓名或者昵称
 * @version v1.0.0
 * @Package : org.example.exportdocxdemo
 * @Description : TODO
 * @Create on : 2025/4/24 17:21
 **/

@RestController
@RequestMapping("/api/export")
public class DocxExportController {

    @Autowired
    private WordTemplateService wordTemplateService;

    @GetMapping("/word")
    public void exportWord(HttpServletResponse response) throws Exception {
        Map<String, String> textPlaceholders = new HashMap<>();
        textPlaceholders.put("${title}", "销售报告");
        textPlaceholders.put("${date}", new Date().toString());


        List<List<String>> tableData = new ArrayList<>();
        tableData.add(Arrays.asList("月份", "销售额"));
        tableData.add(Arrays.asList("1月", "1200"));
        tableData.add(Arrays.asList("2月", "1500"));
        tableData.add(Arrays.asList("3月", "1800"));


        Map<String, Double> chartData = new LinkedHashMap<>();
        chartData.put("1月", 1200.0);
        chartData.put("2月", 1500.0);
        chartData.put("3月", 1800.0);


        wordTemplateService.exportToWord("/templates/temp.docx", textPlaceholders, tableData, chartData, response);
    }
}
