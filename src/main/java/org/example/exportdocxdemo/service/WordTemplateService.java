package org.example.exportdocxdemo.service;


import org.apache.poi.xwpf.usermodel.*;
import org.example.exportdocxdemo.util.ChartGenerator;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
/**
 * @author 自己的姓名或者昵称
 * @version v1.0.0
 * @Package : org.example.exportdocxdemo
 * @Description : TODO
 * @Create on : 2025/4/24 17:22
 **/
@Service
public class WordTemplateService {
    public void exportToWord(String templatePath,
                             Map<String, String> textPlaceholders,
                             List<List<String>> tableData,
                             Map<String, Double> chartData,
                             HttpServletResponse response) throws Exception {

        try (InputStream is = getClass().getResourceAsStream(templatePath);
             XWPFDocument doc = new XWPFDocument(is)) {

            replaceTextPlaceholders(doc, textPlaceholders);
            insertTable(doc, tableData);
            insertChartImage(doc, chartData);

            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.setHeader("Content-Disposition", "attachment; filename=report.docx");

            try (OutputStream os = response.getOutputStream()) {
                doc.write(os);
            }
        }
    }

    private void replaceTextPlaceholders(XWPFDocument doc, Map<String, String> textPlaceholders) {
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                if (text != null) {
                    for (Map.Entry<String, String> entry : textPlaceholders.entrySet()) {
                        if (text.contains(entry.getKey())) {
                            text = text.replace(entry.getKey(), entry.getValue());
                            run.setText(text, 0);
                        }
                    }
                }
            }
        }
    }

    private void insertTable(XWPFDocument doc, List<List<String>> tableData) {
        XWPFTable table = doc.createTable();
        for (List<String> rowData : tableData) {
            XWPFTableRow row = table.createRow();
            for (String cell : rowData) {
                row.createCell().setText(cell);
            }
        }
        table.removeRow(0); // remove the first default row
    }

    private void insertChartImage(XWPFDocument doc, Map<String, Double> chartData) throws Exception {
        BufferedImage chartImage = ChartGenerator.generateLineChart(chartData);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        ImageIO.write(chartImage, "png", os);

        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.addPicture(new ByteArrayInputStream(os.toByteArray()),
                Document.PICTURE_TYPE_PNG, "chart.png", 500 * 9525, 300 * 9525);
    }
}
