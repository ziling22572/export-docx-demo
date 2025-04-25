package org.example.exportdocxdemo.util;

/**
 * @author 自己的姓名或者昵称
 * @version v1.0.0
 * @Package : org.example.exportdocxdemo
 * @Description : TODO
 * @Create on : 2025/4/24 17:31
 **/

import org.knowm.xchart.*;
import java.awt.image.BufferedImage;
import java.util.*;

public class ChartGenerator {

    public static BufferedImage generateLineChart(Map<String, Double> data) {
        List<Integer> xData = new ArrayList<>();
        List<Double> yData = new ArrayList<>();
        int index = 1;
        for (Double value : data.values()) {
            xData.add(index++);
            yData.add(value);
        }

        XYChart chart = new XYChartBuilder()
                .width(500)
                .height(300)
                .title("销售趋势图（按月序）")
                .xAxisTitle("月份序号")
                .yAxisTitle("销售额")
                .build();
        chart.addSeries("销量", xData, yData);

        return BitmapEncoder.getBufferedImage(chart);
    }
}
