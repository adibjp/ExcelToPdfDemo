package com.demo;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

import com.lowagie.text.Document;
import com.lowagie.text.pdf.DefaultFontMapper;
import com.lowagie.text.pdf.PdfContentByte;
import com.lowagie.text.pdf.PdfTemplate;
import com.lowagie.text.pdf.PdfWriter;

public class Demo {
	public static void main(String[] args) {
		Map<String, Double> barChartDataPoints = readFromExcel("//Users//adityabajpai//Documents//test2/abc.xlsx");
		writeChartToPDF(generateBarChart(barChartDataPoints), 500, 400, "//Users//adityabajpai//Documents//test2/test1.pdf");

	}

	public static Map<String, Double> readFromExcel(String path) {
		Map<String, Double> mp = new TreeMap<String, Double>();
		try {
			FileInputStream file = new FileInputStream(new File(path));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				mp.put(row.getCell(0).getStringCellValue(), row.getCell(1).getNumericCellValue());

			}
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return mp;

	}

//	public static JFreeChart generatePieChart() {
//		DefaultPieDataset dataSet = new DefaultPieDataset();
//		JFreeChart chart = ChartFactory.createPieChart("Test", dataSet, true, true, false);
//
//		return chart;
//	}

	public static JFreeChart generateBarChart(Map<String, Double> dataPoints) {
		DefaultCategoryDataset dataSet = new DefaultCategoryDataset();
        for (Map.Entry<String,Double> point : dataPoints.entrySet())  {
    		dataSet.setValue(point.getValue(), "Expense", point.getKey());
        	
        }
		JFreeChart chart = ChartFactory.createBarChart("Demo label", "Expense", "Brand",
				dataSet, PlotOrientation.VERTICAL, false, true, false);

		return chart;
	}

	public static void writeChartToPDF(JFreeChart chart, int width, int height, String fileName) {
		PdfWriter writer = null;

		Document document = new Document();

		try {
			writer = PdfWriter.getInstance(document, new FileOutputStream(fileName));
			document.open();
			PdfContentByte contentByte = writer.getDirectContent();
			PdfTemplate template = contentByte.createTemplate(width, height);
			Graphics2D graphics2d = template.createGraphics(width, height, new DefaultFontMapper());
			Rectangle2D rectangle2d = new Rectangle2D.Double(0, 0, width, height);
			chart.draw(graphics2d, rectangle2d);
			graphics2d.dispose();
			contentByte.addTemplate(template, 0, 0);

		} catch (Exception e) {
			e.printStackTrace();
		}
		document.close();
	}
}
