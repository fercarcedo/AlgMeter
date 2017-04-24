package fergaral.algmeter;

import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

/**
 * Utils related to spreadsheets and 
 * graphs within them
 * 
 * @author fercarcedo
 *
 */
public class SpreadsheetUtils {
	/**
	 * Creates a cell in the document with the given position and text
	 * 
	 * @param row cell row
	 * @param colIndex cell column
	 * @param value cell content
	 * @return created cell
	 */
	public static XSSFCell createCell(XSSFRow row, int colIndex, String value) {
		XSSFCell cell = row.createCell(colIndex);
		cell.setCellValue(value);
		return cell;
	}
	
	/**
	 * Creates a cell in the document with the given position and value
	 * 
	 * @param row cell row
	 * @param colIndex cell column
	 * @param value cell numerical content
	 * @return created cell
	 */
	public static XSSFCell createCell(XSSFRow row, int colIndex, long value) {
		XSSFCell cell = row.createCell(colIndex);
		cell.setCellValue(value);
		return cell;
	}
	
	/**
	 * Sets category axis title
	 * 
	 * @param chart graph
	 * @param axisIdx axis id
	 * @param title title of the axis
	 */
	public static void setCatAxisTitle(XSSFChart chart, int axisIdx, String title) {
	    CTCatAx valAx = chart.getCTChart().getPlotArea().getCatAxArray(axisIdx);
	    CTTitle ctTitle = valAx.addNewTitle();
	    ctTitle.addNewLayout();
	    ctTitle.addNewOverlay().setVal(false);
	    CTTextBody rich = ctTitle.addNewTx().addNewRich();
	    rich.addNewBodyPr();
	    rich.addNewLstStyle();
	    CTTextParagraph p = rich.addNewP();
	    p.addNewPPr().addNewDefRPr();
	    p.addNewR().setT(title);
	    p.addNewEndParaRPr();
	}

	/**
	 * Sets value axis title
	 * 
	 * @param chart graph
	 * @param axisIdx axis id
	 * @param title title of the axis
	 */
	public static void setValueAxisTitle(XSSFChart chart, int axisIdx, String title) {
	    CTValAx valAx = chart.getCTChart().getPlotArea().getValAxArray(axisIdx);
	    CTTitle ctTitle = valAx.addNewTitle();
	    ctTitle.addNewLayout();
	    ctTitle.addNewOverlay().setVal(false);
	    CTTextBody rich = ctTitle.addNewTx().addNewRich();
	    rich.addNewBodyPr();
	    rich.addNewLstStyle();
	    CTTextParagraph p = rich.addNewP();
	    p.addNewPPr().addNewDefRPr();
	    p.addNewR().setT(title);
	    p.addNewEndParaRPr();
	}
	
	/**
	 * Creates a chart in the given worksheet
	 * and positioned by anchor
	 * 
	 * @param worksheet current document
	 * @param anchor chart position
	 * @return created chart
	 */
	public static XSSFChart createChart(XSSFSheet worksheet, XSSFClientAnchor anchor){
        XSSFDrawing drawing = worksheet.createDrawingPatriarch();
        // Define anchor points in the worksheet to position the chart 
        return drawing.createChart(anchor);
	}

	public static void drawGraph(XSSFSheet worksheet, XSSFChart chart, String catAxisTitle, String valueAxisTitle, 
			CellRangeAddress nValuesRange, CellRangeAddress timeValuesRange) {
		ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);   
        
        SpreadsheetUtils.setCatAxisTitle(chart, 0, catAxisTitle);
        SpreadsheetUtils.setValueAxisTitle(chart, 0, valueAxisTitle);
        
        // Pass data into the chart
        // Cell Range Address is defined as First row, last row, first column, last column 
        ChartDataSource<Number> nValues = DataSources.fromNumericCellRange(worksheet, nValuesRange);
        ChartDataSource<Number> timeValues = DataSources.fromNumericCellRange(worksheet, timeValuesRange);

		// Create data for the chart 
        LineChartData chartData = chart.getChartDataFactory().createLineChartData();   
        chartData.addSeries(nValues, timeValues);
       
        chart.plot(chartData, new ChartAxis[] { bottomAxis, leftAxis });
	}
}
