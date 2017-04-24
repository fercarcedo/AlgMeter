package fergaral.algmeter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class used to produce an Excel spreadsheet
 * and graph of an algorithm execution time
 * 
 * @author fercarcedo
 * 
 */
public class Renderer {
	private Map<Long, Long> data;
	private String filename;

	public Renderer(Map<Long, Long> data, String filename) {
		this.data = data;
		this.filename = filename;
	}

	/**
	 * Generates a spreadsheet and graph of the 
	 * algorithm execution times
	 * 
	 */
	public void render() throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet worksheet = workbook.createSheet(filename);

		createHeaders(worksheet);
		renderData(worksheet);
		createGraph(worksheet);
		save(workbook);
	}
	
	/**
	 * Creates headers for the data table
	 */
	protected void createHeaders(XSSFSheet worksheet) {
		XSSFRow headerRow = worksheet.createRow(0);
		SpreadsheetUtils.createCell(headerRow, 0, "n");
		SpreadsheetUtils.createCell(headerRow, 1, "time");
	}
	
	/**
	 * Renders algorithm times into a table
	 */
	protected void renderData(XSSFSheet worksheet) {
		int row = 1;
		
		for (Map.Entry<Long, Long> dataEntry : data.entrySet()) {
			XSSFRow entryRow = worksheet.createRow(row++);
			SpreadsheetUtils.createCell(entryRow, 0, dataEntry.getKey());
			SpreadsheetUtils.createCell(entryRow, 1, dataEntry.getValue());
		}	
	}
	
	/**
	 * Renders a graph representing the algorithm workload / time
	 */
	protected void createGraph(XSSFSheet worksheet) {
        XSSFChart chart = SpreadsheetUtils.createChart(worksheet, 
        		new XSSFClientAnchor(0, 0, 0, 0, 3, 1, 13, 11));
        
        SpreadsheetUtils.drawGraph(worksheet, chart, "n", "time", 
        		new CellRangeAddress(1, data.size(), 0, 0), 
        		new CellRangeAddress(1, data.size(), 1, 1));
	}
	
	/**
	 * Saves the resulting file
	 * 
	 * @param workbook current document
	 * @throws IOException
	 */
	protected void save(Workbook workbook) throws IOException {
		FileOutputStream fileOut = new FileOutputStream(filename + ".xlsx");
		workbook.write(fileOut);
		fileOut.close();
		workbook.close();
	}
}
