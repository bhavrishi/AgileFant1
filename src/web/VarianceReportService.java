package fi.hut.soberit.agilefant.web;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
/*import org.apache.poi.hpsf.HPSFException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;*/
import java.util.Map;

import org.apache.poi.hpsf.HPSFException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import com.opensymphony.xwork2.Action;
import com.opensymphony.xwork2.ActionSupport;

import fi.hut.soberit.agilefant.config.DBConnectivity;
import fi.hut.soberit.agilefant.db.VarianceReportDao;
import fi.hut.soberit.agilefant.model.Expense;
import fi.hut.soberit.agilefant.model.VarReport;

/**
 * Service class to hold state of Expense at project level. This class has
 * functions to facilitate generation of Report at Project level.
 *
 */

@Component("varianceReportAction")
@Scope("prototype")
public class VarianceReportService extends ActionSupport {
	private InputStream inputStream;

	@Autowired
	private VarianceReportDao varianceReportDao;

	public String getVarianceReport() throws HPSFException, SQLException {

		ArrayList<VarReport> data = new ArrayList<VarReport>();

		HashMap<String, Expense> projectMap = new HashMap<String, Expense>();

		projectMap = varianceReportDao.getDetails();
		ArrayList dataWritten = new ArrayList();
		Double difference;

		// Write new Excel workbook (Use new Apache POI jar):
		ArrayList<String> headers = new ArrayList<String>();
		headers.add("Project Name");
		headers.add("Project Forecast");
		headers.add("Project Actuals");
		headers.add("%difference");

		// Write Data
		// New Map traversal:
		Iterator it = projectMap.entrySet().iterator();
		while (it.hasNext()) {
			Map.Entry map = (Map.Entry) it.next();
			System.out.println(map.getKey() + " = " + map.getValue());
			difference = (double) ((Math.abs(
					((Expense) map.getValue()).getEstimatedExpense() - ((Expense) map.getValue()).getActualExpense()))
					* 100);
			Integer estimate = (int) ((Expense) map.getValue()).getEstimatedExpense();
			Integer actual = (int) ((Expense) map.getValue()).getActualExpense();
			difference = difference / estimate;
			ArrayList cells = new ArrayList();
			cells.add(map.getKey());
			// Converting to String
			cells.add(estimate.toString());
			cells.add(actual.toString());
			cells.add(difference.toString());
			dataWritten.add(cells);

		}

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet();
		int rowIdx = 0;
		short cellIdx = 0;

		// Write Report Header:
		HSSFRow hssfHeader = sheet.createRow(rowIdx);
		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		for (String value : headers) {
			HSSFCell hssfCell = hssfHeader.createCell(cellIdx++);
			hssfCell.setCellStyle(cellStyle);
			hssfCell.setCellValue((String) value);
		}

		// Write Report Data:
		rowIdx = 1;
		for (Iterator rows = dataWritten.iterator(); rows.hasNext();) {
			ArrayList row = (ArrayList) rows.next();
			HSSFRow hssfRow = sheet.createRow(rowIdx++);
			cellIdx = 0;
			for (Iterator cells = row.iterator(); cells.hasNext();) {
				HSSFCell hssfCell = hssfRow.createCell(cellIdx++);
				hssfCell.setCellValue((String) cells.next());
				// System.out.print((String) cells.next());
			}
		}

		wb.setSheetName(0, "VarianceReport");

		try {
			ByteArrayOutputStream outs = new ByteArrayOutputStream();
			wb.write(outs);
			setInputStream(new ByteArrayInputStream(outs.toByteArray()));
			outs.close();
		} catch (IOException e) {
			throw new HPSFException(e.getMessage());
		}

		return Action.SUCCESS;
	}

	public InputStream getInputStream() {
		return inputStream;
	}

	public void setInputStream(InputStream inputStream) {
		this.inputStream = inputStream;
	}

}
