package fi.hut.soberit.agilefant.web;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.hpsf.HPSFException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import com.opensymphony.xwork2.Action;
import com.opensymphony.xwork2.ActionSupport;

import fi.hut.soberit.agilefant.config.DBConnectivity;

//@Component("varianceReportAction")
@Scope("prototype")
public class VarianceReportAction extends ActionSupport{
	private InputStream inputStream;

	public String getVarianceReport() throws HPSFException, SQLException {
		Connection conn = null;
		Statement stmt = null;
		ResultSet rs = null;
		conn = DBConnectivity.dataBaseConnect();
		ArrayList data = new ArrayList();		
		stmt = conn.createStatement();
		String sql = "select hr.task_id, ta.name, sum(hr.minutesSpent), ta.originalestimate from hourentries hr, tasks ta where ta.id = hr.task_id group by task_id;";
		rs = stmt.executeQuery(sql);
		while(rs.next()){
			ArrayList cells = new ArrayList();
			cells.add(rs.getString("task_id"));
			cells.add(rs.getString("name"));
			cells.add(String.valueOf(Integer.parseInt(rs.getString("sum(hr.minutesSpent)"))/60));
			cells.add(String.valueOf(Integer.parseInt(rs.getString("originalestimate"))/60));
			cells.add(String.valueOf((Integer.parseInt(rs.getString("originalestimate"))/60) - Integer.parseInt((rs.getString("sum(hr.minutesSpent)")))/60));
			data.add(cells);
			System.out.println(data.toString());
		}
		ArrayList headers = new ArrayList();
		headers.add("TaskID");
		headers.add("TaskName");
		headers.add("Effort Spent (Hrs)");
		headers.add("Original Estimate (Hrs)");
		headers.add("Variance (Hrs)");
		exportToExcel("VarianceReport",headers, data);
		return Action.SUCCESS;
	}

	public void exportToExcel(String sheetName, ArrayList headers, ArrayList data) throws HPSFException {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet();
		int rowIdx = 0;
		short cellIdx = 0;

		// Header
		HSSFRow hssfHeader = sheet.createRow(rowIdx);
		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		for (Iterator cells = headers.iterator(); cells.hasNext();) {
			HSSFCell hssfCell = hssfHeader.createCell(cellIdx++);
			hssfCell.setCellStyle(cellStyle);
			hssfCell.setCellValue((String) cells.next());
		}
		// Data
		rowIdx = 1;
		for (Iterator rows = data.iterator(); rows.hasNext();) {
			ArrayList row = (ArrayList) rows.next();
			HSSFRow hssfRow = sheet.createRow(rowIdx++);
			cellIdx = 0;
			for (Iterator cells = row.iterator(); cells.hasNext();) {
				HSSFCell hssfCell = hssfRow.createCell(cellIdx++);
				hssfCell.setCellValue((String) cells.next());
			}
		}

		wb.setSheetName(0, sheetName);
		try {
			ByteArrayOutputStream outs = new ByteArrayOutputStream();
			wb.write(outs);
			setInputStream(new ByteArrayInputStream(outs.toByteArray()));
			outs.close();
		} catch (IOException e) {
			throw new HPSFException(e.getMessage());
		}
	}

	public InputStream getInputStream() {
		return inputStream;
	}

	public void setInputStream(InputStream inputStream) {
		this.inputStream = inputStream;
	}
}
