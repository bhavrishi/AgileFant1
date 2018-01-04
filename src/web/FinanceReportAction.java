package fi.hut.soberit.agilefant.web;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.EmptyStackException;
import java.util.Iterator;
import java.util.List;

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

@Component("financeReportAction")
@Scope("prototype")
public class FinanceReportAction extends ActionSupport {
	private InputStream inputStream;
	
	public String getReports() throws HPSFException, SQLException{
		ArrayList finaceList = getFinanceReport();
		ArrayList fundsList = getAllotedFundsReport();
		exportToExcel((ArrayList)finaceList.get(0), (ArrayList)finaceList.get(1), (ArrayList)fundsList.get(0), (ArrayList)fundsList.get(1));
		return Action.SUCCESS;
	}

	public ArrayList getFinanceReport() throws HPSFException, SQLException {
		Connection conn = null;
		Statement stmt = null;
		ResultSet rs = null;
		conn = DBConnectivity.dataBaseConnect();
		ArrayList data = new ArrayList();		
		stmt = conn.createStatement();
		String sql = "select hr.user_id, us.initials, ta.name, ta.state, sum(hr.minutesSpent), ta.originalestimate, us.cost from hourentries hr join users us on us.id = hr.user_id join tasks ta on ta.id = hr.task_id group by task_id;";
		rs = stmt.executeQuery(sql);
		while(rs.next()){
			ArrayList cells = new ArrayList();
			cells.add(rs.getString("user_id"));
			cells.add(rs.getString("us.initials"));
			cells.add(rs.getString("ta.name"));
			String state = rs.getString("ta.state");
			if(Integer.parseInt(state) == 5){
				state = "Completed";
			} else if (Integer.parseInt(state) == 1){
				state = "InProgress";
			}else if (Integer.parseInt(state) == 0){
				state = "NotStarted";
			}
			cells.add(state);
			cells.add(String.valueOf(Integer.parseInt(rs.getString("sum(hr.minutesSpent)"))/60));
			cells.add(rs.getString("cost"));
			cells.add(String.valueOf((Integer.parseInt(rs.getString("sum(hr.minutesSpent)")))*(Integer.parseInt(rs.getString("cost")))/60));
			cells.add(String.valueOf(Integer.parseInt(rs.getString("originalestimate"))/60));	
			cells.add(String.valueOf(((Integer.parseInt(rs.getString("originalestimate"))/60) - Integer.parseInt((rs.getString("sum(hr.minutesSpent)")))/60)*(Integer.parseInt(rs.getString("cost")))));
			data.add(cells);
		}
		ArrayList headers = new ArrayList();
		headers.add("UserID");
		headers.add("UserName");
		headers.add("TaskName");
		headers.add("TaskStatus");
		headers.add("EffortSpent (Hrs)");
		headers.add("Cost/hr (USD)");
		headers.add("MoneySpent (USD)");
		headers.add("AssignedHours (Hrs)");
		headers.add("ResidualAmount (USD)");
		//exportToExcel(headers, data);
		ArrayList hd = new ArrayList();
		hd.add(headers);
		hd.add(data);
		System.out.println("Executed getFinanceReport...");
		return hd;
	}
	
	public ArrayList getAllotedFundsReport() throws HPSFException, SQLException {
		Connection conn = null;
		Statement stmt = null;
		ResultSet rs = null;
		conn = DBConnectivity.dataBaseConnect();
		ArrayList data = new ArrayList();		
		stmt = conn.createStatement();
		String sql = "select us.id, us.initials, ta.name, ta.state, ta.originalestimate, us.cost from users us join task_user tu on us.id = tu.responsibles_id join tasks ta on ta.id = tu.tasks_id;";
		rs = stmt.executeQuery(sql);
		while(rs.next()){
			ArrayList cells = new ArrayList();
			cells.add(rs.getString("us.id"));
			cells.add(rs.getString("us.initials"));
			cells.add(rs.getString("ta.name"));
			String state = rs.getString("ta.state");
			if(Integer.parseInt(state) == 5){
				state = "Completed";
			} else if (Integer.parseInt(state) == 1){
				state = "InProgress";
			}else if (Integer.parseInt(state) == 0){
				state = "NotStarted";
			}
			cells.add(state);
			cells.add(String.valueOf(Integer.parseInt(rs.getString("originalestimate"))/60));
			cells.add(rs.getString("cost"));
			cells.add(String.valueOf((Integer.parseInt(rs.getString("originalestimate")))*(Integer.parseInt(rs.getString("cost")))/60));	
			data.add(cells);
		}
		ArrayList headers = new ArrayList();
		headers.add("UserID");
		headers.add("UserName");
		headers.add("TaskName");
		headers.add("TaskStatus");
		headers.add("AssignedHours (Hrs)");
		headers.add("Cost/hr (USD)");
		headers.add("AllotedFunds (USD)");
		//exportToExcel(headers, data);
		ArrayList hd = new ArrayList();
		hd.add(headers);
		hd.add(data);
		System.out.println("Executed getAllotedFunds Method...");
		return hd;
	}

	public void exportToExcel(ArrayList headers, ArrayList data, ArrayList headersn, ArrayList datan) throws HPSFException {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("ForecastReport");
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
		//wb.setSheetName(0, sheetName);
		HSSFSheet sheet1 = wb.createSheet("FinanceReport");
		int rowIdxn = 0;
		short cellIdxn = 0;

		// Header
		HSSFRow hssfHeader1 = sheet1.createRow(rowIdxn);
		HSSFCellStyle cellStyle1 = wb.createCellStyle();
		cellStyle1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		for (Iterator cells = headersn.iterator(); cells.hasNext();) {
			HSSFCell hssfCell = hssfHeader1.createCell(cellIdxn++);
			hssfCell.setCellStyle(cellStyle);
			hssfCell.setCellValue((String) cells.next());
		}
		// Data
		rowIdxn = 1;
		for (Iterator rows = datan.iterator(); rows.hasNext();) {
			ArrayList row = (ArrayList) rows.next();
			HSSFRow hssfRow = sheet1.createRow(rowIdxn++);
			cellIdxn = 0;
			for (Iterator cells = row.iterator(); cells.hasNext();) {
				HSSFCell hssfCell = hssfRow.createCell(cellIdxn++);
				hssfCell.setCellValue((String) cells.next());
			}
		}
		//wb.setSheetName(0, sheetName);
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
