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

@Component("testReportAction")
@Scope("prototype")
public class TestReportAction {
	private InputStream inputStream;

	public String getTestReport() throws HPSFException, SQLException {
		Connection conn = null;
		Statement stmt = null;
		ResultSet rs = null;
		conn = DBConnectivity.dataBaseConnect();
		ArrayList data = new ArrayList();		
		stmt = conn.createStatement();
		//String sql = "SELECT user_id, Months,TotalHours, Rate, DecemberCost FROM (SELECT user_id,monthname(date) AS Months,SUM(minutesSpent / 60) AS TotalHours,Cost AS Rate,(Cost * SUM(minutesSpent / 60)) AS DecemberCost FROM hourentries hr JOIN users us ON hr.user_id = us.id GROUP BY user_id , Months) t GROUP BY user_id , Months;";
		String sql="SELECT\n" + 
				"  Resource,Rate,Role,\n" + 
				"  sum(if(month(date) = 1, TotalHours, 0))  AS 'Jan Hours',\n" + 
				"  sum(if(month(date) = 2, TotalHours, 0))  AS 'Feb Hours',\n" + 
				"  sum(if(month(date) = 3, TotalHours, 0))  AS 'Mar Hours',\n" + 
				"  sum(if(month(date) = 4, TotalHours, 0))  AS 'Apr Hours',\n" + 
				"  sum(if(month(date) = 5, TotalHours, 0))  AS 'May Hours',\n" + 
				"  sum(if(month(date) = 6, TotalHours, 0))  AS 'Jun Hours',\n" + 
				"  sum(if(month(date) = 7, TotalHours, 0))  AS 'Jul Hours',\n" + 
				"  sum(if(month(date) = 8, TotalHours, 0))  AS 'Aug Hours',\n" + 
				"  sum(if(month(date) = 9, TotalHours, 0))  AS 'Sep Hours',\n" + 
				"  sum(if(month(date) = 10, TotalHours, 0)) AS 'Oct Hours',\n" + 
				"  sum(if(month(date) = 11, TotalHours, 0)) AS 'Nov Hours',\n" + 
				"  sum(if(month(date) = 12, TotalHours, 0)) AS 'Dec Hours',\n" + 
				"  sum(if(month(date) = 1, MonthlyCost, 0))  AS 'Jan Cost',\n" + 
				"  sum(if(month(date) = 2, MonthlyCost, 0))  AS 'Feb Cost',\n" + 
				"  sum(if(month(date) = 3, MonthlyCost, 0))  AS 'Mar Cost',\n" + 
				"  sum(if(month(date) = 4, MonthlyCost, 0))  AS 'Apr Cost',\n" + 
				"  sum(if(month(date) = 5, MonthlyCost, 0))  AS 'May Cost',\n" + 
				"  sum(if(month(date) = 6, MonthlyCost, 0))  AS 'Jun Cost',\n" + 
				"  sum(if(month(date) = 7, MonthlyCost, 0))  AS 'Jul Cost',\n" + 
				"  sum(if(month(date) = 8, MonthlyCost, 0))  AS 'Aug Cost',\n" + 
				"  sum(if(month(date) = 9, MonthlyCost, 0))  AS 'Sep Cost',\n" + 
				"  sum(if(month(date) = 10, MonthlyCost, 0)) AS 'Oct Cost',\n" + 
				"  sum(if(month(date) = 11, MonthlyCost, 0)) AS 'Nov Cost',\n" + 
				"  sum(if(month(date) = 12, MonthlyCost, 0)) AS 'Dec Cost'\n" + 
				" from\n" + 
				" (SELECT \n" + 
				"       fullName as Resource,hr.user_id,name as Role,\n" + 
				"       date,\n" + 
				"           sum(minutesSpent / 60) AS TotalHours,\n" + 
				"           Cost as Rate,\n" + 
				"           (Cost * SUM(minutesSpent / 60)) AS MonthlyCost\n" + 
				"   FROM\n" + 
				"       hourentries hr\n" + 
				"   JOIN users us ON hr.user_id = us.id join team_user tu on tu.User_id=us.id join teams t on t.id=tu.Team_id\n" + 
				"   GROUP BY user_id) t\n" + 
				"GROUP BY Resource,Role;";
		
		rs = stmt.executeQuery(sql);
		while(rs.next()){
			ArrayList cells = new ArrayList();
			cells.add(rs.getString(1));
			
			cells.add(rs.getString(2));
			
			cells.add(rs.getString(3));
		
			cells.add(rs.getString(4));
			
			cells.add(rs.getString(5));
			
			cells.add(rs.getString(6));
			cells.add(rs.getString(7));
			cells.add(rs.getString(8));
			cells.add(rs.getString(9));
			cells.add(rs.getString(10));
			cells.add(rs.getString(11));
			cells.add(rs.getString(12));
			cells.add(rs.getString(13));
			cells.add(rs.getString(14));
			cells.add(rs.getString(15));
			cells.add(rs.getString(16));
			cells.add(rs.getString(17));
			cells.add(rs.getString(18));
			cells.add(rs.getString(19));
			cells.add(rs.getString(20));
			cells.add(rs.getString(21));
			cells.add(rs.getString(22));
			cells.add(rs.getString(23));
			cells.add(rs.getString(24));
			cells.add(rs.getString(25));
			cells.add(rs.getString(26));
			cells.add(rs.getString(27));
			
			data.add(cells);
			System.out.println(data.toString());
		}
		ArrayList headers = new ArrayList();
		headers.add("Resource");
		headers.add("Rate");
		headers.add("Role");
		headers.add("Jan Hours");
		headers.add("Feb Hours");
		headers.add("Mar Hours");
		headers.add("Apr Hours");
		headers.add("May Hours");
		headers.add("Jun Hours");
		headers.add("Jul Hours");
		headers.add("Aug Hours");
		headers.add("Sep Hours");
		headers.add("Oct Hours");
		headers.add("Nov Hours");
		headers.add("Dec Hours");
		headers.add("Jan Cost");
		headers.add("Feb Cost");
		headers.add("Mar Cost");
		headers.add("Apr Cost");
		headers.add("May Cost");
		headers.add("Jun Cost");
		headers.add("Jul Cost");
		headers.add("Aug Cost");
		headers.add("Sep Cost");
		headers.add("Oct Cost");
		headers.add("Nov Cost");
		headers.add("Dec Cost");
		exportToExcel("TestReport",headers, data);
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
				hssfCell.setCellValue( (String) cells.next());
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

