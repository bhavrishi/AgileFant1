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

@Component("resourceCapReportAction")
@Scope("prototype")
public class ResourceCapReportAction {
	private InputStream inputStream;

	public String getResourceCapReport() throws HPSFException, SQLException {
		Connection conn = null;
		Statement stmt = null;
		ResultSet rs = null;
		conn = DBConnectivity.dataBaseConnect();
		ArrayList data = new ArrayList();		
		stmt = conn.createStatement();
		//String sql = "SELECT user_id, Months,TotalHours, Rate, DecemberCost FROM (SELECT user_id,monthname(date) AS Months,SUM(minutesSpent / 60) AS TotalHours,Cost AS Rate,(Cost * SUM(minutesSpent / 60)) AS DecemberCost FROM hourentries hr JOIN users us ON hr.user_id = us.id GROUP BY user_id , Months) t GROUP BY user_id , Months;";
		String sql="select Resource,Role,MonthCapacity,HoursEstimated,100*(t3.MonthCapacity-t3.HoursEstimated)/MonthCapacity as '%Free'\n" + 
				"from(\n" + 
				"SELECT \n" + 
				"    fullName as Resource,name as Role,(TIMESTAMPDIFF(DAY, CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), DATE_ADD(CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), INTERVAL 1 MONTH)) - CASE\n" + 
				"                WHEN\n" + 
				"                    WEEKDAY(CONCAT(MONTH(NOW()), '-01', '-', YEAR(NOW()))) = 1\n" + 
				"                        AND WEEKDAY(CONCAT(MONTH(NOW()), '-02', '-', YEAR(NOW()))) = 2\n" + 
				"                THEN\n" + 
				"                    9\n" + 
				"                WHEN\n" + 
				"                    WEEKDAY(CONCAT(MONTH(NOW()), '-01', '-', YEAR(NOW()))) = 7\n" + 
				"                        AND WEEKDAY(CONCAT(MONTH(NOW()), '-02', '-', YEAR(NOW()))) = 1\n" + 
				"                THEN\n" + 
				"                    10\n" + 
				"                WHEN\n" + 
				"                    WEEKDAY(CONCAT(MONTH(NOW()), '-01', '-', YEAR(NOW()))) = 6\n" + 
				"                        AND TIMESTAMPDIFF(DAY, CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), DATE_ADD(CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), INTERVAL 1 MONTH)) = 30\n" + 
				"                THEN\n" + 
				"                    9\n" + 
				"                WHEN\n" + 
				"                    WEEKDAY(CONCAT(MONTH(NOW()), '-01', '-', YEAR(NOW()))) = 6\n" + 
				"                        AND TIMESTAMPDIFF(DAY, CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), DATE_ADD(CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), INTERVAL 1 MONTH)) = 31\n" + 
				"                THEN\n" + 
				"                    10\n" + 
				"				when\n" + 
				"                    TIMESTAMPDIFF(DAY, CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), DATE_ADD(CONCAT(YEAR(NOW()), '-', MONTH(NOW()), '-01'), INTERVAL 1 MONTH)) = 29\n" + 
				"                    AND \n" + 
				"                   WEEKDAY(CONCAT(MONTH(NOW()), '-01', '-', YEAR(NOW()))) in (1,7)   then 9              \n" + 
				"                   ELSE 8\n" + 
				"            END - DATEDIFF(t2.endDate, t2.startDate) )* 8 AS MonthCapacity,SUM(effortleft)/60 as 'HoursEstimated'\n" + 
				"FROM\n" + 
				"    (SELECT \n" + 
				"        h.user_id,hr.story_id,h.endDate,h.startDate,t.effortleft\n" + 
				"    FROM\n" + 
				"        holiday h join hourentries hr\n" + 
				"        on h.user_id=hr.user_id\n" + 
				"     join tasks t\n" + 
				"    on t.story_id = hr.story_id\n" + 
				"        JOIN\n" + 
				"    stories s ON t.story_id = s.id\n" + 
				"        JOIN\n" + 
				"    backlogs b ON b.id = s.backlog_id) t2\n" + 
				"join users u on u.id=t2.user_id join team_user tu on t2.user_id=tu.User_id join teams te on tu.Team_id=te.id group by Resource) t3;";
		rs = stmt.executeQuery(sql);
		while(rs.next()){
			ArrayList cells = new ArrayList();
			cells.add(rs.getString(1));
			
			cells.add(rs.getString(2));
			
			cells.add(rs.getString(3));
		
			cells.add(rs.getString(4));
			cells.add(rs.getString(5));
			
		
			
			data.add(cells);
			System.out.println(data.toString());
		}
		ArrayList headers = new ArrayList();
		headers.add("Resource");
		headers.add("Role");
		headers.add("Month Capacity");
		headers.add("Hours Estimated");
		headers.add("%Free");
		
		exportToExcel("ResourceCapacityReport",headers, data);
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
