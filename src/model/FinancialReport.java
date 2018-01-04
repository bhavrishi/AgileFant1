// Added by Naren Vaishnavi and Vinay

package fi.hut.soberit.agilefant.web;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.Console;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Array;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hpsf.HPSFException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;
import org.springframework.context.annotation.Scope;
import org.springframework.format.datetime.DateFormatter;
import org.springframework.stereotype.Component;

import com.opensymphony.xwork2.Action;

import fi.hut.soberit.agilefant.config.DBConnectivity;

import fi.hut.soberit.agilefant.model.*;


@Component("financialAction")
@Scope("prototype")
public class FinancialReport {


	private InputStream inputStream;

	public String getReports(){
		try{
			ArrayList finaceList = getFinanceVarianceReport();
			exportToExcel((ArrayList)finaceList.get(0), (ArrayList)finaceList.get(1));
		}catch(Exception e)
		{
			e.printStackTrace();
		}
		return Action.SUCCESS;
	}

	public ArrayList getFinanceVarianceReport() throws HPSFException, SQLException {
		Connection conn = null;
		Statement stmt = null;
		ResultSet rs = null;

		ArrayList<Project> projects = new ArrayList<Project>();
		ArrayList<TaskHourEntry> hourEntries = new ArrayList<TaskHourEntry>();
		ArrayList<Task> tasks = new ArrayList<Task>();
		ArrayList<User> users = new ArrayList<User>();
		ArrayList<Story> stories = new ArrayList<Story>();

		conn = DBConnectivity.dataBaseConnect();
		ArrayList data = new ArrayList();	

		stmt = conn.createStatement();
		String sql = "select id, name from backlogs where backlogtype = 'Project'";
		rs = stmt.executeQuery(sql);
		System.out.println(rs);
		while(rs.next()) {
			Project project = new Project();
			project.setId(Integer.parseInt(rs.getString("id")));
			project.setName(rs.getString("name"));
			projects.add(project);

		}
		rs = null;

		sql = "select date, minutesSpent,user_id, task_id from hourentries where DTYPE='TaskHourEntry'";
		rs = stmt.executeQuery(sql);
		try {
			while(rs.next()) {
				TaskHourEntry the = new TaskHourEntry();
				the.setMinutesSpent(Integer.parseInt(rs.getString("minutesSpent")));
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-mm-dd HH:MM:SS.SSS");
				Date d1 = sdf.parse(rs.getString("date"));
				Date d2 = new Date(d1.getYear(), d1.getMonth(), d1.getDate(), d1.getHours(), d1.getMinutes(), d1.getSeconds());
				DateTime dt = new DateTime(d2.getTime()); 
				the.setDate(dt);
				User user = new User();
				user.setId(Integer.parseInt(rs.getString("user_id")));
				the.setUser(user);
				Task task = new Task();
				task.setId(Integer.parseInt(rs.getString("task_id")));
				the.setTask(task);
				hourEntries.add(the);

			}
		}catch (Exception e) {
			// TODO: handle exception
			System.out.println(e);
		}
		rs = null;

		sql = "select id,originalestimate,effortleft,story_id from tasks";
		System.out.println(sql);
		rs = stmt.executeQuery(sql);
		while(rs.next()) {

			Task t = new Task();
			t.setId(Integer.parseInt(rs.getString("id")));
			ExactEstimate originalEstimate =  new ExactEstimate(Integer.parseInt(rs.getString("originalestimate")));
			t.setOriginalEstimate(originalEstimate);
			ExactEstimate effortLeft =  new ExactEstimate(Integer.parseInt(rs.getString("effortleft")));
			t.setEffortLeft(effortLeft);
			Story story = new Story();
			story.setId(Integer.parseInt(rs.getString("story_id")));
			t.setStory(story);
			tasks.add(t);

		}

		rs = null;
		sql = "select id,cost from users";
		rs = stmt.executeQuery(sql);

		while(rs.next()) {
			User u = new User();
			u.setId(Integer.parseInt(rs.getString("id")));
			u.setCost(Integer.parseInt(rs.getString("cost")));
			users.add(u);

		}

		rs = null;
		sql = "select id, backlog_id from stories";
		rs = stmt.executeQuery(sql);

		while(rs.next()) {
			Story s = new Story();
			s.setId(Integer.parseInt(rs.getString("id")));
			Backlog b = new Backlog() {

				@Override
				public boolean isStandAlone() {
					// TODO Auto-generated method stub
					return false;
				}
			};
			b.setId(Integer.parseInt(rs.getString("backlog_id")));
			s.setBacklog(b);

			stories.add(s);
		}

		for(int i=0;i<projects.size();i++) {
			Project p = projects.get(i);
			float hoursSpent = 0;
			float totalCost = 0;
			ArrayList<Story> storiesUnderProject = new ArrayList<Story>();
			int pid = p.getId();
			for(Story s:stories) {
				Backlog b = s.getBacklog();
				if(b.getId()==p.getId()) {
					storiesUnderProject.add(s);
				}
			}
			for(int j=0;j<storiesUnderProject.size();j++) {
				Story s = storiesUnderProject.get(j);
				ArrayList<Task> tasksUnderStory = new ArrayList<Task>();
				for(Task t:tasks) {
					Story story = t.getStory();
					if(s.getId()==story.getId()) {
						tasksUnderStory.add(t);
					}
				}

				for(Task t:tasksUnderStory) {
					for(TaskHourEntry the: hourEntries) {
						Task task = the.getTask();
						if(task.getId()==t.getId()) {
							User user = the.getUser();
							for(User u:users) {
								if(u.getId()==user.getId()) {
									hoursSpent = hoursSpent + the.getMinutesSpent()/60;
									totalCost = totalCost+u.getCost()*(the.getMinutesSpent()/60);
								}
							}
						}
					}
				}
			}
			ArrayList cells = new ArrayList();
			cells.add(p.getId());
			cells.add(p.getName());
			cells.add(hoursSpent);
			cells.add(totalCost);
			data.add(cells);
		}

		ArrayList headers = new ArrayList();
		headers.add("ProjectID");
		headers.add("Project Name");

		headers.add("Hours Spent");

		headers.add("Cost");
		ArrayList hd = new ArrayList();
		hd.add(headers);
		hd.add(data);
		System.out.println("Executed getFinanceReport...");
		return hd;
	}


	public void exportToExcel(ArrayList headers, ArrayList data) throws HPSFException {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("FinanceReport");
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
				hssfCell.setCellValue(String.valueOf(cells.next()));
			}
		}
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
