package csm.csm;

import java.awt.Color;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ClassScheduleMaker {
	
	public static void main(String[] args) {
		
		ClassScheduleMaker ScheduleMaker = new ClassScheduleMaker();
		
		System.out.println("Please enter the subject list's file name: ");
		Scanner input = new Scanner(System.in);
		String filename = input.nextLine();
		
		ScheduleMaker.getSubjectsFromFile(filename);
		
		System.out.println("Please select a color scheme: ");
		System.out.println("[1] - Black and White");
		System.out.println("[Any key] - Colored");
		System.out.println("Your choice: ");
		
		String colorScheme = input.nextLine();
		ScheduleMaker.setColorScheme(colorScheme);
		
		
		System.out.println("Please enter your name: ");
		String headerName = input.nextLine();
		
		System.out.println("Please enter the output filename (don't include extension): ");
		String outputFileName = input.nextLine();
		
		ScheduleMaker.createExcelFile(headerName, outputFileName, colorScheme);
		
	}
	
	ArrayList<Subject> subjects;
	
	public ClassScheduleMaker() {
		subjects = new ArrayList<Subject>();
	}
	
	/**
	 * gets this schedule's subject list from an existing file
	 * @param filename the subject list's file name
	 */
	public void getSubjectsFromFile(String filename) {
		File file =  new File(filename);
		try {
			BufferedReader br = new BufferedReader(new FileReader(file));
			String line;
			while((line = br.readLine()) != null){
				String[] vars = line.split(",");
				Subject subject = new Subject(vars[0]);
				int day = -1;
				String room = "";
				for(int j = 1; j < vars.length; j++) {
					String[] schedString = vars[j].split("/");
					switch(schedString[0]) {
						case "M": day = 0; break;
						case "T": day = 1; break;
						case "W": day = 2; break;
						case "TH": day = 3; break;
						case "F": day = 4; break;
						case "S": day = 5; break;
						default: day = 6; break;
					}
					
					String[] timeString = schedString[1].split("-");
					int startTimeHour = Integer.parseInt(timeString[0].substring(0,timeString[0].indexOf(':')));
					if(timeString[0].contains("PM") && startTimeHour != 12) startTimeHour += 12;
					else if(timeString[0].contains("AM") && startTimeHour == 12) startTimeHour = 0;
					int startTimeMin = Integer.parseInt(timeString[0].substring(timeString[0].indexOf(":") + 1, timeString[0].indexOf(":") + 3));
					
					int endTimeHour = Integer.parseInt(timeString[1].substring(0,timeString[1].indexOf(':')));
					if(timeString[1].contains("PM") && endTimeHour != 12) endTimeHour += 12;
					else if(timeString[1].contains("AM") && endTimeHour == 12) endTimeHour = 0;
					int endTimeMin = Integer.parseInt(timeString[1].substring(timeString[1].indexOf(":") + 1, timeString[1].indexOf(":") + 3));
				
					room = schedString[2];
					
					int[] start = {startTimeHour,startTimeMin};
					int[] end = {endTimeHour,endTimeMin};
					Schedule subjectSchedule = new Schedule(day,start,end,room);
					subject.schedules.add(subjectSchedule);
					
				}
				
				subjects.add(subject);
				
			}
			
			if(this.hasConflictingSubjects()) {
				System.err.println("The file has conflicting subject schedules!");
				System.err.println("Please double check its contents!");
				System.exit(0);
			}
			
			if(this.hasExceedingScheduleTime()) {
				System.err.println("The file has schedules beyond 7:00 AM - 8:30 PM!");
				System.err.println("Please double check its contents!");
				System.exit(0);
			}
			
			if(this.hasExceedingSubjectCount()) {
				System.err.println("The file has too many subjects!");
				System.exit(0);
			}
			
			if(this.hasInvalidScheduleDuration()){
				System.err.println("The file has an invalid schedule duration!");
				System.err.println("Please double check its contents!");
				System.exit(0);
			}
			
			System.out.println("Subjects loaded!");
			
		} 
		catch (ArrayIndexOutOfBoundsException e) {	
			System.err.println("Invalid file format!");
			System.exit(0);
		}catch (FileNotFoundException e) {
			System.err.println("File not found!");
			System.exit(0);
		}catch (IOException e) {
			System.err.println("I/O Error!");
			e.printStackTrace();
			System.exit(0);
		}
	}
	/**
	 * checks if this schedule has conflicting subjects
	 * @return true if the schedule has conflicting subjects, false otherwise
	 */
	public boolean hasConflictingSubjects() {
		ArrayList<Schedule> allSchedules = new ArrayList<Schedule>();
		for(Subject s: subjects) {
			for(Schedule sc: s.schedules)
				allSchedules.add(sc);
		}
		
		for(int i = 0; i < allSchedules.size(); i++) {
			Schedule a = allSchedules.get(i);
			for(int j = i + 1; j < allSchedules.size(); j++) {
				Schedule b = allSchedules.get(j);
				if(a.isConflicting(b))
					return true;
			}
		}
		
		return false;
			
	}
	/**
	 * checks if this Schedule has more than 14 subjects
	 * @return true if this schedule's subject count exceeds 14 subjects, false otherwise
	 */
	public boolean hasExceedingSubjectCount() {
		return subjects.size() > 14;
	}
	
	/**
	 * checks if this Schedule has subjects with time exceeding 7:00 AM - 8:30 PM
	 * @return true if this schedule has exceeding times, false otherwise
	 */
	public boolean hasExceedingScheduleTime() {
		for(Subject s: subjects) {
			for(Schedule sc: s.schedules) {
				if(sc.start[0] < 7 || sc.end[0] < 7)
					return true;
				if(sc.end[0] == 7 && sc.end[1] == 0)
					return true;
				if(sc.start[0] == 20 && sc.start[1] == 30)
					return true;
				if(sc.start[0] > 20 || sc.end[0] > 20)
					return true;
			}
		}
		
		return false;
	}
	/**
	 * checks if the subject durations have end times less than or equal to their corresponding start times
	 * @return true if a subject has an invalid schedule duration, false otherwise
	 */
	public boolean hasInvalidScheduleDuration(){
		for(Subject s: subjects){
			for(Schedule sc: s.schedules){
				if(sc.getEndY() - sc.getStartY() <= 0)
					return true;
			}
		}
		return false;
	}
	/**
	 * sets the Formatted Excel File's Color Scheme
	 * @param choice - 1 if black and white, any otherwise if colored
	 */
	public void setColorScheme(String choice) {
		if(choice.equals("1")) {
			for(Subject subject: subjects)
				subject.color = new Color(255,255,255);
		}
		else {
			for(int i = 0; i < subjects.size(); i++) {
				Subject subject = subjects.get(i);
				subject.color = Subject.colorPallete[i];
			}
		}
		System.out.println("Color scheme set!");
	}
	
	/**
	 * creates this schedule into a formatted excel file
	 * @param headerName - The Schedule's Header Name
	 * @param outputFileName - The output file name
	 * @param colorScheme - The selected color scheme (1 if black and white, any otherwise if colored)
	 */
	public void createExcelFile(String headerName, String outputFileName, String colorScheme) {
		
		// Steps in Creating the Excel File
	    // 1. Create the .xlsx file and create a single sheet
	    // 2. Create 30 rows (2 header rows, 1 day row, and 27 thirty minute block rows)
	    // 		2a. Create each cell and apply default styling to each:
	    //			- Calibri 8 Font
	    //			- 11.86 points column width
	    //			- Dotted Border
	    //		2b. Set row height to 12.75 points
	    // 3. Create The Schedule Header
	    //		- Merge the entire top 2 rows of the excel file
	    //		- Set the title according to the user's input
	    //		- Apply Impact font
	    //		- Apply the selected color scheme
	    // 4. Create Day Headers
	    //		- MON TUES WED THURS FRI SAT SUN
	    // 5. Create Time Headers
	    //		- Hour count starts at 7
	    //		- Every even (0-based) index has a :00 - :30 minute schedule
		//		- Every odd (0-based) index has a :30 - :00 minute schedule
	    //		- The first 9 blocks are AM schedules, the rest are PM
	    //		- Reset Hour Count to 1 after reaching the 12th hour
	    // 6. Fill Out Subject Blocks
	    //		- Apply custom fill color to each affected subject cell
		//		- If the subject is only 30 minutes:
		//			- Place the subject name on the designated cell
		//		- Else:
		//			- Place the subject name on the top cell
		//			- Place the room name on the cell immediately below
		//			- Remove excess dotted borders for cells in between the subject's schedule range
	    // 7. Finalize Border Styles
	    //		- Apply a thin border around each Day Header
	    //		- Apply a thick black border around the table's edges
	    // 8. Write the WorkBook Object into the File
	
		
		// Step 1 - Create the .xlsx file and create a single sheet
		XSSFWorkbook wb = new XSSFWorkbook();  
	        try  (OutputStream os = new FileOutputStream(outputFileName + ".xlsx")) {  
	        	Sheet sheet = wb.createSheet("Schedule");
	        	sheet.getPrintSetup().setLandscape(true);
	        	
	        	// Step 2 - Create 30 rows
	        	
	            Font font = wb.createFont();  // Calibri 8 Font
	            font.setFontHeightInPoints((short)8);  
	            font.setFontName("Calibri");  
	        	
	        	for(int i = 0; i < 30; i++) {
	        		Row row = sheet.createRow(i);
	        		for(int j = 0; j < 8; j++) {
	        			
	        			// Step 2a
	        			Cell cell = row.createCell(j);
	        			
	        			XSSFCellStyle style = wb.createCellStyle();
	                    style.setFont(font);
	                    style.setAlignment(HorizontalAlignment.CENTER);
	                    style.setBorderBottom(BorderStyle.DOTTED);
	                    style.setBorderLeft(BorderStyle.DOTTED);
	                    style.setBorderRight(BorderStyle.DOTTED);
	                    style.setBorderTop(BorderStyle.DOTTED);
	                    cell.setCellStyle(style);
	                    
	                    
	        			sheet.setColumnWidth(j, 3225); // 11.86 points
	        		}
	        		// Step 2b
	        		row.setHeight((short) 260); // 12.75 points
	        	}
	        	
	        	// Step 3 - Create The Schedule Header
	        	sheet.addMergedRegion(new CellRangeAddress(0,1,0,7));
	        	Row HRow = sheet.getRow(0);
	        	Cell HCell = HRow.getCell(0);
	        	XSSFCellStyle style = wb.createCellStyle();
	        	HCell.setCellValue(headerName+"'s Schedule");
	        	
	        	Font font2 = wb.createFont();  
	            font2.setFontHeightInPoints((short)12);  
	            font2.setFontName("Impact");
	            style.setFont(font2);
	            style.setAlignment(HorizontalAlignment.CENTER);
	            
	            if(!colorScheme.equals("1")) {
	            	style.setFillForegroundColor(new XSSFColor(new java.awt.Color(214, 227, 188)));
	            	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	            }
	            HCell.setCellStyle(style);
	            
	            // Step 4 - Create Day Headers
	            String[] days = {"MON","TUES","WED","THURS","FRI","SAT","SUN"};
	            int DXOffset = 1, DYOffset = 2;
	            Row DRow = sheet.getRow(DYOffset);
	            for(int i = 0; i < 7; i++) {
	            	Cell DCell = DRow.getCell(DXOffset + i);
	            	DCell.setCellValue(days[i]);
	            }
	            
	            // Step 5 - Create Time Headers
	            int hr = 7;
	            
	            int TYOffset = 3;
	        	for(int i = 0; i < 27; i++) {
	        		Row TRow = sheet.getRow(TYOffset + i);
	        		Cell TCell = TRow.getCell(0);
	        		String time = "";
	        		String mdm = i < 9 ? " AM" : " PM";
	        		if(i % 2 == 0) {
	        			time = hr + ":00" + " - " + hr + ":30";
	        		}
	        		else {
	        			if(hr == 12) {
	        				time = hr +":30 - " + "1:00";
	        				hr = 1;
	        			}				
	        
	        			else time = hr + ":30 - " + ++hr + ":00";
	        		}
	        		time += mdm;
	        		TCell.setCellValue(time);
	        	}
	        	
	        	// Step 6 - Fill Out Subject Blocks
	        	for(Subject subject: subjects) {
		    		for(Schedule sched: subject.schedules) {
		    			for(int i = sched.getStartY(); i <= sched.getEndY(); i++) {
		        			Row SRow = sheet.getRow(i);
		        			Cell SCell = SRow.getCell(sched.getX());
		        			XSSFCellStyle Sstyle = (XSSFCellStyle) SCell.getCellStyle();
		    				Sstyle.setFillForegroundColor(new XSSFColor(subject.color));
		    	            Sstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		    	            
		    	            
		    	            // Conditional Schedule Styling
		    	            
		    	            if(sched.isSingleBlock()) {
		    	            	Font subjectFont = wb.createFont();  
		    	                subjectFont.setFontHeightInPoints((short)8);  
		    	                subjectFont.setFontName("Calibri");
		    	                subjectFont.setBold(true);
		    	                
		    	            	SCell.setCellValue(subject.getFormattedName());
		    	            	Sstyle.setFont(subjectFont);
		    	            	Sstyle.setBorderBottom(BorderStyle.NONE);
		    	            }
		    	            else {
			    	            if(i == sched.getStartY()) {
			    	            	Font subjectFont = wb.createFont();  
			    	                subjectFont.setFontHeightInPoints((short)8);  
			    	                subjectFont.setFontName("Calibri");
			    	                subjectFont.setBold(true);
			    	                
			    	            	SCell.setCellValue(subject.getFormattedName());
			    	            	Sstyle.setFont(subjectFont);
			    	            	Sstyle.setBorderBottom(BorderStyle.NONE);
			    	            }
			    	            if(i == sched.getStartY() + 1) {
			    	            	Font subjectFont = wb.createFont();  
			    	                subjectFont.setFontHeightInPoints((short)8);  
			    	                subjectFont.setFontName("Calibri");
			    	                subjectFont.setItalic(true);
			    	                
			    	                SCell.setCellValue(sched.room);
			    	            	Sstyle.setFont(subjectFont);
			    	                
			    	            }
			    	            if(i > sched.getStartY() && i < sched.getEndY()) {
			    	            	Sstyle.setBorderTop(BorderStyle.NONE);
			    	            	Sstyle.setBorderBottom(BorderStyle.NONE);
			    	            }
			    	            if(i == sched.getEndY())
			    	            	Sstyle.setBorderTop(BorderStyle.NONE);
		    	            }
		    	            
		    	            SCell.setCellStyle(Sstyle);
		        		}
		    		}
	        	}
	    		
	    		// Step 7 - Finalize Border Styles
	    		
	    		// a. - Thin Day Borders
	    		for(int i = 0; i < 8; i++) {
	    			Row ARow = sheet.getRow(2);
	    			Row BRow = sheet.getRow(3);
	    			
	    			Cell ACell = ARow.getCell(i);
	    			XSSFCellStyle AStyle = (XSSFCellStyle) ACell.getCellStyle();
	    			
	    			Cell BCell = BRow.getCell(i);
	    			XSSFCellStyle BStyle = (XSSFCellStyle) BCell.getCellStyle();
	    			
	    			AStyle.setBorderLeft(BorderStyle.THIN);
	    			AStyle.setBorderRight(BorderStyle.THIN);
	    			AStyle.setBorderBottom(BorderStyle.THIN);
	    			
	    			BStyle.setBorderTop(BorderStyle.THIN);
	    			
	    		}
	    		
	    		// b. Thick Left and Right Borders 
	    		for(int i = 0; i < 30; i++) {
	    			Row SRow = sheet.getRow(i);
	    			Cell leftCell = SRow.getCell(0);
	    			Cell rightCell = SRow.getCell(7);
	    			
	    			XSSFCellStyle leftstyle = (XSSFCellStyle) leftCell.getCellStyle();
	    			XSSFCellStyle rightstyle = (XSSFCellStyle) rightCell.getCellStyle();
	    			
	    			leftstyle.setBorderLeft(BorderStyle.THICK);
	    			rightstyle.setBorderRight(BorderStyle.THICK);
	    		}
	    		
	    		// c. Thick Top and Bottom Borders 
	    		Row firstRow = sheet.getRow(0);
	    		Cell firstCell = firstRow.getCell(0);
	    		
	    		XSSFCellStyle firststyle = (XSSFCellStyle) firstCell.getCellStyle();
	    		firststyle.setBorderTop(BorderStyle.THICK);
	    		
	    		Row secondRow = sheet.getRow(1);
	    		Cell secondCell = secondRow.getCell(0);
	    		
	    		XSSFCellStyle secondstyle = (XSSFCellStyle) secondCell.getCellStyle();
	    		secondstyle.setBorderBottom(BorderStyle.THICK);
	    		
	    		for(int i = 0; i < 8; i++) {
	    			Row topRow = sheet.getRow(0);
	    			Row topRow2 = sheet.getRow(1);
	    			Row botRow = sheet.getRow(29);
	    			
	    			Cell topCell = topRow.getCell(i);
	    			Cell topCell2 = topRow2.getCell(i);
	    			Cell botCell = botRow.getCell(i);
	    			
	    			XSSFCellStyle topStyle = (XSSFCellStyle) topCell.getCellStyle();
	    			XSSFCellStyle topStyle2 = (XSSFCellStyle) topCell2.getCellStyle();
	    			XSSFCellStyle botStyle = (XSSFCellStyle) botCell.getCellStyle();
	    			
	    			topStyle.setBorderTop(BorderStyle.THICK);
	    			topStyle2.setBorderBottom(BorderStyle.THICK);
	    			botStyle.setBorderBottom(BorderStyle.THICK);
	    			
	    		}
	   
	    		// Step 8 - Write the WorkBook Object into the File
	            wb.write(os);
	            wb.close();
	            System.out.println("Excel file created!");
	        }catch(Exception e) {  
	            System.out.println(e.getMessage());  
	        }
	}

}