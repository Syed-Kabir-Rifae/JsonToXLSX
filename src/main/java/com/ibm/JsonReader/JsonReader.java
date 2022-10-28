package com.ibm.JsonReader;
import java.io.*;
import net.minidev.json.JSONArray;
import net.minidev.json.JSONObject;
import net.minidev.json.parser.JSONParser;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JsonReader{

	public static void main(String[] args) {
		try {
			
			//Taking Input JSON File
			JSONParser parser = new JSONParser();
			JSONArray a = (JSONArray)parser.parse(new FileReader(("C:\\Users\\003C6G744\\Desktop\\Json\\Sample.json")));
			
			//Create workbook in .xlsx format
			Workbook workbook = new XSSFWorkbook();
		
			//Create Sheet
			Sheet sh = workbook.createSheet("Employee Data");
			
			//Create top row with column headings
			String[] colHeadings = {"Employee Name","Employee ID","Location"," Role"," Rating","Date of Joining","    Salary"};
			//We want to make it bold with a foreground color.
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short)12);
			headerFont.setColor(IndexedColors.BLACK.index);
			
			
			//Create a CellStyle with the font
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerStyle.setFillForegroundColor(IndexedColors.YELLOW1.index);
			
			
			//Create the header row
			Row headerRow = sh.createRow(0);
			
			
			//Iterate over the column headings to create columns
			for(int i=0;i<colHeadings.length;i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(colHeadings[i]);
				cell.setCellStyle(headerStyle);
			}
			//Freeze Header Row
			sh.createFreezePane(0, 1);

	
			CreationHelper creationHelper= workbook.getCreationHelper();
			CellStyle dateStyle = workbook.createCellStyle();
			dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("₹ #,##0.00"));
		    String arr[] ={"name","id","location","role","rating","date"};
			
			int rownum =1;
			 for (Object i : a)
			  {
				 int k=0;
			  
				JSONObject p = (JSONObject) i;
			    Row row = sh.createRow(rownum++);
			    
			    while(k<6) {
				 Cell cell=row.createCell(k);
				CellStyle cellSt = workbook.createCellStyle();
				cellSt.setWrapText(true);
				cellSt.setAlignment(HorizontalAlignment.CENTER);
				cell.setCellValue((String) p.get(arr[k]));
				cell.setCellStyle(cellSt);
				k++;
				 }
			    
			    
				Cell dateCell = row.createCell(6);
				dateCell.setCellValue((int) p.get("salary"));
				dateCell.setCellStyle(dateStyle);
				
			  }	

			//Autosize columns
			for(int i=0;i<colHeadings.length;i++) {
				sh.autoSizeColumn(i);
			}
			//Write the output to file
			FileOutputStream fileOut = new FileOutputStream("C:\\Users\\003C6G744\\Desktop\\Json\\employee.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
			System.out.println("Completed");
		}
		catch(Exception e) {
			e.printStackTrace();
		}}}
