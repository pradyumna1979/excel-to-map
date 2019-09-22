/**
 * 
 */
package com.blogspot.na5cent.exom.util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * @author pradyumna.k.khadanga
 *
 */
public class ExcelWriter {

	public static void writeToXlsx(List<Map<String,String>> dataMap){

		List<String> headers= dataMap.stream()
				.flatMap(map->map.keySet().stream())
				.distinct()
				.collect(Collectors.toList());

	
		String[] columns = (String[]) headers.toArray(new String[headers.size()]);

		Workbook workbook = new XSSFWorkbook();     // new HSSFWorkbook() for generating `.xls` file

		/* CreationHelper helps us create instances for various things like DataFormat,
	           Hyperlink, RichTextString etc in a format (HSSF, XSSF) independent way */
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Company");

		// Create a Font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Creating cells
		for(int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		CellStyle styleRed = workbook.createCellStyle();
		styleRed.setFillForegroundColor(IndexedColors.GREEN.getIndex());		  
		Font font = workbook.createFont();
		font.setColor(IndexedColors.RED.getIndex());
		styleRed.setFont(font);

		CellStyle styleGreen = workbook.createCellStyle();
		styleGreen.setFillForegroundColor(IndexedColors.GREEN.getIndex());		  
		Font fontGreen = workbook.createFont();
		fontGreen.setColor(IndexedColors.GREEN.getIndex());
		styleGreen.setFont(fontGreen);



		// Create Other rows and cells with employees data
		int rowNum = 1;

		for(Map<String,String> map: dataMap) {

			for(int i=1;i<columns.length;i++) {
				Row row = sheet.createRow(rowNum++);
				row.createCell(i).setCellValue(map.get(columns[i]));

			}
		}
		// Resize all columns to fit the content size
		for(int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream("C:///Users//pradyumna//sonar_monitor.xlsx");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			workbook.write(fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}
