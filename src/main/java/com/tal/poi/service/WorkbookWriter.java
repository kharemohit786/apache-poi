package com.tal.poi.service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class WorkbookWriter {

	
	static String fileLoc = "src/test/resources/reportFile.xls";

	public static void WriteExcel() {
		HSSFWorkbook workbook = new HSSFWorkbook();
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] { "ID", "NAME", "LASTNAME" });
		data.put("2", new Object[] { 1, "Amit", "Shukla" });
		data.put("3", new Object[] { 2, "Lokesh", "Gupta" });
		data.put("4", new Object[] { 3, "John", "Adwards" });
		data.put("5", new Object[] { 4, "Brian", "Schultz" });
		
		createSheet(workbook, "sheet1", data);
		createSheet(workbook, "sheet2", data);

	}

	public static void createSheet(HSSFWorkbook workbook, String sheetName, Map<String, Object[]> data) {

		HSSFSheet sheet = workbook.createSheet(sheetName);

		updateCells(sheet, data);

		try {
			FileOutputStream out = new FileOutputStream(new File(fileLoc));
			workbook.write(out);
			workbook.close();
			out.close();
			System.out.println("report.xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void updateCells(HSSFSheet sheet, Map<String, Object[]> data) {
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}
	}
}
