package org.thoufeeq;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;


public class Thoufeeq {

	
		
		@Parameters({"inputFile","outputFile"})
		@Test
		
		public void meth1(String TargetSheetPathAndName,String NewSheetPathAndName ) {

		

		
		if (TargetSheetPathAndName != null && !"".equals(TargetSheetPathAndName.trim())) {

		try {

		File targetFile = new File(TargetSheetPathAndName.trim());

		FileInputStream inputStream = new FileInputStream(targetFile);

		XSSFWorkbook inputWorkbook = new XSSFWorkbook(inputStream);

		int targetSheetCount = inputWorkbook.getNumberOfSheets();

		System.out.println("Total no. of sheet(s) in the Target Workbook: " + targetSheetCount);

		File outputFile = new File(NewSheetPathAndName.trim());

		FileOutputStream outputStream = new FileOutputStream(outputFile);

		XSSFWorkbook outputWorkbook = new XSSFWorkbook();

		// Step #2 : Creating sheets with the same name as appearing in target workbook.

		for (int i = 0; i < targetSheetCount; i++) {

		XSSFSheet targetSheet = inputWorkbook.getSheetAt(i);

		String inputSheetName = inputWorkbook.getSheetName(i);

		XSSFSheet outputSheet = outputWorkbook.createSheet(inputSheetName);

		copyExcelWB(targetSheet, outputSheet);

		}

		// Step #4 : Write all the sheets in the new Workbook using FileOutStream Object (Step 3 is mentioned below)

		outputWorkbook.write(outputStream);

		outputStream.close();

		}

		catch (Exception ex) {

		System.out.println("Please check the target sheet given path and name: " + TargetSheetPathAndName);

		System.out.println();

		ex.printStackTrace();

		}

		}

		}

		public static void copyExcelWB(XSSFSheet targetSheet, XSSFSheet outputSheet) {

		int rowCount = targetSheet.getLastRowNum();

		System.out.println("There are " + rowCount + " rows in the Target workbook with sheet name -" + ""

		+ targetSheet.getSheetName() + "");

		int currentRowIndex = 0;

		if (rowCount > 0) {

		Iterator<Row> rowIterator = targetSheet.iterator();

		while (rowIterator.hasNext()) {

		int currentCellIndex = 0;

		Iterator<Cell> cellIterator = ((Row) rowIterator.next()).cellIterator();

		while (cellIterator.hasNext()) {

		// Step #3: Creating new Row, Cell and Input value in the newly created sheet.

		String cellData = cellIterator.next().toString();

		if (currentCellIndex == 0)

		outputSheet.createRow(currentRowIndex).createCell(currentCellIndex).setCellValue(cellData);

		else

		outputSheet.getRow(currentRowIndex).createCell(currentCellIndex).setCellValue(cellData);
		
		currentCellIndex++;

		}

		currentRowIndex++;

		}

		System.out.println("Total " + (currentRowIndex - 1) + " rows are Copied in the new Workbook with sheet name- "

		+ "" + outputSheet.getSheetName() + "");

		}
	
		

		}
		
		}
		

		