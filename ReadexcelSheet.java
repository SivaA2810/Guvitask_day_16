package org.wrokInExcelsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadexcelSheet {
public static void main(String[] args) throws IOException {
	
	XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\HP\\Eclipse workspace new\\ExcerciseOpet\\workExcel\\workbook.xlsx");
	XSSFSheet sheet = book.getSheet("Sheet2");
	XSSFRow row1 = sheet.getRow(0);
	int allRows = sheet.getPhysicalNumberOfRows();
	int allcells = row1.getPhysicalNumberOfCells();
	for (int i = 0; i < allRows; i++) {
		System.out.println();
		XSSFRow row = sheet.getRow(i);
	for (int j=0; j<allcells;j++) {
		XSSFCell cell2 = row.getCell(j);
		System.out.print("\t"+cell2+"\t");
	}
	}
	
}
}
