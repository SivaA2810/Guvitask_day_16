package org.wrokInExcelsheet;

import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readexcel_anotherMethod {
public static void main(String[] args) throws IOException {
	
	XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\HP\\Eclipse workspace new\\ExcerciseOpet\\workExcel\\workbook.xlsx");
	XSSFSheet sheet = book.getSheet("Sheet2");
	int rowcount = sheet.getLastRowNum(); 
	
	XSSFRow allrow = sheet.getRow(0);
	int colomncount = allrow.getLastCellNum();
	
	String[][] data = new String[rowcount][colomncount];	
	for (int i = 1; i <=rowcount; i++) {
	XSSFRow row = sheet.getRow(i);
	System.out.println();
	for (int j = 0; j <colomncount ; j++) {
		XSSFCell cell = row.getCell(j);
		data [i-1][j]= cell.getStringCellValue();
		
		System.out.print("\t"+cell.getStringCellValue()+"\t");
	}
}
	System.out.println("datas fetched Successfully");
}

}
