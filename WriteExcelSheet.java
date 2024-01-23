package org.wrokInExcelsheet;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelSheet {
	public static void main(String[] args) throws IOException  {
		
	XSSFWorkbook book = new XSSFWorkbook();
	XSSFSheet sheet = book.createSheet("Sheet2");
	Object [][] datas= {
						{"name",    "age", "city"},
						{"abdul",   "28",  "trichy"},
						{"shankar", "29",  "karur"},
						{"kavin",   "28",  "erode"},
						{"jacob",   "29",  "tanjore"},
						};
	int rowcount = 0;
	for (Object[] row1 : datas) {
			XSSFRow row = sheet.createRow(rowcount++);
			int colomncount = 0;
	for (Object col : row1) {
		XSSFCell cell = row.createCell(colomncount++);
		
		if(col instanceof String) {
			cell.setCellValue((String) col);
		}
		else if(col instanceof Integer){
			cell.setCellValue((Integer) col);
		}
	}				
	}
	FileOutputStream stream = new FileOutputStream("C:\\Users\\HP\\Eclipse workspace new\\ExcerciseOpet\\workExcel\\workbook.xlsx");
	book.write(stream);
	
	System.out.println("Updated successfully");
	
		
	}
	}


