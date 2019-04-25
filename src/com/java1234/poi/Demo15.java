package com.java1234.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/*
数据格式
*/
public class Demo15 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); //定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个sheet页");  //创建第一个sheet页
		CellStyle style;
		//创建数据格式
		DataFormat format=wb.createDataFormat();
		Row row;
		Cell cell;
		short rowNum=0;
		short colNum=0;
		//创建一行
		row=sheet.createRow(rowNum++);
		cell=row.createCell(colNum);
		cell.setCellValue(111111.25);
		
		style=wb.createCellStyle();
		style.setDataFormat(format.getFormat("0.0")); // 设置数据格式
		cell.setCellStyle(style);
		
		row=sheet.createRow(rowNum++);
		cell=row.createCell(colNum);
		cell.setCellValue(1111111.25);
		style=wb.createCellStyle();
		style.setDataFormat(format.getFormat("#,##0.000"));
		cell.setCellStyle(style);
		
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
