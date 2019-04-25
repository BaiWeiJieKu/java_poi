package com.java1234.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/*
字体处理
*/
public class Demo12 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); //定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个sheet页");  //创建第一个sheet页
		Row row=sheet.createRow(1); //创建一个行
		
		// 创建一个字体处理类
		Font font=wb.createFont();
		font.setFontHeightInPoints((short)24);//高度
		font.setFontName("Courier New");
		font.setItalic(true);//斜体
		font.setStrikeout(true);
		
		CellStyle style=wb.createCellStyle();
		style.setFont(font);
		
		Cell cell=row.createCell((short)1);
		cell.setCellValue("This is test of fonts");
		cell.setCellStyle(style);
		
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
