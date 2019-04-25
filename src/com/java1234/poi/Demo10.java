package com.java1234.poi;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/*
填充色
*/
public class Demo10 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); //定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个sheet页");  //创建第一个sheet页
		Row row=sheet.createRow(1); //创建一个行
		
		Cell cell=row.createCell(1);//创建一列
		cell.setCellValue("XX");
		CellStyle cellStyle=wb.createCellStyle();//创建一个列样式
		cellStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex()); // 背景色
		cellStyle.setFillPattern(CellStyle.BIG_SPOTS);  
		cell.setCellStyle(cellStyle);
		
		
		Cell cell2=row.createCell(2);//创建第二列
		cell2.setCellValue("YYY");
		CellStyle cellStyle2=wb.createCellStyle();
		cellStyle2.setFillForegroundColor(IndexedColors.RED.getIndex()); // 前景色
		cellStyle2.setFillPattern(CellStyle.SOLID_FOREGROUND);  
		cell2.setCellStyle(cellStyle2);
		
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
