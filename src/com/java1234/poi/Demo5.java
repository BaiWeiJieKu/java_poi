package com.java1234.poi;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo5 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); //定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个sheet页");  //创建第一个sheet页
		Row row=sheet.createRow(0); //创建一行
		Cell cell=row.createCell(0); //创建一个单元格 第一列
		cell.setCellValue(new Date());  // 给单元格设置一个时间值
		
		row.createCell(1).setCellValue(1);
		row.createCell(2).setCellValue("一个字符串");
		row.createCell(3).setCellValue(true);
		row.createCell(4).setCellValue(HSSFCell.CELL_TYPE_NUMERIC);
		row.createCell(5).setCellValue(false);
		
		FileOutputStream fileOut=new FileOutputStream("c:\\新工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
