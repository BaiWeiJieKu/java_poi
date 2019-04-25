package com.java1234.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo3 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); // 定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个sheet页");  // 创建第一个sheet页
		Row row=sheet.createRow(0); // 创建一个行
		Cell cell=row.createCell(0); //创建一个单元格，第一列
		cell.setCellValue(1);  // 给单元格设置值为1
		
		row.createCell(1).setCellValue(1.2);   //创建一个单元格 第二列值是1.2
		
		row.createCell(2).setCellValue("字符串"); //创建一个单元格 第三列值是字符串
		
		row.createCell(3).setCellValue(false);  //创建一个单元格 第四列值是布尔类型
		
		FileOutputStream fileOut=new FileOutputStream("c:\\新工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
