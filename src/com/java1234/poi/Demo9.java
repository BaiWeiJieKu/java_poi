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
边框样式
*/
public class Demo9 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook(); //定义一个新的工作簿
		Sheet sheet=wb.createSheet("第一个sheet页");  //创建第一个sheet页
		Row row=sheet.createRow(1); //创建一个行
		
		Cell cell=row.createCell(1); // 创建一个单元格
		cell.setCellValue(4);
		
		CellStyle cellStyle=wb.createCellStyle(); 
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN); //底部边框
		cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 底部颜色
		
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);  // 左边边框
		cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex()); // 左边框颜色
		
		cellStyle.setBorderRight(CellStyle.BORDER_THIN); // 右边框
		cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());  // 右边框颜色
		
		cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); // 上边框
		cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());  // 上边框颜色
		
		cell.setCellStyle(cellStyle);
		FileOutputStream fileOut=new FileOutputStream("c:\\工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
