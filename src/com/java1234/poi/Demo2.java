package com.java1234.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class Demo2 {

	public static void main(String[] args) throws Exception {
		
		Workbook wb=new HSSFWorkbook(); // 定义一个新的工作簿
		wb.createSheet("第一个sheet页");  // 创建一个sheet页
		wb.createSheet("第二个sheet页");  // 创建第二个sheet页
		FileOutputStream fileOut=new FileOutputStream("c:\\多个sheet页的工作簿.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
