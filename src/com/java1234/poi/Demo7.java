package com.java1234.poi;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class Demo7 {

	public static void main(String[] args) throws Exception{
		InputStream is=new FileInputStream("c:\\二货名单.xls");
		POIFSFileSystem fs=new POIFSFileSystem(is);
		HSSFWorkbook wb=new HSSFWorkbook(fs);
		//抽取所有数据
		ExcelExtractor excelExtractor=new ExcelExtractor(wb);
		excelExtractor.setIncludeSheetNames(false);// 不需要sheet页的名字
		System.out.println(excelExtractor.getText());
	}
	

}
