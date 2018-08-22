package com.cong.app;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map.Entry;

import org.junit.Test;

import com.cong.excel.ReadExcel;
import com.cong.excel.WriteExcel;

/**
 * 启动类
 * 
 * @author cong
 */
public class Client {

	/**
	 * 将json文件写入excel中(不分版本)
	 * @throws IOException 
	 */
	@Test
	public void writeExcelForXSSF() throws IOException {
		WriteExcel.writeXSSF("data.json");
	}
	
	/**
	 * 从excel中读取数据(不分版本)
	 * @throws IOException 
	 */
	@Test
	public void readExcelForXSSF() throws IOException {
		LinkedHashMap<String ,String> content = ReadExcel.readXSSF("hello.xlsx");
		for(Entry<String,String> item : content.entrySet()) {
			System.out.printf("%s:%s\t", item.getKey(),item.getValue());
		}
		long i = 1000000000;
		System.out.println(String.format("%,d", i));
	}
	
	
}
