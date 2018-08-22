package com.cong.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map.Entry;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;

/**
 * 将数据写入excel
 * @author cong
 *
 */
public class WriteExcel {

	@SuppressWarnings("resource")
	public static void writeXSSF(String fileName) throws IOException {
		//创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //新建工作表
        XSSFSheet sheet = workbook.createSheet("hello");
        
        //得到文件的路径
        String path = WriteExcel.class.getClassLoader().getResource(fileName).getPath();
        if(path == null || "".equals(path)) {
        	return;
        }
        File file = new File(path);
        
        //读取文件内容
        String content = FileUtils.readFileToString(file,"UTF-8");
        //将内容转为LinkedHashMap(为了保证有序)
        LinkedHashMap<String, String> contentMap = JSONObject.parseObject(content, new TypeReference<LinkedHashMap<String, String>>(){});
        //遍历
        int index = 0;//索引
        for(Entry<String,String> item : contentMap.entrySet()) {
        	//创建工作表的行,从0开始
        	XSSFRow row = sheet.createRow(index);
        	//创建行的列,从0开始
        	XSSFCell cell1 = row.createCell(0);
        	XSSFCell cell2 = row.createCell(1);
        	//给单元格赋值
        	cell1.setCellValue(item.getKey());
        	cell2.setCellValue(item.getValue());
        	index ++;
        }
        
        //创建输出流
        FileOutputStream fos = new FileOutputStream(new File("/hello.xlsx"));
        workbook.write(fos);
        //关闭流
        workbook.close();
        fos.close();
	}
}
