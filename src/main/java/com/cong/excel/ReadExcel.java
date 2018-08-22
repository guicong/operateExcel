package com.cong.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 从excel中读取数据
 * @author cong
 *
 */
public class ReadExcel {

	
	public static LinkedHashMap<String ,String> readXSSF(String fileName) throws IOException {
		//定义一个LinkedHashMap存放数据
		LinkedHashMap<String ,String> content = new LinkedHashMap<String, String>();
        //得到文件的路径
        String path = ReadExcel.class.getClassLoader().getResource(fileName).getPath();
        if(path == null || "".equals(path)) {
        	return null;
        }
        File file = new File(path);
        
        //创建输入流
        FileInputStream fis = new FileInputStream(file);
    	XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
    	
    	//遍历工作簿中的工作表
    	for(int numSheet = 0;numSheet < xssfWorkbook.getNumberOfSheets();numSheet++) {
    		//得到工作表
    		XSSFSheet sheet = xssfWorkbook.getSheetAt(numSheet);
    		if(sheet == null) {
    			continue;
    		}
    		//遍历工作表中的行
    		for(int numRow = 0;numRow < sheet.getLastRowNum();numRow++) {
    			//得到行
    			XSSFRow row = sheet.getRow(numRow);
    			//得到每列的值并放入LinkedHashMap中
    			content.put(getValue(row.getCell(0)), getValue(row.getCell(1)));
    		}
    	}
    	//关闭流
    	fis.close();
    	xssfWorkbook.close();
    	return content;
	}
	
	/**
	 * 将单元格的数据转为字符串
	 */
	public static String getValue(XSSFCell cell) {
		if(cell == null) {
			return "";
		}
		//都按文本格式读取
		cell.setCellType(CellType.STRING);
		return cell.getStringCellValue();
	}
	
}
