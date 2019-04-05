package com.pras.test;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Row;

import com.google.gson.Gson;

public class JsonParser {

	public static void main(String[] args) throws IOException, org.json.simple.parser.ParseException, IllegalAccessException, InstantiationException {
		  // 1.read the json file
        //JSONObject jsonObject = readJson();
		
		VolumeContainer vc = readJson();
		HashMap<String, String> hm = new HashMap<String,String>();
		for(Volume v: vc.volumes){
		  hm.put(v.getId(), v.getName());  
		}
		
		createExcelFile(vc);
        System.out.println(hm);

	}

	public static VolumeContainer readJson() throws IOException, org.json.simple.parser.ParseException {
	    String filePath = "src/main/resources/sample.json";
	    //FileInputStream inputStream = new FileInputStream("src/com/product/resource/config.properties");
	    FileReader reader = new FileReader(filePath);

	    /*    JSONParser jsonParser = new JSONParser();
	    return (JSONObject) jsonParser.parse(reader);*/
	    
	    Gson g = new Gson();
	    VolumeContainer vc = g.fromJson(reader, VolumeContainer.class);
	    return vc;
	}
	
	public static void createExcelFile(VolumeContainer vc) throws IOException, IllegalAccessException, InstantiationException {
	    FileOutputStream fileOut = new FileOutputStream("src/main/resources/Response.xls");
	    HSSFWorkbook workbook = new HSSFWorkbook();
	    HSSFSheet worksheet = workbook.createSheet("product details");
	    HSSFRow row1 = worksheet.createRow((short) 0);
	    short index = 0;

	    //create header
/*	    for (String header : getHeader()) {
	        HSSFCell cellA1 = row1.createCell(index);
	        cellA1.setCellValue(header);
	        HSSFCellStyle cellStyle = workbook.createCellStyle();
	        cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
	        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	        cellA1.setCellStyle(cellStyle);
	        index++;
	    }*/
	    
	    for (int i = 0 ; i < 2 ; i++) {
	        HSSFCell cellA1 = row1.createCell(index);
	        cellA1.setCellValue("Column "+ (i+1));
	        HSSFCellStyle cellStyle = workbook.createCellStyle();
	        cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
	        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	        cellA1.setCellStyle(cellStyle);
	        index++;
	    }

	    //create rows
	    index = 1;
	    for(Volume v: vc.volumes){
	        HSSFRow excelRow = worksheet.createRow(index);
/*	        short flag = 0;
	        for (String field : row.getField()) {
	            HSSFCell cellA1 = excelRow.createCell(flag);
	            cellA1.setCellValue(field);
	            flag++;
	        }*/
	        
            HSSFCell cellA1 = excelRow.createCell(0);
            cellA1.setCellValue(v.getId());
            
            HSSFCell cellA2 = excelRow.createCell(1);
            cellA2.setCellValue(v.getName());
            
	        index++;
	    }

	    workbook.write(fileOut);
	    fileOut.flush();
	    fileOut.close();
	}
}

