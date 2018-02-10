package com.sample.test.exceltest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class PrepareConfig {
	private Properties prop = new Properties();

	public Properties getProp() {
		return prop;
	}
	public void setProp(Properties prop) {
		this.prop = prop;
	}
	PrepareConfig(){
		InputStream input = null;

		try {

			input = this.getClass().getClassLoader().getResourceAsStream("config.properties");
			prop.load(input);
		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			if (input != null) {
				try {
					input.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	public static void main(String[] args) throws IOException {
		new PrepareConfig().readData();
	}


	public void readData(){
		FileInputStream excelFile = null;
		XSSFWorkbook workbook = null;
		FileOutputStream outputStream = null;
		XSSFWorkbook writeWorkbook = null;
		Map<String, String> sidEmpId = new HashMap<String, String>();
		try {
			Sheet writeSheet;
				int rowCount = 0;
				excelFile = new FileInputStream(new File("I:\\BL_15398_Gbl_Client_Access_Annexure - July.xlsx"));
				workbook = new XSSFWorkbook(excelFile);
				for(short s=1; s<12; s++){
					Sheet datatypeSheet = workbook.getSheetAt(s);
					System.out.println(datatypeSheet.getSheetName());
					for (int rn=datatypeSheet.getFirstRowNum(); rn<=datatypeSheet.getLastRowNum(); rn++) {
						Row row = datatypeSheet.getRow(rn);
						String key = "";
						String value = "";
						if (row == null) {
							// There is no data in this row, handle as needed
						} else {

							// Row "rn" has data
							for (int cn=0; cn<row.getLastCellNum(); cn++) {
								Cell cell = row.getCell(cn);

								if (cell == null) {
									// This cell is empty/blank/un-used, handle as needed
								} else if(cell.getAddress().toString().startsWith("G") ||
										cell.getAddress().toString().startsWith("F")){
									if (cell.getCellTypeEnum() == CellType.STRING) {
										String cellValue = (String)cell.getStringCellValue();

										if(cell.getAddress().toString().startsWith("G")){
											key = cellValue;
											sidEmpId.put(key, value);
										}

									}
									else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
										if(cell.getAddress().toString().startsWith("F")){
											value = Double.toString(cell.getNumericCellValue());	
										}
									}
								}
							}
						}
					}
					System.out.println(sidEmpId);
				}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		finally{
			if(excelFile != null){
				try {
					workbook.close();
					excelFile.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		
		for(String key : sidEmpId.keySet()){
			System.out.println(key+"="+sidEmpId.get(key));
		}
	}

}
