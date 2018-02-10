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


public class CompareExcel {
	Properties prop = new Properties();

	CompareExcel(){
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
		new CompareExcel().readData();;
	}


	public void readData(){
		FileInputStream excelFile = null;
		XSSFWorkbook workbook = null;
		FileOutputStream outputStream = null;
		XSSFWorkbook writeWorkbook = null;
		int count=0;
		Map<String, Double> hoursData = new HashMap<String, Double>();
		try {
			Sheet writeSheet;
			File folder = new File("I:\\August");
			File[] files  = folder.listFiles();
			for(File file : files){
				int rowCount = 0;
				excelFile = new FileInputStream(new File(file.getAbsolutePath()));
				workbook = new XSSFWorkbook(excelFile);
				Sheet datatypeSheet = workbook.getSheetAt(0);
				for (int rn=datatypeSheet.getFirstRowNum(); rn<=datatypeSheet.getLastRowNum(); rn++) {
					Row row = datatypeSheet.getRow(rn);
					String key = "";
					if (row == null) {
						// There is no data in this row, handle as needed
					} else {

						// Row "rn" has data
						for (int cn=0; cn<row.getLastCellNum(); cn++) {
							Cell cell = row.getCell(cn);

							if (cell == null) {
								// This cell is empty/blank/un-used, handle as needed
							} else if(cell.getAddress().toString().startsWith("K") ||
									cell.getAddress().toString().startsWith("O")||
									cell.getAddress().toString().startsWith("B")){
								//String cellStr = fmt.formatCell(cell);
								// Do something with the value
								if (cell.getCellTypeEnum() == CellType.STRING) {
									String cellValue = (String)cell.getStringCellValue();
									if(!"Resource SID (if known)".equals(cellValue) && 
											!"Hours Worked".equals(cellValue) &&
											!"OT Hrs".equals(cellValue)){
										key = cellValue;
										System.out.println("key "+ key);
										key = prop.getProperty(key);
										if(key != null)
											count++;
										else
											System.out.println("NOT "+key);
									}
								} else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
									Double cellValue = (Double) cell.getNumericCellValue();
									//System.out.println(cellValue);
									if(hoursData.get(key) != null){
										cellValue = hoursData.get(key) + cellValue;
									}
									hoursData.put(key, cellValue);
								}
								else if (cell.getCellTypeEnum() == CellType.FORMULA) {
									double cellValue = cell.getNumericCellValue();
									//System.out.println(cellValue);
									if(hoursData.get(key) != null){
										cellValue = hoursData.get(key) + cellValue;
									}
									hoursData.put(key, cellValue);
								} 

							}
						}
					}
				}
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
		//System.out.println(hoursData);
		for(String key : hoursData.keySet()){
			System.out.println(key+"="+hoursData.get(key));
		}
		System.out.println("count is" + count);
	}

}
