package com.sample.test.exceltest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
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


public class MainExcel {
	Properties prop = new Properties();
	Properties rateProp = new Properties();
	Properties cols = new Properties();
	MainExcel(){
		InputStream input = null;

		try {
			FileReader reader=new FileReader("src/main/resources/config.properties");
			prop.load(reader);
			reader=new FileReader("src/main/resources/rate.properties");
			rateProp.load(reader);
			reader=new FileReader("src/main/resources/columns.properties");
			cols.load(reader);
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
		MainExcel me = new MainExcel();
		me.readData();
	}

	public void setCellStyle(XSSFWorkbook writeWorkbook, Cell cell, boolean header, boolean hyperLink, boolean sHeader){
		XSSFCellStyle backgroundStyle = writeWorkbook.createCellStyle();
		backgroundStyle.setBorderBottom(BorderStyle.THIN);
		backgroundStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderLeft(BorderStyle.THIN);
		backgroundStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderRight(BorderStyle.THIN);
		backgroundStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderTop(BorderStyle.THIN);
		backgroundStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
		XSSFFont font = writeWorkbook.createFont();
		font.setBold(false);
		if(header || sHeader){
			if(header){
				backgroundStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
			}
			else{
				backgroundStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			}
			backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			font.setBold(true);
		}

		font.setFontName("Calibri");
		font.setFontHeight(9);
		if(hyperLink){
			font.setUnderline(Font.U_SINGLE);
			font.setColor(IndexedColors.BLUE.getIndex());
		}
		backgroundStyle.setFont(font);

		cell.setCellStyle(backgroundStyle);
	}

	public void writeHyperlinkToAnnexure(int column, Row summaryRow, Sheet writeSheet, XSSFWorkbook writeWorkbook, String reference){
		Cell summaryCell = summaryRow.createCell(column);
		CreationHelper helper= writeWorkbook.getCreationHelper();
		summaryCell.setCellValue(reference);
		Hyperlink link2 = helper.createHyperlink(HyperlinkType.DOCUMENT);
		link2.setAddress("'"+ reference +"'!A1");
		summaryCell.setHyperlink(link2);
		setCellStyle(writeWorkbook, summaryCell, false, true, false);
	}

	public void writeFormulaToAnnexure(int column, Row summaryRow, Sheet writeSheet, XSSFWorkbook writeWorkbook, String formula, CellStyle style){
		Cell summaryCell = summaryRow.createCell(column);
		summaryCell.setCellFormula(formula);
		summaryCell.setCellStyle(style);
	}

	public void createAnnexureHeader(XSSFWorkbook writeWorkbook, Row annexureRow, int annexureRowCount, int annexureColumnCount){
		Cell annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Project ID");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Project Name");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Annexure");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Annex Key");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Invoice #");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue(" Gross ");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Invoice #");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("4% Discount");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Invoice#	");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Credit for Unbilled + Others");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue("Net");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		annexureCell.setCellValue(" Verified ");
		setCellStyle(writeWorkbook, annexureCell, true, false, false);

	}
	public void readData(){
		FileInputStream excelFile = null;
		XSSFWorkbook workbook = null;
		FileOutputStream outputStream = null;
		XSSFWorkbook writeWorkbook = null;
		try {
			boolean writeAnnexure = false;
			writeWorkbook = new XSSFWorkbook();
			Sheet annexureWriteSheet = writeWorkbook.createSheet("Summary");
			int annexureRowCount = 1;
			int annexureColumnCount = 1;
			Row annexureRow;
			Cell annexureCell;
			Sheet writeSheet;
			annexureRow = annexureWriteSheet.createRow(annexureRowCount++);
			createAnnexureHeader(writeWorkbook, annexureRow, annexureRowCount, annexureColumnCount);
			File folder = new File("C:\\Users\\559694\\December");
			File[] files  = folder.listFiles();
			for(File file : files){
				int rowCount = 0;
				annexureColumnCount = 1;
				excelFile = new FileInputStream(new File(file.getAbsolutePath()));
				workbook = new XSSFWorkbook(excelFile);
				Sheet datatypeSheet = workbook.getSheetAt(0);
				String sheetName = file.getName().replaceAll("15398-GCA - ", "");
				sheetName  = sheetName.substring(0, sheetName.indexOf("-"));
				writeSheet = writeWorkbook.createSheet(sheetName);
				annexureRow = annexureWriteSheet.createRow(annexureRowCount++);
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				annexureCell.setCellValue("1000180826");
				setCellStyle(writeWorkbook, annexureCell, false, false, false);
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				annexureCell.setCellValue("JPMC CIB - Gbl Client Access");
				setCellStyle(writeWorkbook, annexureCell, false, false, false);
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				writeHyperlinkToAnnexure(annexureColumnCount-1, annexureRow, annexureWriteSheet, writeWorkbook, sheetName);
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				annexureCell.setCellValue("Ak");
				setCellStyle(writeWorkbook, annexureCell, false, false, false);
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				annexureCell.setCellValue("Ak");
				setCellStyle(writeWorkbook, annexureCell, false, false, false);
				List<Integer> validCols = new ArrayList<Integer>();
				for (int rn=datatypeSheet.getFirstRowNum(); rn<=datatypeSheet.getLastRowNum(); rn++) {
					Row row = datatypeSheet.getRow(rn);
					Row writerow = writeSheet.createRow(rn);
					rowCount++;
					if(rowCount ==1){
						Cell writecell = writerow.createCell(0);
						writecell.setCellValue("S.No");
						setCellStyle(writeWorkbook, writecell, false, false, true);
						writecell = writerow.createCell(1);
						writecell.setCellValue("Project ID");
						setCellStyle(writeWorkbook, writecell, false, false, true);
						writecell = writerow.createCell(2);
						writecell.setCellValue("Project Name");
						setCellStyle(writeWorkbook, writecell, false, false, true);
						writecell = writerow.createCell(3);
						writecell.setCellValue("Milestone");
						setCellStyle(writeWorkbook, writecell, false, false, true);
						writecell = writerow.createCell(4);
						writecell.setCellValue("Beeline ID");
						setCellStyle(writeWorkbook, writecell, false, false, true);
						writecell = writerow.createCell(5);
						writecell.setCellValue("Assoc ID");
						setCellStyle(writeWorkbook, writecell, false, false, true);
					}
					else{
						Cell cell = row.getCell(1);
						if(cell != null && cell.getCellTypeEnum() != CellType.BLANK){
							CellStyle style = writeWorkbook.createCellStyle();
							style.cloneStyleFrom(cell.getCellStyle());
							Cell writecell = writerow.createCell(0);
							writecell.setCellValue(rowCount-1);
							writecell.setCellStyle(style);
							writecell = writerow.createCell(1);
							writecell.setCellValue("1000180826");
							writecell.setCellStyle(style);
							writecell = writerow.createCell(2);
							writecell.setCellValue("JPMC CIB - Gbl Client Access");
							writecell.setCellStyle(style);
							writecell = writerow.createCell(3);
							writecell.setCellValue(sheetName);
							writecell.setCellStyle(style);
							writecell = writerow.createCell(4);
							writecell.setCellValue("15398");
							writecell.setCellStyle(style);
						}
					}
					if (row == null) {
						// There is no data in this row, handle as needed
					} else {

						// Row "rn" has data
						for (int cn=1; cn<row.getLastCellNum(); cn++) {
							Cell cell = row.getCell(cn);
							Cell writecell = writerow.createCell(cn+5);
							if( (rowCount ==1 && !validCols.contains(cn)) || (rowCount >1 && validCols.contains(cn)) ){
								if (cell == null || cell.getCellTypeEnum() == CellType.BLANK) {
									// This cell is empty/blank/un-used, handle as needed
								} else {
									System.out.println(sheetName);
									//String cellStr = fmt.formatCell(cell);
									// Do something with the value
									CellStyle style = writeWorkbook.createCellStyle();
									style.cloneStyleFrom(cell.getCellStyle());
									writecell.setCellStyle(style);			
									if (cell.getCellTypeEnum() == CellType.STRING) {
										String cellValue = (String)cell.getStringCellValue();
										if(rowCount ==1){
											if(cols.getProperty(cellValue) != null){
												writecell.setCellValue(cols.getProperty(cellValue));
												validCols.add(cn);
												setCellStyle(writeWorkbook, writecell, false, false, true);
											}
										}
										else{
											writecell.setCellValue(cellValue);
										}
										if(cn ==1 && rowCount!=1){
											Cell zeroCell = writerow.createCell(5);
											zeroCell.setCellStyle(style);
											zeroCell.setCellValue(prop.getProperty(cellValue));
										}
										if("Total Amount".equalsIgnoreCase(cellValue) ){
											writeAnnexure = true;
										}
										else if("4% Flat discount".equalsIgnoreCase(cellValue)){
											writeAnnexure = true;
											for(short cnt =0; cnt <1; cnt++){
												annexureCell = annexureRow.createCell(annexureColumnCount++);
												annexureCell.setCellValue("");
												setCellStyle(writeWorkbook, annexureCell, false, false, false);
											}

										}
										else if("Net Amount".equalsIgnoreCase(cellValue)){
											writeAnnexure = true;
											for(short cnt =0; cnt <2; cnt++){
												annexureCell = annexureRow.createCell(annexureColumnCount++);
												annexureCell.setCellValue("");
												setCellStyle(writeWorkbook, annexureCell, false, false, false);
											}
										}
									} else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
										writecell.setCellValue((Double) cell.getNumericCellValue());
									}
									else if (cell.getCellTypeEnum() == CellType.FORMULA) {

										//writecell.setCellFormula(cell.getCellFormula());
										writecell.setCellValue(cell.getNumericCellValue());
										if(writeAnnexure){
											System.out.println("writeAnnexure true");
											writeAnnexure = false;
											String formula ="'"+ sheetName + "'!"+writecell.getAddress().toString();
											writeFormulaToAnnexure(annexureColumnCount++, annexureRow, annexureWriteSheet, writeWorkbook, formula, style);

										}
									} 
									else{
										writecell.setCellValue("");
									}
									//if(rowCount ==1){
										//setCellStyle(writeWorkbook, writecell, false, false, true);
									//}
								}
							}
						}
					}
				}
			}
			writeSheet = writeWorkbook.createSheet("Rate");
			int rowCnt =0;
			for(Object keys : rateProp.keySet()){
				Row writerow = writeSheet.createRow(rowCnt++);
				Cell writecell = writerow.createCell(0);
				String key = (String)keys;
				if(rowCnt ==1){
					writecell.setCellValue("Billing Rate");
					setCellStyle(writeWorkbook, writecell, false, false, true);
				}
				else{
					writecell.setCellValue(key);
				}
				writecell = writerow.createCell(1);
				if(rowCnt ==1){
					writecell.setCellValue("Project Role");
					setCellStyle(writeWorkbook, writecell, false, false, true);
				}
				else{
					writecell.setCellValue(rateProp.getProperty(key));
				}
			}
			outputStream = new FileOutputStream("C:\\Users\\559694\\December\\Annexure.xlsx");
			writeWorkbook.write(outputStream);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		catch(Exception e){
			e.printStackTrace();
		}
		finally{
			if(excelFile != null){
				try {
					workbook.close();
					excelFile.close();
					writeWorkbook.close();
					outputStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}

}
