package com.sample.test.exceltest;

import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.Arrays;
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
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class CopyOfMainExcel {
	Properties prop = new Properties();
	Properties rateProp = new Properties();
	Properties cols = new Properties();
	List<String> headerCols;
	double annexureDiscountAmount = 0.0;
	double annexureGrossAmount = 0.0;
	double annexureNetAmount = 0.0;
	
	CopyOfMainExcel(){
		InputStream input = null;

		try {
			FileReader reader=new FileReader("src/main/resources/config.properties");
			prop.load(reader);
			reader=new FileReader("src/main/resources/rate.properties");
			rateProp.load(reader);
			InputStream is = new FileInputStream("src/main/resources/columns.properties");
			cols.load(new DataInputStream(is));
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
		CopyOfMainExcel me = new CopyOfMainExcel();
		Properties headerProps = new Properties();
		InputStream is = new FileInputStream("src/main/resources/header.properties");
		headerProps.load(new DataInputStream(is));
		String header = headerProps.getProperty("HeaderSequence");
		String[] st = header.split(",");
		me.headerCols = Arrays.asList(st);
		me.readData();
	}

	public void setCellStyle(XSSFWorkbook annexureWorkbook, Cell cell, short colorCode, 
			boolean hyperLink, short currencyFormat){
		XSSFCellStyle backgroundStyle = annexureWorkbook.createCellStyle();
		backgroundStyle.setBorderBottom(BorderStyle.THIN);
		backgroundStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderLeft(BorderStyle.THIN);
		backgroundStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderRight(BorderStyle.THIN);
		backgroundStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderTop(BorderStyle.THIN);
		backgroundStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setWrapText(false);
		XSSFFont font = annexureWorkbook.createFont();
		font.setBold(false);
		if(colorCode >=0){
			backgroundStyle.setFillForegroundColor(colorCode);
			backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			backgroundStyle.setAlignment(HorizontalAlignment.CENTER);
			font.setBold(true);
		}
		
		font.setFontName("Calibri");
		font.setFontHeight(9);
		if(hyperLink){
			font.setUnderline(Font.U_SINGLE);
			font.setColor(IndexedColors.BLUE.getIndex());
		}
		backgroundStyle.setFont(font);

		if(currencyFormat > 0){
			backgroundStyle.setDataFormat(currencyFormat);
		}
		cell.setCellStyle(backgroundStyle);
	}

	public void writeHyperlinkToAnnexure(int column, Row summaryRow, Sheet milestoneSheet, XSSFWorkbook annexureWorkbook, String reference){
		Cell summaryCell = summaryRow.createCell(column);
		CreationHelper helper= annexureWorkbook.getCreationHelper();
		summaryCell.setCellValue(reference);
		Hyperlink link2 = helper.createHyperlink(HyperlinkType.DOCUMENT);
		link2.setAddress("'"+ reference +"'!A1");
		summaryCell.setHyperlink(link2);
		//setCellStyle(annexureWorkbook, summaryCell, false, true, false);
		setCellStyle(annexureWorkbook, summaryCell, (short)-1, true, (short)0); 
	}

	public void writeFormulaToAnnexure(int column, Row summaryRow, Sheet milestoneSheet, XSSFWorkbook annexureWorkbook, String formula, CellStyle style){
		Cell summaryCell = summaryRow.createCell(column);
		summaryCell.setCellFormula(formula);
		summaryCell.setCellStyle(style);
	}

	public void createAnnexureHeader(Sheet summarySheet, XSSFWorkbook annexureWorkbook, int annexureRowCount, int annexureColumnCount){
		Row row = summarySheet.createRow((short) 0);

		createMergedCells(summarySheet, row, annexureWorkbook, "JPMC CIB - Gbl Client Access", IndexedColors.SEA_GREEN.getIndex(),0,0,0,11,(short)0);

		row = summarySheet.createRow((short) 1);
        createMergedCells(summarySheet, row, annexureWorkbook, "Project Details", IndexedColors.PALE_BLUE.getIndex(),1, 1, 0, 1,(short)0);

        createMergedCells(summarySheet, row, annexureWorkbook, "Annexure Details",IndexedColors.PALE_BLUE.getIndex(), 1,1,2,3,(short)0);
        
        createMergedCells(summarySheet, row, annexureWorkbook, "Gross Amount", IndexedColors.PALE_BLUE.getIndex(),1,1,4,5,(short)0);
        
        createMergedCells(summarySheet, row, annexureWorkbook, "Discount Amount",  IndexedColors.PALE_BLUE.getIndex(),1,1,6,9,(short)0);
        
        row = summarySheet.getRow((short) 1);
        Cell cell = row.createCell((short) 10);
        cell.setCellValue(new XSSFRichTextString("Net Amount"));
        setCellStyle(annexureWorkbook, cell, IndexedColors.PALE_BLUE.getIndex(), false, (short)0);
        
        row = summarySheet.createRow((short) 2);
		String[] headerCols = {"Project ID","Project Name","Annexure","Annex Key","Invoice #","Gross","Invoice#","4% Discount","Invoice#","Credit for Unbilled + Others","Net","Verified"};
		for(String headerCol : headerCols){
			Cell annexureCell = row.createCell(annexureColumnCount++);
			summarySheet.autoSizeColumn(annexureColumnCount);
			annexureCell.setCellValue(headerCol);
			setCellStyle(annexureWorkbook, annexureCell, IndexedColors.AQUA.getIndex(), false, (short)0);
		}
	}
	private void createMergedCells(Sheet summarySheet, Row row,
			XSSFWorkbook annexureWorkbook, String cellValue, short cellColor, int firstRow, int lastRow, int firstCol, int lastCol,
			short currencyFormat) {
		Cell cell;
        cell = row.createCell((short) firstCol);
        summarySheet.autoSizeColumn(firstCol);
        if(currencyFormat >0){
        	cell.setCellValue(Double.valueOf(cellValue));
        }
        else{
        	cell.setCellValue(new XSSFRichTextString(cellValue));
        }
        setCellStyle(annexureWorkbook, cell, cellColor, false, currencyFormat);
        setFooterBorderStyle(summarySheet, firstRow, lastRow, firstCol, lastCol );
	}
	public void readData(){
		FileInputStream excelFile = null;
		XSSFWorkbook workbook = null;
		FileOutputStream outputStream = null;
		XSSFWorkbook annexureWorkbook = null;
		try {
			boolean writeAnnexure = false;
			annexureWorkbook = new XSSFWorkbook();
			Sheet summarySheet = annexureWorkbook.createSheet("Summary");
			int annexureRowCount = 3;
			int annexureColumnCount = 0;
			Row annexureRow;
			Cell annexureCell;
			Sheet milestoneSheet;
			//annexureRow = summarySheet.createRow(annexureRowCount++);
			createAnnexureHeader(summarySheet, annexureWorkbook, annexureRowCount, annexureColumnCount);
			File folder = new File("I:\\AnnexureWorkspace\\December\\December");
			File[] files  = folder.listFiles();
			for(File file : files){
				int rowCount = 0;
				annexureColumnCount = 0;
				excelFile = new FileInputStream(new File(file.getAbsolutePath()));
				workbook = new XSSFWorkbook(excelFile);
				Sheet datatypeSheet = workbook.getSheetAt(0);
				String sheetName = file.getName().replaceAll("15398-", "");
				sheetName  = sheetName.substring(0, sheetName.indexOf("-Template"));
				milestoneSheet = annexureWorkbook.createSheet(sheetName);
				annexureRow = summarySheet.createRow(annexureRowCount++);
				summarySheet.autoSizeColumn(annexureColumnCount);
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				summarySheet.autoSizeColumn(annexureColumnCount);
				annexureCell.setCellValue("1000180826");
				//setCellStyle(annexureWorkbook, annexureCell, false, false, false);
				setCellStyle(annexureWorkbook, annexureCell, (short)-1, false, (short)0); 
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				summarySheet.autoSizeColumn(annexureColumnCount);
				annexureCell.setCellValue("JPMC CIB - Gbl Client Access");
				//setCellStyle(annexureWorkbook, annexureCell, false, false, false);
				setCellStyle(annexureWorkbook, annexureCell, (short)-1, false, (short)0); 
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				summarySheet.autoSizeColumn(annexureColumnCount);
				writeHyperlinkToAnnexure(annexureColumnCount-1, annexureRow, summarySheet, annexureWorkbook, sheetName);
				annexureCell = annexureRow.createCell(annexureColumnCount++);
				summarySheet.autoSizeColumn(annexureColumnCount);
				annexureCell.setCellValue(prop.getProperty(sheetName));
				//setCellStyle(annexureWorkbook, annexureCell, false, false, false);
				setCellStyle(annexureWorkbook, annexureCell, (short)-1, false, (short)0); 
				int milestoneHeaderColumnCount = 0;
				Row writerow = milestoneSheet.createRow(rowCount++);
				for(String key : headerCols){
					Cell writecell = writerow.createCell(milestoneHeaderColumnCount++);
					milestoneSheet.autoSizeColumn(milestoneHeaderColumnCount);
					writecell.setCellValue(key);
					//setCellStyle(annexureWorkbook, writecell, false, false, true);
					setCellStyle(annexureWorkbook, writecell, IndexedColors.GREY_25_PERCENT.getIndex(), false, (short)0); 
				}
				Double totalLeaveDays = 0.0;
				Double totalBillingLoss = 0.0;
				Double totalBillableHours = 0.0;
				Double totalAmount = 0.0;
				for (int rn=datatypeSheet.getFirstRowNum()+1; rn<=datatypeSheet.getLastRowNum(); rn++) {
					Row row = datatypeSheet.getRow(rn);
					writerow = milestoneSheet.createRow(rn);
					rowCount++;
					milestoneHeaderColumnCount = 0;
					Double leaveDays = 0.0 ;
					for(String key : headerCols){
						Cell writecell = writerow.createCell(milestoneHeaderColumnCount++);
						milestoneSheet.autoSizeColumn(milestoneHeaderColumnCount);
						String strValue = cols.getProperty(key);
						System.out.println(sheetName+" "+key+" "+strValue);
						if( "-1".equals(strValue) || "CONST".equals(strValue) || "FORMULA".equals(strValue) ){
							if(datatypeSheet.getRow(rn) != null &&
									datatypeSheet.getRow(rn).getCell(0) != null && 
									datatypeSheet.getRow(rn).getCell(0).getCellTypeEnum() != CellType.BLANK){
								if("S.No".equalsIgnoreCase(key)){
									writecell.setCellValue(rn);
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
								}else if("Project ID".equalsIgnoreCase(key)){
									writecell.setCellValue(Long.valueOf("1000180826"));
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
								}else if("Project Name".equalsIgnoreCase(key)){
									writecell.setCellValue("JPMC CIB - Gbl Client Access");
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
								}
								else if("Assoc ID".equalsIgnoreCase(key)){
									Cell sidCell = datatypeSheet.getRow(rn).getCell(Integer.valueOf(cols.getProperty("SID")));
									if(sidCell != null && 
											sidCell.getCellTypeEnum() != CellType.BLANK){
										String employeeId = prop.getProperty(sidCell.getStringCellValue());
										if(employeeId != null){
											writecell.setCellValue(Integer.valueOf(employeeId));
										}
										//setCellStyle(annexureWorkbook, writecell, false, false, false);
										setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
									}
								}
								/*else if("Leave (Days)".equalsIgnoreCase(key)){
									Cell billableDaysCell = datatypeSheet.getRow(rn).getCell(Integer.valueOf(cols.getProperty("Billable Days")));
									if(billableDaysCell != null && 
											billableDaysCell.getCellTypeEnum() != CellType.BLANK){
										Cell readCell = row.getCell(Integer.valueOf(cols.getProperty("Billable Days")));
										Double workingHours = (Double) readCell.getNumericCellValue();
										DecimalFormat df = new DecimalFormat("#.#");
										Double days = workingHours/8;
										leaveDays  = Double.valueOf(prop.getProperty("WorkingDays"))-days;
										totalLeaveDays+=leaveDays;
										writecell.setCellValue(Double.valueOf(df.format(leaveDays)));
										setCellStyle(annexureWorkbook, writecell, false, false, false);
									}
								}*/
								else if("Billable Loss".equalsIgnoreCase(key)){
									if(leaveDays  > 0){
										Cell readCell = row.getCell(Integer.valueOf(cols.getProperty("Rate/Hr")));
										Double rate = (Double) readCell.getNumericCellValue();
										double billingLoss = rate*8*leaveDays;
										totalBillingLoss+=billingLoss;
										writecell.setCellValue(billingLoss);
										//(annexureWorkbook, writecell);
										setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)8);
									}
									else{
										writecell.setCellValue(0);
										//setCellStyle(annexureWorkbook, writecell, false, false, false);
										setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
									}
								}
								else if("Project Role".equalsIgnoreCase(key)){
									writecell.setCellFormula("VLOOKUP(J"+(rn+1)+",Rate!A:B,2,FALSE)");
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
								}
								else if("Rate/day".equalsIgnoreCase(key)){
									Cell readCell = row.getCell(Integer.valueOf(cols.getProperty("Rate/Hr")));
									Double rate = (Double) readCell.getNumericCellValue();
									writecell.setCellValue(rate*8);
									//setCellBorder(annexureWorkbook, writecell);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)8);
								}
								else if("Billable Hrs".equalsIgnoreCase(key)){
									Cell readCell = row.getCell(Integer.valueOf(cols.getProperty("Billable Days")));
									Double workingHours = (Double) readCell.getNumericCellValue();
									Double workingDays = workingHours/8;
									readCell = row.getCell(Integer.valueOf(cols.getProperty("Over time")));
									Double overTime = (Double) readCell.getNumericCellValue();
									Double billableHours = workingDays*8+overTime;
									totalBillableHours+=billableHours;
									writecell.setCellValue(billableHours);
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
								}
								else if("Annexure".equalsIgnoreCase(key)){
									writecell.setCellValue(prop.getProperty(sheetName));
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
								}
								else {
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
								}

							}

						}
						else{
							Cell readCell = row.getCell(Integer.valueOf(strValue));
							if(readCell != null && readCell.getCellTypeEnum() != CellType.BLANK){
								CellStyle style = annexureWorkbook.createCellStyle();
								style.cloneStyleFrom(readCell.getCellStyle());
								style.setWrapText(false);
								writecell.setCellStyle(style);
								if (readCell.getCellTypeEnum() == CellType.STRING) {
									String cellValue = (String)readCell.getStringCellValue();
									if("Location".equalsIgnoreCase(key)){
										if( "Bengaluru".equalsIgnoreCase(cellValue) || "Mumbai".equalsIgnoreCase(cellValue) ){
											cellValue = "Offshore";
										}
										else{
											cellValue = "Onsite";	
										}
									}
									writecell.setCellValue(cellValue);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
								} else if (readCell.getCellTypeEnum() == CellType.NUMERIC) {
									Double doubleValue = (Double) readCell.getNumericCellValue();
									writecell.setCellValue(doubleValue);
									if("Billable Days".equalsIgnoreCase(key)){
										DecimalFormat df = new DecimalFormat("#.#");
										writecell.setCellValue(Double.valueOf(df.format(doubleValue/8)));
									}
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
								}
								else if (readCell.getCellTypeEnum() == CellType.FORMULA) {
									Double doubleValue = (Double) readCell.getNumericCellValue();
									writecell.setCellValue(doubleValue);
									if("$ Amount".equalsIgnoreCase(key)){
										if(datatypeSheet.getRow(rn).getCell(0) != null && 
												datatypeSheet.getRow(rn).getCell(0).getCellTypeEnum() != CellType.BLANK){
											totalAmount += doubleValue;
										}
										setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)8);
									}
									else{
										setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
									}
								}
							}
							else  {
								if(datatypeSheet.getLastRowNum()-3 >=rn){
									//setCellStyle(annexureWorkbook, writecell, false, false, false);
									setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
								}
							}
						}
					}

				}
				setFooterRows(annexureWorkbook, milestoneSheet, totalLeaveDays,
						totalBillingLoss, totalBillableHours, totalAmount);
				int lastRow = milestoneSheet.getLastRowNum();
				int amountCol = Integer.valueOf(prop.getProperty("$ Amount"));
				annexureColumnCount = writeAnnexureAmouts(annexureWorkbook,
						annexureColumnCount, annexureRow, milestoneSheet,
						sheetName, lastRow, amountCol, 3);

				annexureColumnCount = writeAnnexureAmouts(annexureWorkbook,
						annexureColumnCount, annexureRow, milestoneSheet,
						sheetName, lastRow, amountCol, 2);

				annexureColumnCount = writeAnnexureAmouts(annexureWorkbook,
						annexureColumnCount, annexureRow, milestoneSheet,
						sheetName, lastRow, amountCol, 1);

				annexureColumnCount = writeAnnexureAmouts(annexureWorkbook,
						annexureColumnCount, annexureRow, milestoneSheet,
						sheetName, lastRow, amountCol, 0);
			}
			
			annexureRow = summarySheet.createRow(annexureRowCount++);
			annexureCell = annexureRow.createCell(annexureColumnCount++);
			
			createMergedCells(summarySheet, annexureRow, annexureWorkbook, "GCA Total",  IndexedColors.LIGHT_GREEN.getIndex(), annexureRowCount-1,annexureRowCount-1,0,3,(short)0);
			annexureCell = annexureRow.createCell(4);
			setCellStyle(annexureWorkbook, annexureCell, IndexedColors.LIGHT_GREEN.getIndex(), false, (short)0);
			annexureCell = annexureRow.createCell(5);
			annexureCell.setCellValue(annexureGrossAmount);
			//setCellStyle(annexureWorkbook, annexureCell, false, true, true);
			setCellStyle(annexureWorkbook, annexureCell, IndexedColors.LIGHT_GREEN.getIndex(), false, (short)8); 
			createMergedCells(summarySheet, annexureRow, annexureWorkbook, Double.toString(annexureDiscountAmount),  IndexedColors.LIGHT_GREEN.getIndex(), annexureRowCount-1,annexureRowCount-1,6,9,(short)8);
			annexureCell = annexureRow.createCell(10);
			annexureCell.setCellValue(annexureNetAmount);
			setCellStyle(annexureWorkbook, annexureCell, IndexedColors.LIGHT_GREEN.getIndex(), false, (short)8); 
			
			milestoneSheet = annexureWorkbook.createSheet("Rate");
			int rowCnt =0;
			for(Object keys : rateProp.keySet()){
				Row writerow = milestoneSheet.createRow(rowCnt++);
				Cell writecell = writerow.createCell(0);
				milestoneSheet.autoSizeColumn(0);
				String key = (String)keys;
				if(rowCnt ==1){
					writecell.setCellValue("Billing Rate");
					//setCellStyle(annexureWorkbook, writecell, false, false, true);
					setCellStyle(annexureWorkbook, writecell, IndexedColors.GREY_25_PERCENT.getIndex(), false, (short)0); 
				}
				else{
					Double rate = 0.0;
					if(!"-".equals(key)){
						rate = Double.valueOf(key);
					}
					writecell.setCellValue(rate);
					//setCellBorder(annexureWorkbook, writecell);
					setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)8);
				}
				writecell = writerow.createCell(1);
				milestoneSheet.autoSizeColumn(1);
				if(rowCnt ==1){
					writecell.setCellValue("Project Role");
					//setCellStyle(annexureWorkbook, writecell, false, false, true);
					setCellStyle(annexureWorkbook, writecell, IndexedColors.GREY_25_PERCENT.getIndex(), false, (short)0); 
				}
				else{
					writecell.setCellValue(rateProp.getProperty(key));
					//setCellStyle(annexureWorkbook, writecell, false, false, false);
					setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
				}
			}
			System.out.println(annexureGrossAmount +" "+ annexureDiscountAmount +" " +annexureNetAmount);
			outputStream = new FileOutputStream("I:\\AnnexureWorkspace\\December\\December\\Annexure.xlsx");
			annexureWorkbook.write(outputStream);
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
					annexureWorkbook.close();
					outputStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
	private int writeAnnexureAmouts(XSSFWorkbook annexureWorkbook,
			int annexureColumnCount, Row annexureRow, Sheet milestoneSheet,
			String sheetName, int lastRow, int amountCol, int amountRow) {
		Cell annexureCell;
		Row writerow;
		writerow = milestoneSheet.getRow(lastRow-amountRow);
		Cell cell = writerow.getCell(amountCol);
		milestoneSheet.autoSizeColumn(amountCol);
		if(amountRow == 3){
			annexureGrossAmount += cell.getNumericCellValue();
		}
		else if (amountRow ==2 || amountRow ==1){
			annexureDiscountAmount += cell.getNumericCellValue();	
		}
		else{
			annexureNetAmount += cell.getNumericCellValue();
		}
		if(amountRow >0){
			milestoneSheet.autoSizeColumn(annexureColumnCount);
			annexureCell = annexureRow.createCell(annexureColumnCount++);
			//setCellBorder(annexureWorkbook, annexureCell);
			setCellStyle(annexureWorkbook, annexureCell, (short)-1, false, (short)8);
		}
		milestoneSheet.autoSizeColumn(annexureColumnCount);
		annexureCell = annexureRow.createCell(annexureColumnCount++);
		String formula ="'"+ sheetName + "'!"+cell.getAddress().toString();
		annexureCell.setCellFormula(formula);
		//setCellBorder(annexureWorkbook, annexureCell);
		setCellStyle(annexureWorkbook, annexureCell, (short)-1, false, (short)8);
		return annexureColumnCount;
	}
	private void setFooterRows(XSSFWorkbook annexureWorkbook,
			Sheet milestoneSheet, Double totalLeaveDays,
			Double totalBillingLoss, Double totalBillableHours,
			Double totalAmount) {
		Row writerow;
		milestoneSheet.shiftRows(milestoneSheet.getLastRowNum()-2, milestoneSheet.getLastRowNum(), 1);
		writerow = milestoneSheet.createRow(milestoneSheet.getLastRowNum()-3);
		Cell writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Leave (Days)")));
		writecell.setCellValue(totalLeaveDays);
		//setCellStyle(annexureWorkbook, writecell, false, false, false);
		setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0); 
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Billable Loss")));
		writecell.setCellValue(totalBillingLoss);
		//setCellStyle(annexureWorkbook, writecell, false, false, false);
		setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Billable Hrs")));
		writecell.setCellValue(totalBillableHours);
		//setCellStyle(annexureWorkbook, writecell, false, false, false);
		setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Project Role")));
		//setCellStyle(annexureWorkbook, writecell, false, false, false);
		setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Rate/day")));
		//setCellStyle(annexureWorkbook, writecell, false, false, false);
		setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("$ Amount")));
		writecell.setCellValue(totalAmount);
		//setCellBorder(annexureWorkbook, writecell);
		setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)8);
		writerow = milestoneSheet.getRow(milestoneSheet.getLastRowNum()-2);
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Leave (Days)")));
		writecell.setCellValue(new XSSFRichTextString("Gross"));
		//setCellStyle(annexureWorkbook, writecell, true, false, false);
		setCellStyle(annexureWorkbook, writecell, IndexedColors.AQUA.getIndex(), false, (short)0);
		setFooterBorderStyle(milestoneSheet, milestoneSheet.getLastRowNum()-2, milestoneSheet.getLastRowNum()-2,
				Integer.valueOf(prop.getProperty("Leave (Days)")), Integer.valueOf(prop.getProperty("Billable Hrs")));

		writerow = milestoneSheet.getRow(milestoneSheet.getLastRowNum()-1);
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Leave (Days)")));
		writecell.setCellValue(new XSSFRichTextString("Onetime Discount"));
		//setCellStyle(annexureWorkbook, writecell, true, false, false);
		setCellStyle(annexureWorkbook, writecell, IndexedColors.AQUA.getIndex(), false, (short)0);
		setFooterBorderStyle(milestoneSheet, milestoneSheet.getLastRowNum()-1, milestoneSheet.getLastRowNum()-1,
				Integer.valueOf(prop.getProperty("Leave (Days)")), Integer.valueOf(prop.getProperty("Billable Hrs")));

		milestoneSheet.shiftRows(milestoneSheet.getLastRowNum(), milestoneSheet.getLastRowNum(), 1);

		writerow = milestoneSheet.createRow(milestoneSheet.getLastRowNum()-1);
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Leave (Days)")));
		writecell.setCellValue(new XSSFRichTextString("Credit for Unbilled"));
		//setCellStyle(annexureWorkbook, writecell, true, false, false);
		setCellStyle(annexureWorkbook, writecell, IndexedColors.AQUA.getIndex(), false, (short)0);
		setFooterBorderStyle(milestoneSheet, milestoneSheet.getLastRowNum()-1, milestoneSheet.getLastRowNum()-1,
				Integer.valueOf(prop.getProperty("Leave (Days)")), Integer.valueOf(prop.getProperty("Billable Hrs")));
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("$ Amount")));
		//setCellStyle(annexureWorkbook, writecell, false, false, false);
		setCellStyle(annexureWorkbook, writecell, (short)-1, false, (short)0);

		writerow = milestoneSheet.getRow(milestoneSheet.getLastRowNum());
		writecell = writerow.createCell(Integer.valueOf(prop.getProperty("Leave (Days)")));
		writecell.setCellValue(new XSSFRichTextString("Net"));
		//setCellStyle(annexureWorkbook, writecell, true, false, false);
		setCellStyle(annexureWorkbook, writecell, IndexedColors.AQUA.getIndex(), false, (short)0);
		setFooterBorderStyle(milestoneSheet, milestoneSheet.getLastRowNum(), milestoneSheet.getLastRowNum(),
				Integer.valueOf(prop.getProperty("Leave (Days)")), Integer.valueOf(prop.getProperty("Billable Hrs")));
	}
	private void setFooterBorderStyle(Sheet milestoneSheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		CellRangeAddress cell = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		RegionUtil.setBorderBottom(BorderStyle.THIN, cell, milestoneSheet);
		RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(),cell, milestoneSheet);
		RegionUtil.setBorderLeft(BorderStyle.THIN, cell, milestoneSheet);
		RegionUtil.setLeftBorderColor(IndexedColors.BLACK.getIndex(),cell, milestoneSheet);
		RegionUtil.setBorderRight(BorderStyle.THIN, cell, milestoneSheet);
		RegionUtil.setRightBorderColor(IndexedColors.BLACK.getIndex(),cell, milestoneSheet);
		RegionUtil.setBorderTop(BorderStyle.THIN, cell, milestoneSheet);
		RegionUtil.setTopBorderColor(IndexedColors.BLACK.getIndex(),cell, milestoneSheet);
		milestoneSheet.addMergedRegion(cell);
	}
}
