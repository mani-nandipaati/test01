package com.sample.test.exceltest;

/* ====================================================================
Licensed to the Apache Software Foundation (ASF) under one or more
contributor license agreements.  See the NOTICE file distributed with
this work for additional information regarding copyright ownership.
The ASF licenses this file to You under the Apache License, Version 2.0
(the "License"); you may not use this file except in compliance with
the License.  You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
==================================================================== */

import java.io.FileOutputStream;
import java.io.IOException;

//import javafx.scene.paint.Color;



import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
* An example of how to merge regions of cells.
*/
public class TestMerge {
 public static void main(String[] args) throws IOException {
     try (XSSFWorkbook wb = new XSSFWorkbook()) { //or new HSSFWorkbook();
         Sheet sheet = wb.createSheet("new sheet");

         Row row = sheet.createRow((short) 0);
         Cell cell = row.createCell((short) 0);
         cell.setCellValue(new XSSFRichTextString("JPMC CIB - Gbl Client Access"));
         setCellStyle(wb, cell, true, false, false);

         sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));

         // Write the output to a file
         try (FileOutputStream fileOut = new FileOutputStream("merging_cells.xlsx")) {
             wb.write(fileOut);
         }
     }
 }
 
 public static void setCellStyle(XSSFWorkbook writeWorkbook, Cell cell, boolean header, boolean hyperLink, boolean sHeader){
		XSSFCellStyle backgroundStyle = writeWorkbook.createCellStyle();
		/*backgroundStyle.setBorderBottom(BorderStyle.THIN);
		backgroundStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderLeft(BorderStyle.THIN);
		backgroundStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderRight(BorderStyle.THIN);
		backgroundStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		backgroundStyle.setBorderTop(BorderStyle.THIN);
		backgroundStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());*/
		//backgroundStyle.setFillBackgroundColor(XSSFColor.toXSSFColor(Color.));((short)001);
		XSSFFont font = writeWorkbook.createFont();
		font.setBold(false);
		if(header || sHeader){
			if(header){
				backgroundStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
				backgroundStyle.setAlignment(HorizontalAlignment.CENTER);
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
 
}
