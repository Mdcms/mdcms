package com.md.mdcms.xlsx;

/*
 * Copyright 2017 Midrange Dynamics GmbH. All Rights reserved.
 *
 * This software is the proprietary information of GmbH
 * Use is subject to license and non-disclosure terms.
 */

/**
 * Michael Morgan
 * 21.11.2017
 */

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSVtoXLSX {

//	private static WritableCellFormat integerFormat = new WritableCellFormat(
//			NumberFormats.INTEGER);
//
//	private static WritableCellFormat floatFormat = new WritableCellFormat(
//			NumberFormats.THOUSANDS_FLOAT);
//
//	private static WritableFont arial9Font = new WritableFont(
//			WritableFont.ARIAL, 9);
//
//	private static WritableCellFormat arial9Format = new WritableCellFormat(
//			arial9Font);
//
//	private static WritableFont arial9TotalFont = new WritableFont(
//			WritableFont.ARIAL, 9, WritableFont.BOLD, true);
//	
//	private static WritableCellFormat arial9TotalIntegerFormat = new WritableCellFormat(
//			arial9TotalFont, NumberFormats.THOUSANDS_INTEGER);
//
//	private static WritableCellFormat arial9TotalFloatFormat = new WritableCellFormat(
//			arial9TotalFont, NumberFormats.THOUSANDS_FLOAT);
//	
//	private static WritableFont courier11BoldBlackFont = new WritableFont(
//			WritableFont.COURIER, 11, WritableFont.BOLD, false,
//			UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
//	
//	private static WritableFont arial10BoldBlackFont = new WritableFont(
//			WritableFont.ARIAL, 10, WritableFont.BOLD, false,
//			UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
//
//	private static WritableCellFormat headerFormat = new WritableCellFormat(
//			courier11BoldBlackFont);
//	
//	private static WritableCellFormat footerFormat = new WritableCellFormat(
//			courier11BoldBlackFont);
//
//	private static WritableCellFormat colHeaderFormat = new WritableCellFormat(
//			arial10BoldBlackFont);
//
//	//	private static String[] columnsToExclude;
//
//	private static HashMap columnWidth = new HashMap();

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// check if correct number of arguments were passed
		if (args.length != 23) {
			System.out.println("Invalid number of Parameters passed");
			System.out.println("Parameters expected = 23");
			System.out.println("Parameters passed = " + args.length);
			System.out.println("Expected parameters: ");
			System.out.println("1) csv file path");
			System.out.println("2) excel file path");
			System.out.println("3) Header1");
			System.out.println("4) Header2");
			System.out.println("5) Header3");
			System.out.println("6) Header4");
			System.out.println("7) Header5");
			System.out.println("8) Header6");
			System.out.println("9) Header7");
			System.out.println("10) Header8");
			System.out.println("11) Header9");
			System.out.println("12) Footer1");
			System.out.println("13) Footer2");
			System.out.println("14) Footer3");
			System.out.println("15) Footer4");
			System.out.println("16) Footer5");
			System.out.println("17) Footer6");
			System.out.println("18) Field Types1");
			System.out.println("19) Field Types2");
			System.out.println("20) Field Types3");
			System.out.println("21) Field Types4");
			System.out.println("22) Date order");
			System.out.println("23) Date Separator");
			System.exit(1);
		}

		try {
			System.setProperty("java.awt.headless", "true");
			File csvFile = new File(args[0]);
			File xlsFile = new File(args[1]);
			String dateOrder = args[21];
			String dateSep = args[22];

			// prep CSV
			String lineIn;
			BufferedReader br = new BufferedReader(new FileReader(csvFile));

			// Workbook Settings
			Workbook wb = new XSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet1");

			
//			wb.setLocale(new Locale("en", "EN"));
//			WritableWorkbook workbook = Workbook.createWorkbook(xlsFile, ws);
//			WritableSheet sheet = workbook.createSheet("Table1", 0);
//			SheetSettings settings = new SheetSettings(sheet);
//			settings.setFitToPages(true);
//			settings.setPaperSize(PaperSize.A4);
//			settings.setOrientation(PageOrientation.LANDSCAPE);

			// set cell formats			
//			arial9Format.setShrinkToFit(true);
//			arial9Format.setWrap(true);
//			arial9TotalIntegerFormat.setShrinkToFit(true);
//			arial9TotalIntegerFormat.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
//			arial9TotalFloatFormat.setShrinkToFit(true);
//			arial9TotalFloatFormat.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
//			colHeaderFormat.setBackground(Colour.GREY_25_PERCENT);
//			colHeaderFormat.setWrap(true);
//			colHeaderFormat.setShrinkToFit(true);
//			colHeaderFormat.setVerticalAlignment(VerticalAlignment.TOP);
//			floatFormat.setShrinkToFit(true);
//			floatFormat.setWrap(true);
//			integerFormat.setShrinkToFit(true);
//			integerFormat.setWrap(true);
//			headerFormat.setBackground(Colour.GREY_25_PERCENT);
//			headerFormat.setShrinkToFit(true);
//			footerFormat.setBackground(Colour.GREY_25_PERCENT);
//			footerFormat.setShrinkToFit(true);

			// cell(column, row)
			int colnr = 0;
			int rownr = 0;
			int firstHeaderRow = 0;

			// Headers
			String[] header = {args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9], args[10]};
			int lastHeaderRow = firstHeaderRow;
			boolean headerFound = false;
			for (int j = 8; j > -1; j--) {
				if (!"".equals(header[j].trim()) || (headerFound)) {
					Label label = new Label(0, j + firstHeaderRow, header[j].replaceAll("\\s+$",""),
							headerFormat);
					sheet.addCell(label);
					if (!headerFound) {
						headerFound = true;
						lastHeaderRow = j + firstHeaderRow;
					}
				}
			}

			// fill field type list
			String fieldTypes = args[17].trim() + args[18].trim() + args[19].trim() + args[20].trim();
			String[] fieldType = fieldTypes.split(",");

			// table data
			int columnHeadingRow = lastHeaderRow;
			if (headerFound) {
				columnHeadingRow += 2;
			}
			rownr = columnHeadingRow;
			int firstDataRow = 0;
			int lastDataRow = 0;
			String[] char13 = {"m", "w", "A", "B", "C", "D", "E", "G", "H", "K", "M", "N", "O", "P", "Q", "R", "S", "U", "V", "W"};
			double factor = 1.0;
			int width;
			double w;
			double charWidth;
			String value;

			// loop through CSV lines
			lineIn = br.readLine();
			while (lineIn != null && !"".equals(lineIn)) {
				String[] fields = lineIn.split("\t");

				// loop through columns in line
				for (int i = 0; (i < fields.length && i < fieldType.length); i++) {
					if (!fieldType[i].substring(0, 1).equals("E")) {
						value = fields[i];
						value = value.replaceAll("\"", "").trim();

						// column heading
						if (rownr == columnHeadingRow) {
							factor = 1.3;
							Label label = new Label(colnr, rownr, value, colHeaderFormat);
							sheet.addCell(label);
						}

						// column data
						else {
							factor = 1;
							if (firstDataRow == 0) {
								firstDataRow = rownr;
							} 
							lastDataRow = rownr;

							// date field
							if (fieldType[i].equals("D")) {
								if (value.length() == 6) {
									if (dateOrder.equals("DMY")) {
										value = value.substring(4, 6) + dateSep + value.substring(2, 4) + dateSep + value.substring(0, 2);
									} else {
										if (dateOrder.equals("MDY")) {
											value = value.substring(2, 4) + dateSep + value.substring(4, 6) + dateSep + value.substring(0, 2);
										} else {
											value = value.substring(0, 2) + dateSep + value.substring(2, 4) + dateSep + value.substring(4, 6);
										}
									}
								}
								if (value.length() == 8) {
									if (dateOrder.equals("DMY")) {
										value = value.substring(6, 8) + dateSep + value.substring(4, 6) + dateSep + value.substring(0, 4);
									} else {
										if (dateOrder.equals("MDY")) {
											value = value.substring(2, 4) + dateSep + value.substring(4, 6) + dateSep + value.substring(0, 4);
										} else {
											value = value.substring(0, 4) + dateSep + value.substring(4, 6) + dateSep + value.substring(6, 8);
										}
									}
								}
								Label label = new Label(colnr, rownr, value, arial9Format);
								sheet.addCell(label);
							}
						
							// floating point field
							if (fieldType[i].substring(0, 1).equals("F")) {
								try {
									double doubleValue = Double.valueOf(value).doubleValue();
									Number number = new Number(colnr, rownr, doubleValue, floatFormat);
									sheet.addCell(number);
								} catch (Exception e) {
									Label label = new Label(colnr, rownr, value, arial9Format);
									sheet.addCell(label);
								}
							}

							// integer field
							if (fieldType[i].substring(0, 1).equals("I")) {
								try {
									int integerValue = Integer.valueOf(value).intValue();
									Number number = new Number(colnr, rownr, integerValue, integerFormat);
									sheet.addCell(number);
								} catch (Exception e) {
									Label label = new Label(colnr, rownr, value, arial9Format);
									sheet.addCell(label);
								}
							}

							// string field
							if (fieldType[i].equals("S")) {
								Label label = new Label(colnr, rownr, value, arial9Format);
								sheet.addCell(label);
							}
						}

						//		calculate cell width and add column number
						w = 1;
						for (int j = 0; j < value.length(); j++) {
							charWidth = 1;
							for (int k = 0; k < char13.length; k++) {
								if (char13[k].equals(value.substring(j, j + 1))) {
									charWidth = 1.3;
									k = char13.length;
								}
							}
							w = w + (charWidth * factor);
						}
						width = Double.valueOf(String.valueOf(w)).intValue();
						if (width > 80) {
							width = 80;
						}
						concludeColumnWidth(colnr, width);
						colnr++;
					}
				}
				lineIn = br.readLine();
				colnr = 0;
				rownr++;
			}

			for (Iterator iterator = columnWidth.keySet().iterator(); iterator
					.hasNext();) {
				Integer col = (Integer) iterator.next();
				sheet.setColumnView(col.intValue(), ((Integer) columnWidth
						.get(col)).intValue());
			}

			// total row
			colnr = 0;
			int columnCount = 5;
			for (int i = 0; i < fieldType.length; i++) {
				if (!fieldType[i].substring(0, 1).equals("E")) {
					if (fieldType[i].length() > 1) {
						if (fieldType[i].substring(1, 2).equals("T")) {
							String firstCell;
							String lastCell;
							firstCell = CellReferenceHelper.getCellReference(colnr, firstDataRow);
							lastCell = CellReferenceHelper.getCellReference(colnr, lastDataRow);
							value = "SUM(" + firstCell + ":" + lastCell + ")";
							if (fieldType[i].substring(0, 1).equals("F")) {
								Formula formula = new Formula(colnr, rownr, value, arial9TotalFloatFormat);
								sheet.addCell(formula);
							} else {
								Formula formula = new Formula(colnr, rownr, value, arial9TotalIntegerFormat);
								sheet.addCell(formula);
							}
						}
					}
					colnr++;
					if (colnr > columnCount) {
						columnCount = colnr;
					}
				}
			}
			
			// merge the header cells
			if (headerFound) {
				for (int i = firstHeaderRow; i <= lastHeaderRow; i++) {
					sheet.mergeCells(0, i, columnCount - 1, i);
				}
			}

			// Footers
			String[] footer = {args[11], args[12], args[13], args[14], args[15], args[16]};
			boolean footerFound = false;
			rownr++;
			for (int j = 5; j > -1; j--) {
				if (!"".equals(footer[j].trim()) || (footerFound)) {
					Label label = new Label(0, rownr + j, footer[j].replaceAll("\\s+$", ""), footerFormat);
					sheet.addCell(label);
					sheet.mergeCells(0, rownr + j, columnCount - 1, rownr + j);
					footerFound = true;
				}
			}
			
			// write workbook to file
			if (xlsFile.exists()) {
				xlsFile.delete();
			}
			FileOutputStream fileOut = new FileOutputStream(xlsFile);
			wb.write(fileOut);
			fileOut.close();

		} catch (UnsupportedEncodingException e) {
			System.out.println(e.toString());
			System.exit(1);
		} catch (IOException e) {
			System.out.println(e.toString());
			System.exit(1);
		} catch (Exception e) {
			System.out.println(e.toString());
			System.exit(1);
		}
	}

	private static void concludeColumnWidth(int colnr, int width) {
		Integer savedColWidth = (Integer) columnWidth.get(new Integer(colnr));
		if (savedColWidth != null) {
			if (savedColWidth.intValue() < width) {
				columnWidth.put(new Integer(colnr), new Integer(width));
			}
		} else {
			columnWidth.put(new Integer(colnr), new Integer(width));
		}
	}
}