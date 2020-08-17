package com.office.pdfConverter.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class excelService {
	
	
	
	private static Font catFont = new Font(Font.FontFamily.TIMES_ROMAN, 18, Font.BOLD);
	private static Font redFont = new Font(Font.FontFamily.TIMES_ROMAN, 12, Font.NORMAL, BaseColor.RED);
	private static Font subFont = new Font(Font.FontFamily.TIMES_ROMAN, 16, Font.BOLD);
	private static Font smallBold = new Font(Font.FontFamily.TIMES_ROMAN, 12, Font.BOLD);
	private static Font plainFont= new Font(Font.FontFamily.TIMES_ROMAN, 10,
            Font.NORMAL);
	private static int numberOfColumns;
	
	
	public static String exceltoPdf() throws IOException, DocumentException {
		String sourcePath = "/home/jinx/excelsample2.xls";
		String destinationPath = "/home/jinx/fileName.pdf";
		FileInputStream input_document = new FileInputStream(new File(sourcePath));
		if (sourcePath.endsWith(".xlsx")) {
			XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document);
			int length = my_xlsx_workbook.getNumberOfSheets();
			Document iText_xlsx_2_pdf = new Document();
			PdfWriter.getInstance(iText_xlsx_2_pdf, new FileOutputStream(destinationPath));
			iText_xlsx_2_pdf.open();
			for (int i = 0; i < length; i++) {
				int temp = 0;
				Sheet sheet = my_xlsx_workbook.getSheetAt(i);
				addTitlePage(iText_xlsx_2_pdf,sheet.getSheetName());
				Iterator<Row> rowIterator = sheet.iterator();
				PdfPTable table = null;
				PdfPCell table_cell;
				boolean flag = true;
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					int cellNumber = 0;
					if (flag) {
						table = new PdfPTable(row.getLastCellNum());
						flag = false;
					}
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							if (temp == 0) {
								numberOfColumns = row.getLastCellNum();
								PdfPCell c1 = new PdfPCell(new Phrase(cell.getStringCellValue(),plainFont));
								c1.setHorizontalAlignment(Element.ALIGN_CENTER);
								c1.setBackgroundColor(BaseColor.GREEN);
								table.addCell(c1);
								table.setHeaderRows(1);
							} else {
								cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell, table);
							}
							cellNumber++;
							break;
						case Cell.CELL_TYPE_NUMERIC:
							cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell, table);
							cellNumber++;
							break;
						}
					}
					temp = 1;
					if (numberOfColumns != cellNumber) {
						for (int j = 0; j < (numberOfColumns - cellNumber); j++) {
							table.addCell(" ");
						}
					}
				}
				table.setWidthPercentage(100);
				iText_xlsx_2_pdf.add(table);
			}
			iText_xlsx_2_pdf.close();
			input_document.close();
		}
		else if(sourcePath.endsWith(".xls")){
			HSSFWorkbook my_xls_workbook = new HSSFWorkbook(input_document);
			int length = my_xls_workbook.getNumberOfSheets();
			Document iText_xls_2_pdf = new Document();
			PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream(destinationPath));
			iText_xls_2_pdf.open();
			for (int i = 0; i < length; i++) {
				int temp = 0;
				Sheet sheet = my_xls_workbook.getSheetAt(i);
				addTitlePage(iText_xls_2_pdf,sheet.getSheetName());
				Iterator<Row> rowIterator = sheet.iterator();
				PdfPTable table = null;
				PdfPCell table_cell;
				boolean flag = true;
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					int cellNumber = 0;
					if (flag) {
						table = new PdfPTable(row.getLastCellNum());
						flag = false;
					}
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							if (temp == 0) {
								numberOfColumns = row.getLastCellNum();
								PdfPCell c1 = new PdfPCell(new Phrase(cell.getStringCellValue(),plainFont));
								c1.setHorizontalAlignment(Element.ALIGN_CENTER);
								c1.setBackgroundColor(BaseColor.GREEN);
								table.addCell(c1);
								table.setHeaderRows(1);
							} else {
								cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell, table);
							}
							cellNumber++;
							break;
						case Cell.CELL_TYPE_NUMERIC:
							cellNumber = checkEmptyCellAndAddCellContentToPDFTable(cellNumber, cell, table);
							cellNumber++;
							break;
						}
					}
					temp = 1;
					if (numberOfColumns != cellNumber) {
						for (int j = 0; j < (numberOfColumns - cellNumber); j++) {
							table.addCell(" ");
						}
					}
				}
				table.setWidthPercentage(100);
				iText_xls_2_pdf.add(table);
			}
			iText_xls_2_pdf.close();
			input_document.close();
		}
		return "pdf created successfully at " + destinationPath;
	}
	private static int checkEmptyCellAndAddCellContentToPDFTable(int cellNumber, Cell cell, PdfPTable table) {
		if (cellNumber == cell.getColumnIndex()) {
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				PdfPCell c1 = new PdfPCell(new Phrase(String.valueOf((int)cell.getNumericCellValue()),plainFont));
				c1.setHorizontalAlignment(Element.ALIGN_CENTER);
				table.addCell(c1);
			}
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				PdfPCell c1 = new PdfPCell(new Phrase(cell.getStringCellValue(),plainFont));
				c1.setHorizontalAlignment(Element.ALIGN_CENTER);
				table.addCell(c1);;
			}

		} else {
			while (cellNumber < cell.getColumnIndex()) {

				table.addCell(" ");
				cellNumber++;

			}
			if (cellNumber == cell.getColumnIndex()) {
				if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					PdfPCell c1 = new PdfPCell(new Phrase(String.valueOf((int)cell.getNumericCellValue()),plainFont));
					c1.setHorizontalAlignment(Element.ALIGN_CENTER);
					table.addCell(c1);
				}
				if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					PdfPCell c1 = new PdfPCell(new Phrase(cell.getStringCellValue(),plainFont));
					c1.setHorizontalAlignment(Element.ALIGN_CENTER);
					table.addCell(c1);;
				}

			}
			cellNumber = cell.getColumnIndex();
		}

		return cellNumber;

	}
	private static void addTitlePage(Document document,String i) throws DocumentException {
		Paragraph preface = new Paragraph();
		addEmptyLine(preface, 6);
		preface.add(new Paragraph("Title of the document", catFont));
		addEmptyLine(preface, 1);
		preface.add(new Paragraph("Sheet Name is : "+i+" generated by: " + "Ankit Kumar" + ", " + new Date(), smallBold));
		addEmptyLine(preface, 3);
		preface.add(new Paragraph("This document describes something which is very important ", smallBold));
		addEmptyLine(preface, 5);

		//preface.add(new Paragraph("This document is a preliminary version  ;-).", redFont));

		document.add(preface);
		document.newPage();
	}
	private static void addEmptyLine(Paragraph paragraph, int number) {
		for (int i = 0; i < number; i++) {
			paragraph.add(new Paragraph(" "));
		}
	}
}
