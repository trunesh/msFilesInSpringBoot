package com.msfiles.msfilespr.service;

import java.io.IOException;
import java.net.URL;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelGeneratorService {
	
	public Workbook createExcelFile() throws IOException {
		Workbook wb=new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("First cell value");
		
		CellStyle style = wb.createCellStyle();
		style.setBorderBottom(BorderStyle.THICK);
		
		Font f=wb.createFont();
		f.setColor(IndexedColors.GREEN.index);
		f.setBold(true);
		
		style.setFont(f);
		embeded("image1.jpg",wb,sheet);
		return wb;
	}

	private void embeded(String imagePath,Workbook wb,Sheet sheet) throws IOException {
		URL fr=ExcelGeneratorService.class.getResource("/images/" + imagePath);
		byte[] bytes = IOUtils.toByteArray(fr.openStream());
		int pictureIndex  = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
		
		CreationHelper  helper= wb.getCreationHelper();
		
		Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();
		
		anchor.setCol1(1);
		anchor.setCol1(2);
		anchor.setCol1(3);
		
		Picture pict=drawingPatriarch.createPicture(anchor, pictureIndex);
		pict.resize();
	}
}
