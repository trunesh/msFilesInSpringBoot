package com.msfiles.msfilespr.web;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import com.msfiles.msfilespr.service.ExcelGeneratorService;


@RestController
@RequestMapping("/e")
public class ExcelController {

	@Autowired
	ExcelGeneratorService excelGeneratorService;
	
	
	@GetMapping
	public ResponseEntity<StreamingResponseBody> excel() throws IOException {
		Workbook wb=excelGeneratorService.createExcelFile();
		return ResponseEntity
				.ok()
				.contentType(MediaType.APPLICATION_OCTET_STREAM)
				.header(HttpHeaders.CONTENT_DISPOSITION,"inline;fileName=\"excel.xslx\"")
				.body(wb::write);
	}
}
