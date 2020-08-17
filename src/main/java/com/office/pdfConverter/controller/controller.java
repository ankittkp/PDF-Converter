package com.office.pdfConverter.controller;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import com.itextpdf.text.DocumentException;
import com.office.pdfConverter.Service.excelService;
import com.office.pdfConverter.Service.pptService;

@RestController
public class controller {
	
	
	@PostMapping("/pptToPdf")
	public String toPdf() throws InvalidFormatException, IOException, DocumentException {
		String response = pptService.pptToPdf();
		return response;
	}
	@PostMapping("/excelToPdf")
	public String excelToPdf() throws IOException, DocumentException{
		String response = excelService.exceltoPdf();
		return response;
	}
}
