package com.office.pdfConverter.Service;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;


public class docxService {

	public static String docxtoPdf() throws IOException, DocumentException {
		
		String sourcePath = "/home/jinx/sample6.docx";
		String destinationPath = "/home/jinx/fileName2.pdf";
		InputStream doc = new FileInputStream(new File(sourcePath));
		POIFSFileSystem fs = null;
		if (sourcePath.endsWith(".docx")) {
		try {
            
            XWPFDocument document = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(destinationPath));
            PdfConverter.getInstance().convert(document, out, options);
            out.close();
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
		}
		else if(sourcePath.endsWith(".doc")){
			fs = new POIFSFileSystem(new FileInputStream(sourcePath));
			HWPFDocument docs = new HWPFDocument(fs);  
            WordExtractor we = new WordExtractor(docs);
            OutputStream file = new FileOutputStream(new File(destinationPath)); 
            Document document = new Document();
            PdfWriter writer = PdfWriter.getInstance(document, file);
            Range range = docs.getRange();
            document.open();  
            writer.setPageEmpty(true);  
            document.newPage();  
            writer.setPageEmpty(true);  
            String[] paragraphs = we.getParagraphText();  
            for (int i = 0; i < paragraphs.length; i++) {  

                org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);
                paragraphs[i] = paragraphs[i].replaceAll("\\cM?\r?\n", "");  
                //System.out.println("Length:" + paragraphs[i].length());  
                //System.out.println("Paragraph" + i + ": " + paragraphs[i].toString());  

            // add the paragraph to the document  
            document.add(new Paragraph(paragraphs[i]));  
            
            }
            document.close();
            
		}
		return "updated";
	}  
      
}