package com.office.pdfConverter.Service;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
public class pptService {

	public static String pptToPdf() throws IOException, InvalidFormatException, DocumentException {
		String sourcePath = "/home/jinx/final.pptx";
		String destinationPath = "/home/jinx/fileName.pdf";
		FileInputStream inputStream = new FileInputStream(sourcePath);
		double zoom = 2;
	    java.awt.Dimension pgsize = null;
	    Document pdfDocument = new Document();
	    PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument, new FileOutputStream(destinationPath));
	    PdfPTable table = new PdfPTable(1);
	    pdfWriter.open();
	    pdfDocument.open();
	    Image slideImage = null;
	    BufferedImage img = null;
        AffineTransform at = new AffineTransform();
        at.setToScale(zoom, zoom);
		if (sourcePath.endsWith(".pptx")) {
			
	        XMLSlideShow pptx = new XMLSlideShow(inputStream);
	        pgsize = pptx.getPageSize();
	        List<XSLFSlide> slide = pptx.getSlides();
	        for (int j = 0; j < slide.size(); j++) {
                img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setTransform(at);

                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                slide.get(j).draw(graphics);
                graphics.getPaint();
                slideImage = Image.getInstance(img, null);
                table.addCell(new PdfPCell(slideImage, true));
                }
	        pdfDocument.setPageSize(new Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
	        pdfWriter.open();
	        pdfDocument.open();
	        
	        pdfDocument.add(table);
	        pdfDocument.close();
	        pdfWriter.close();
			
		}
		else if(sourcePath.endsWith(".ppt")){
			HSLFSlideShow ppt= new HSLFSlideShow(inputStream);
			pgsize = ppt.getPageSize();
			List<HSLFSlide> slide = ppt.getSlides();
			for (int j = 0; j < slide.size(); j++) {
                img = new BufferedImage((int) Math.ceil(pgsize.width * zoom), (int) Math.ceil(pgsize.height * zoom), BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setTransform(at);

                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
                slide.get(j).draw(graphics);
                graphics.getPaint();
                slideImage = Image.getInstance(img, null);
                table.addCell(new PdfPCell(slideImage, true));
                }
	        pdfDocument.setPageSize(new Rectangle((float) pgsize.getWidth(), (float) pgsize.getHeight()));
	        pdfWriter.open();
	        pdfDocument.open();
	        
	        pdfDocument.add(table);
	        pdfDocument.close();
	        pdfWriter.close();
		}
		
	    
		return "pdf created successfully at " + destinationPath;
	}

}
