package com.Lettergeneration;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;


public class DocumentToPdfConversation 
{
	 public static void ConvertToPDF(XWPFDocument document, String pdfPath) {
	        try {
	            PdfOptions options = PdfOptions.create();
	            OutputStream out = new FileOutputStream(new File(pdfPath,"DonePdf.pdf"));
	            PdfConverter.getInstance().convert(document, out, options);
	            System.out.println("Task completed");
	        } catch (Exception ex) {
	            System.out.println(ex.getMessage());
	            ex.printStackTrace();
	        }
	    }
	 public static void main(String args[]) throws InvalidFormatException, IOException
	 {
		 XWPFDocument doc = new XWPFDocument(OPCPackage.open("D:\\A\\LGI\\output\\Document\\71070031170160011488.docx"));			
		 DocumentToPdfConversation.ConvertToPDF(doc, "D:\\A\\LGI\\More Project BRD\\");
	 }
}
