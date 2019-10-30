package AutomationFramework;
/**
 * @author fendy.ridwan
 * 
 *
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.testng.annotations.Test;

import utility.CustomXWPFDocument;
public class ImageAttachmentInDocument {
	
	public String[] arrayFileName = new String[700];
	public String[] arrayResultStatus = new String[700];
	
	@Test
	public void createResult(String getdata,int counter, int lastrownumber, String getResultStatus) throws Exception
	{
		//Tampung nama gambar sementara dalam bentuk array
		arrayFileName[counter]=getdata;
		arrayResultStatus[counter]=getResultStatus;
		//proses masukan ke dokumen jika sudah di akhir row pada excel saja
		if(counter==lastrownumber){
		XWPFDocument doc = new XWPFDocument();
    	//Make Document landscape mode
    	changeOrientation(doc, "landscape");
    	
    	int nRows = 0;
    	for(String rowValue : arrayFileName)
		{
		    if(rowValue!=null)
		    {
		    	nRows++;
		    }
		}
 	    int nCols = 4;
 	    
 	    XWPFTable table = doc.createTable(nRows+1, nCols);
	    // Set the table style. If the style is not defined, the table style
	    // will become "Normal".
	    CTTblPr tblPr = table.getCTTbl().getTblPr();
	    CTString styleStr = tblPr.addNewTblStyle();
	    styleStr.setVal("StyledTable");
	    // Get a list of the rows in the table
	    
		List<XWPFTableRow> rows = table.getRows();
		System.out.println(rows.get(0).getTableICells());
	    
		int rowCt = 0;
	    //HEADER
		if (rowCt == 0) {
		List<XWPFTableCell> cells = rows.get(0).getTableCells();
		// Header No Script
        XWPFParagraph cell0 = cells.get(0).getParagraphs().get(0);
        XWPFRun rh0 = cell0.createRun();
		rh0.setText("No Script");
		rh0.setBold(true);
        cell0.setAlignment(ParagraphAlignment.CENTER);
        
		// Header Menu Navigation
        XWPFParagraph cell1 = cells.get(1).getParagraphs().get(0);
        XWPFRun rh1 = cell1.createRun();
        rh1.setText("Menu Navigation");
        rh1.setBold(true);
        cell1.setAlignment(ParagraphAlignment.CENTER);
         
		// Header Capture Result
        XWPFParagraph cell2 = cells.get(2).getParagraphs().get(0);
        XWPFRun rh2 = cell2.createRun();
		rh2.setText("Capture Result");
		rh2.setBold(true);
        cell2.setAlignment(ParagraphAlignment.CENTER);
        
		// Header Status
        XWPFParagraph cell3 = cells.get(3).getParagraphs().get(0);
        XWPFRun rh3 = cell3.createRun();
		rh3.setText("Status");
		rh3.setBold(true);
		cell3.setAlignment(ParagraphAlignment.CENTER);
		}
		rowCt = 1;
		int rowResult = 0;
		for(String rowValue : arrayFileName)
		{
			
			if(rowValue!=null)
		    {
				//File img = new File("D:/WORK/Selenium Testing/Project/AutomationTest/AutomationTest/Test Files/attachment/"+rowValue+".jpg");
				File img = new File("/Users/fendyridwan/Documents/Project/AutomationTest/Test Files/attachment/"+rowValue+".jpg");
				String imgFile = img.getName();
		    	int imgFormat = getImageFormat(imgFile);
				
				//CONTAIN
				if (rowCt >= 0) {
				List<XWPFTableCell> cells = rows.get(rowCt).getTableCells();
				// Create containt No Script
	            XWPFParagraph containt0 = cells.get(0).getParagraphs().get(0);
	            XWPFRun rhContaint0 = containt0.createRun();
				rhContaint0.setText(rowValue);
				rhContaint0.setBold(true);
                containt0.setAlignment(ParagraphAlignment.CENTER);
                
				// Create containt Menu Navigation
	            XWPFParagraph containt1 = cells.get(1).getParagraphs().get(0);
	            XWPFRun rhContaint1 = containt1.createRun();
                rhContaint1.setText("-");
                rhContaint1.setBold(true);
                containt1.setAlignment(ParagraphAlignment.CENTER);
                
				// Create containt Capture Result
	            XWPFParagraph containt2 = cells.get(2).getParagraphs().get(0);
	            XWPFRun rhContaint2 = containt2.createRun();
	            rhContaint2.addPicture(new FileInputStream(img), imgFormat, imgFile, Units.toEMU(350), Units.toEMU(210));
                containt2.setAlignment(ParagraphAlignment.CENTER);
	            
				// Create containt Status
	            XWPFParagraph containt3 = cells.get(3).getParagraphs().get(0);
	            XWPFRun rhContaint3 = containt3.createRun();
	            String TestResult = arrayResultStatus[rowResult];
				rhContaint3.setText(TestResult);
				rhContaint3.setBold(true);
				containt3.setAlignment(ParagraphAlignment.CENTER);
				}

		    	rowCt++;
		    }
			rowResult++;
		}
		FileOutputStream out = new FileOutputStream("D:\\AutomationTest\\Test Files\\TestResult.docx");
	    doc.write(out);
	    out.close();
	    System.out.println("Process Completed Successfully");
		}
	}
	
	private static void changeOrientation(XWPFDocument document, String orientation){
	    CTDocument1 doc = document.getDocument();
	    CTBody body = doc.getBody();
	    CTSectPr section = body.addNewSectPr();
	    XWPFParagraph para = document.createParagraph();
	    CTP ctp = para.getCTP();
	    CTPPr br = ctp.addNewPPr();
	    br.setSectPr(section);
	    CTPageSz pageSize = section.isSetPgSz() ? section.getPgSz() : section.addNewPgSz();
	    if(orientation.equals("landscape")){
	        pageSize.setOrient(STPageOrientation.LANDSCAPE);
	        pageSize.setW(BigInteger.valueOf(842 * 20));
	        pageSize.setH(BigInteger.valueOf(595 * 20));
	    }
	    else{
	        pageSize.setOrient(STPageOrientation.PORTRAIT);
	        pageSize.setH(BigInteger.valueOf(842 * 20));
	        pageSize.setW(BigInteger.valueOf(595 * 20));
	    }
	}

	private static int getImageFormat(String imgFile) {
		int format;
		if (imgFile.endsWith(".emf"))
			format = XWPFDocument.PICTURE_TYPE_EMF;
		else if (imgFile.endsWith(".wmf"))
			format = XWPFDocument.PICTURE_TYPE_WMF;
		else if (imgFile.endsWith(".pict"))
			format = XWPFDocument.PICTURE_TYPE_PICT;
		else if (imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg"))
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		else if (imgFile.endsWith(".png"))
			format = XWPFDocument.PICTURE_TYPE_PNG;
		else if (imgFile.endsWith(".dib"))
			format = XWPFDocument.PICTURE_TYPE_DIB;
		else if (imgFile.endsWith(".gif"))
			format = XWPFDocument.PICTURE_TYPE_GIF;
		else if (imgFile.endsWith(".tiff"))
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		else if (imgFile.endsWith(".eps"))
			format = XWPFDocument.PICTURE_TYPE_EPS;
		else if (imgFile.endsWith(".bmp"))
			format = XWPFDocument.PICTURE_TYPE_BMP;
		else if (imgFile.endsWith(".wpg"))
			format = XWPFDocument.PICTURE_TYPE_WPG;
		else {
			return 0;
		}
		return format;
	}

	
	
}