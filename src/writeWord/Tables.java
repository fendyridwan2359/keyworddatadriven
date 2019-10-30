package writeWord;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Tables {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		
		XWPFDocument doc = new XWPFDocument();
		
		XWPFTable table = doc.createTable();
		
		XWPFParagraph paragraph;
		
		XWPFRun run;
		
		XWPFTableRow row0 = table.getRow(0);
		XWPFTableCell cell0 = row0.getCell(0);
		
		//cell0.setText("test");
		
		paragraph = doc.createParagraph();
		run = paragraph.createRun();
		run.setUnderline(UnderlinePatterns.WORDS);
		table = doc.createTable(2, 2);
		table.getCTTbl().getTblPr().setTblBorders(null);
		
		try {
			FileInputStream fis = new FileInputStream("C:/Users/fendy.ridwan/Pictures/ITS/attachment 1.jpg");
			 run.addPicture(fis, XWPFDocument.PICTURE_TYPE_JPEG, "image.jpg", Units.toEMU(100), Units.toEMU(100));			 
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		
		
		
		run.setText("Chicken pot pie");
		run.setColor("FF9900");
		
		try {
			FileInputStream fis = new FileInputStream("C:/Users/fendy.ridwan/Pictures/ITS/attachment 2.jpg");
			 run.addPicture(fis, XWPFDocument.PICTURE_TYPE_JPEG, "image.jpg", Units.toEMU(100), Units.toEMU(100));
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		run.setText("Chicken pot pie");
		run.setColor("FF9900");
		try{
			FileOutputStream output = new FileOutputStream("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\TestResult.docx");
			doc.write(output);
			output.close();
		}catch(Exception e){
			e. printStackTrace();
		}
		
		
	}

}
