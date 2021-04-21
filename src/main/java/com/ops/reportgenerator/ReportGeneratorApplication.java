package com.ops.reportgenerator;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.chrono.JapaneseChronology;
import java.time.chrono.JapaneseDate;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ReportGeneratorApplication {
	public static String output = "rest-with-spring.docx";
	public static void test() throws Exception{
//		XWPFDocument document = new XWPFDocument();
//        XWPFParagraph title = document.createParagraph();
//        title.setAlignment(ParagraphAlignment.CENTER);
//        XWPFRun titleRun = title.createRun();
//        titleRun.setText("Build Your REST API with Spring");
//        titleRun.setColor("009933");
//        titleRun.setBold(true);
//        titleRun.setFontFamily("Courier");
//        titleRun.setFontSize(20);
//        FileOutputStream out = new FileOutputStream(output);
//        document.write(out);
//        out.close();
//        document.close();
//		   FileInputStream fis = new FileInputStream("file/test.docx");
//		   XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
//		   XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
//		   System.out.println(extractor.getText());
		System.out.println("hdshgfsgfh");
		
		try {
			FileInputStream fis = new FileInputStream("file/test.docx");
			XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
			Iterator bodyElementIterator = xdoc.getBodyElementsIterator();
			String dayString = "日";
			LocalDate date = LocalDate.now();
			
			while (bodyElementIterator.hasNext()) {
				IBodyElement element = (IBodyElement) bodyElementIterator.next();

				if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
					List<XWPFTable> tableList = element.getBody().getTables();
					tableList.get(0).getRow(1).getCell(0).setText("sdfsdg");
					
					tableList.get(0).getRow(1).getCell(0).removeParagraph(0);
					XWPFParagraph addParagraph = tableList.get(0).getRow(1).getCell(0).addParagraph();
//					addParagraph.setAlignment(ParagraphAlignment.RIGHT);  
				       XWPFRun run = addParagraph.createRun();
				       run.setText(date.getDayOfMonth()+dayString);
				       DateTimeFormatter japaneseEraDtf = DateTimeFormatter.ofPattern("GGGGy年M月d日")
				               .withChronology(JapaneseChronology.INSTANCE)
				               .withLocale(Locale.JAPAN);

				       
				       
				       LocalDate gregorianDate = LocalDate.parse(date.toString());
				       JapaneseDate japaneseDate = JapaneseDate.from(gregorianDate);
				       String hidzuke =japaneseDate.format(japaneseEraDtf);
				       System.out.println(hidzuke+"nihon date");
				       
						tableList.get(0).getRow(1).getCell(2).removeParagraph(0);
						 addParagraph = tableList.get(0).getRow(1).getCell(2).addParagraph();
						 addParagraph.setAlignment(ParagraphAlignment.RIGHT);
					        run = addParagraph.createRun();
					       run.setText(hidzuke);
					System.out.println(tableList.get(0).getRow(1).getCell(0).getText()+" fromdd"+japaneseEraDtf);
					
					for (XWPFTable table : tableList) {
						System.out.println("Total Number of Rows of Table:" + table.getNumberOfRows());
						for (int i = 0; i < table.getRows().size(); i++) {

							for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
								System.out.println(table.getRow(i).getCell(j).getText());
							}
						}
					}
				}
			}
			FileOutputStream fos = new FileOutputStream("file/newText.docx");
			xdoc.write(fos);
			fos.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}
	
	public static void main(String[] args) {
		
		try {
			test();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		SpringApplication.run(ReportGeneratorApplication.class, args);
	}

}
