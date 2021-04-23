package com.ops.reportgenerator.controller;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.chrono.JapaneseChronology;
import java.time.chrono.JapaneseDate;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.http.MediaType;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ops.reportgenerator.entitites.GetDoc;

@RestController
@RequestMapping("/changeReport")
public class FileChanger {
	@CrossOrigin(origins = "http://localhost:3000")
	@PostMapping(value="/single_download",consumes = { "multipart/form-data" })
	public static void test( @ModelAttribute GetDoc body) throws Exception{

//		   System.out.println(extractor.getText());
		System.out.println("hdshgfsgfh"+body.getFile());
//		body.getFile().transferTo(tempFile);
//		InputStream stream = new FileInputStream(tempFile);
		InputStream inputStream =  new BufferedInputStream(body.getFile().getInputStream());
		System.out.println(body.getName()+"input stream"+inputStream);
		try {
			FileInputStream fis = new FileInputStream("file/test.docx");
			XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(inputStream));
			Iterator bodyElementIterator = xdoc.getBodyElementsIterator();
			String dayString = "日";
			String name ="オリップsd";
			LocalDate date = LocalDate.now();
			
			while (bodyElementIterator.hasNext()) {
				IBodyElement element = (IBodyElement) bodyElementIterator.next();

				if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
					List<XWPFTable> tableList = element.getBody().getTables();
					tableList.get(0).getRow(1).getCell(0).setText("sdfsdg");
					
					tableList.get(0).getRow(1).getCell(0).removeParagraph(0);
					XWPFParagraph addParagraph = tableList.get(0).getRow(1).getCell(0).addParagraph();
					addParagraph.setStyle("LO-normal");
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
						 addParagraph.setStyle("LO-normal");
						 addParagraph.setAlignment(ParagraphAlignment.RIGHT);
					        run = addParagraph.createRun();
					       run.setText(hidzuke);
					       
							tableList.get(0).getRow(0).getCell(1).removeParagraph(0);
							 addParagraph = tableList.get(0).getRow(0).getCell(1).addParagraph();
							 addParagraph.setStyle("LO-normal");
							 addParagraph.setAlignment(ParagraphAlignment.CENTER);
						        run = addParagraph.createRun();
						       run.setText("＜"+name+"の本日の業務ご報告＞");
						    
								tableList.get(0).getRow(3).getCell(2).removeParagraph(0);
								 addParagraph = tableList.get(0).getRow(3).getCell(2).addParagraph();
							
								 addParagraph.setStyle("LO-normal");
								 addParagraph.setAlignment(ParagraphAlignment.RIGHT);
							        run = addParagraph.createRun();
							       run.setText(name);
					System.out.println(tableList.get(0).getRow(1).getCell(0).getText()+" fromdd"+date);
					
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
			FileOutputStream fos = new FileOutputStream("/home/olip/Documents/daily/"+date+"社長報告「日報」オリップ.docx");
			xdoc.write(fos);
			fos.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}
	
}
