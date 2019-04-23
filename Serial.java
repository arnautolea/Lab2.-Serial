package mvn.Lab2;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Reader;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;

public class Serial {
	
	private static final String csvFile = "valuesBNM.csv";
	public static void main (String[] args) throws IOException, ParserConfigurationException, SAXException, TransformerException {
	//create workbook
	XSSFWorkbook wb = new XSSFWorkbook();
        XSSFCreationHelper createHelper = wb.getCreationHelper(); 
		try (Reader reader = Files.newBufferedReader(Paths.get(csvFile));
		            CSVReader csvReader = new CSVReaderBuilder(reader).build();
		        ) {
		        	// Reading All Records at once into a List<String[]>
		            List<String[]> records = csvReader.readAll();
			for (String[] record : records) {
            		System.out.print(record[0]);//output url from csv
			 //read xml from url           	
			 URL url = new URL(record[0]); 
			 URLConnection conn = url.openConnection();
			 //build document as new instance
			 DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			 DocumentBuilder builder = factory.newDocumentBuilder();
			 Document doc = builder.parse(conn.getInputStream());
			 //transform instance 
			 TransformerFactory transformerFactory= TransformerFactory.newInstance();
			 Transformer xform = transformerFactory.newTransformer();
			 xform.transform(new DOMSource(doc), new StreamResult(System.out)); //output xml in console 
			 //read document, normalize xml content
			 doc.getDocumentElement().normalize();
			 //create node list for root element
			 NodeList dList = doc.getElementsByTagName("ValCurs");
         			//start parse
		 		for (int d = 0; d < dList.getLength(); d++) {
		 		Node dNode = dList.item(d);
		 			if (dNode.getNodeType() == Node.ELEMENT_NODE) {
		 			Element dElement = (Element) dNode;
		 			//read attribute Date of root element, create sheets with similar name
		 			XSSFSheet sheet = wb.createSheet(dElement.getAttribute("Date")); 
		 			//Create first row as header
		 			Row row = sheet.createRow(0);
		 			row.createCell(0).setCellValue("ID");
		 			row.createCell(1).setCellValue("NumCode");
		 			row.createCell(2).setCellValue("CharCode");
		 			row.createCell(3).setCellValue("Nominal");
		 			row.createCell(4).setCellValue("Name");
		 			row.createCell(5).setCellValue("Value");
		 			//create node list for tag Valute 
		 			NodeList nList = dElement.getElementsByTagName("Valute");
		 				//get length of list
		 				for (int temp = 0; temp < nList.getLength(); temp++) {
		 				Node nNode = nList.item(temp);
		 					if (nNode.getNodeType() == Node.ELEMENT_NODE) {
		 					//create row for every element
		 					Row row1 = sheet.createRow(temp+1);
		 					Element nElement = (Element) nNode;
		 					//read elements by tag, start writing from second row (temp+1), first is header
		 					row1.createCell(0).setCellValue(createHelper
		 					.createRichTextString(nElement.getAttribute("ID")));
		 					row1.createCell(1).setCellValue(createHelper.createRichTextString(nElement
		 					.getElementsByTagName("NumCode").item(0).getTextContent()));
		 					row1.createCell(2).setCellValue(createHelper.createRichTextString(nElement
		 					.getElementsByTagName("CharCode").item(0).getTextContent()));
		 					row1.createCell(3).setCellValue(createHelper.createRichTextString(nElement
		 					.getElementsByTagName("Nominal").item(0).getTextContent()));
		 					row1.createCell(4).setCellValue(createHelper.createRichTextString(nElement
		 					.getElementsByTagName("Name").item(0).getTextContent()));
		 					row1.createCell(5).setCellValue(createHelper.createRichTextString(nElement
		 					.getElementsByTagName("Value").item(0).getTextContent()));
		 					}//if
		 				}//for
		 			}//if
		 		}//for
			}//for
		}//try
		//create output stream for writing values to file 
	try (OutputStream fileOut = new FileOutputStream("valcuuuuuuurs.xlsx")) {
         wb.write(fileOut);
         wb.close();
        System.out.println("\nFile changed");
     
    	} catch (FileNotFoundException e) {
			e.printStackTrace();
    	} catch (IOException e1) {
			e1.printStackTrace();
		}
	}//main
}//class
