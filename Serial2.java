package mvn.Lab2;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.thoughtworks.xstream.XStream;

public class Serial2 {
	
	private static final String csvFile = "valuesBNM2.csv";
	public static void main (String[] args) throws IOException, ParserConfigurationException, SAXException, TransformerException {
		
		XStream xstream = new XStream();
        xstream.processAnnotations(ValCurs.class);
        xstream.processAnnotations(Valute.class);
        xstream.allowTypesByWildcard(new String[] {"mvn.Lab2"});
       
		//create workbook
        XSSFWorkbook wb = new XSSFWorkbook();
        try (
	        	Reader reader = Files.newBufferedReader(Paths.get(csvFile));
	            CSVReader csvReader = new CSVReaderBuilder(reader).build();
	        ) {
	        	// Reading All Records at once into a List<String[]>
	            List<String[]> records = csvReader.readAll();
	            for (String[] record : records) {
	            	System.out.println(record[0]);	
	            		// Reading xml from url
		            	URL url = new URL("https://bnm.md/en/official_exchange_rates?get_xml=1&date=" + record[0]);
		            	InputStream input = url.openStream();
		            	String xml = IOUtils.toString(input, "utf-8");
		                
		            	ValCurs valCurs = (ValCurs) xstream.fromXML(xml);
		            	XSSFSheet sheet = wb.createSheet(valCurs.getDate());

					//Create first row as header
		 			Row row = sheet.createRow(0);
		 			row.createCell(0).setCellValue("NumCode");
		 			row.createCell(1).setCellValue("CharCode");
		 			row.createCell(2).setCellValue("Nominal");
		 			row.createCell(3).setCellValue("Name");
		 			row.createCell(4).setCellValue("Value");
		 			row.createCell(5).setCellValue("ID");
 		 			
		 				int nextRow = 1;
		 				for (Valute currentVal : valCurs.getValutes()) {
		 					System.out.println(currentVal);
		 					Row row2 = sheet.createRow(nextRow++);
		 					row2.createCell(0).setCellValue(currentVal.getNumCode());
		 					row2.createCell(1).setCellValue(currentVal.getСharCode());
		 					row2.createCell(2).setCellValue(currentVal.getNominal());
		 					row2.createCell(3).setCellValue(currentVal.getName());
		 					row2.createCell(4).setCellValue(currentVal.getValue());
		 					row2.createCell(5).setCellValue(currentVal.getId());
		 				}//for
		 			input.close();	
        		  }//for
        }//try
        
		//create output stream for writing values to workbook  
	try (OutputStream fileOut = new FileOutputStream("valcurs2.xlsx")) {
         wb.write(fileOut);
         wb.close();
        System.out.println("\nFile changed");
     
    	} catch (FileNotFoundException e1) {
			e1.printStackTrace();
    	} catch (IOException e2) {
			e2.printStackTrace();
		}
	}//main
}//class
