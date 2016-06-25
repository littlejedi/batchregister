package batchregister;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import javax.ws.rs.core.MediaType;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.liangzhi.commons.domain.StemTestSysRegisterRequest;
import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.ClientResponse;
import com.sun.jersey.api.client.WebResource;
import com.sun.jersey.api.client.filter.HTTPBasicAuthFilter;

public class BatchExamsClear {
	
	private static final String PASS = "coreapi!123";
	
	public static void main(String[] args) throws Exception {
		File myFile = new File("C://littlejedi/liangzhi/2016shengwu.xlsx");
		//BufferedReader fis = new BufferedReader(new InputStreamReader(new FileInputStream(myFile), "UTF8"));
		FileInputStream fis = new FileInputStream(myFile);

		// Finds the workbook instance for XLSX file
		XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

		// Return first sheet from the XLSX workbook
		XSSFSheet mySheet = myWorkBook.getSheetAt(1);

		// Get iterator to all the rows in current sheet
		Iterator < Row > rowIterator = mySheet.iterator();
		
		rowIterator.next();

		// Traversing over each row of XLSX file
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Cell name = row.getCell(2);
			Cell nationalId = row.getCell(5);
			name.setCellType(Cell.CELL_TYPE_STRING);
			nationalId.setCellType(Cell.CELL_TYPE_STRING);
			String nameString = name.getStringCellValue();
			String nationalIdString = nationalId.getStringCellValue();
			System.out.println("Resetting exam times for name: " + nameString + " , nationalId: " + nationalIdString);
			Client client = Client.create();
	        client.addFilter(new HTTPBasicAuthFilter("root", PASS));
	        client.setConnectTimeout(60000);
	        client.setReadTimeout(60000);
	        WebResource webResource = null;
            ClientResponse response = null;
			boolean update = true;
			if (update) {
	            if (nationalIdString != null) {
	                try {
		            	webResource = client.resource("http://www.stemcloud.cn:8080/testsys/resetexamtimes/" + nationalIdString);
		            	response = webResource.get(ClientResponse.class);
	                } catch (Exception e) {
	                	System.out.println("Error updating for" + nationalIdString);
	                }
	            }
			}
		}
	}


}
