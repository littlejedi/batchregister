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

public class BatchRegister {

	public static void main(String[] args) throws Exception {
		File myFile = new File("C://java/shenbao4.xlsx");
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
			Cell phone = row.getCell(8);
			phone.setCellType(Cell.CELL_TYPE_STRING);
			Cell email = row.getCell(9);
			String nameStr = name.getStringCellValue();
			String nationalIdStr = nationalId.getStringCellValue();
			String phoneStr = null;
			if (phone.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				//throw new Exception("shouldn't reach here");
				double d = phone.getNumericCellValue();
				//phoneStr = String.valueOf(new BigDecimal(phone.getNumericCellValue()).setScale(0, RoundingMode.HALF_UP).intValue());
			}
			else if (phone.getCellType() == Cell.CELL_TYPE_STRING) {
				phoneStr = phone.getStringCellValue();
			} else {
				throw new Exception("shouldn't reach here");
			}
			String emailStr = email.getStringCellValue();
			StemTestSysRegisterRequest req = new StemTestSysRegisterRequest();
			req.setCardID(nationalIdStr);
			req.setName(nameStr);
			req.setEmail(emailStr);
			req.setPhone(phoneStr);
			Client client = Client.create();
			WebResource webResource = client.resource("http://hostname/api/users/stemtestsysregister");
			ClientResponse response = webResource.type(MediaType.APPLICATION_JSON + ";charset=utf-8")
					.post(ClientResponse.class, req);
			if (response.getStatus() != 200) {
				throw new Exception("response is not 200");
			}
			System.out.println(req.toString());
		}
	}

}
