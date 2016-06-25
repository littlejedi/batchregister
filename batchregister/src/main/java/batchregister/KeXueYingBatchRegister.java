package batchregister;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import javax.ws.rs.core.MediaType;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jasypt.util.password.BasicPasswordEncryptor;

import com.liangzhi.commons.domain.User;
import com.liangzhi.commons.domain.UserCredentials;
import com.liangzhi.commons.domain.UserRegistration;
import com.liangzhi.commons.domain.UserType;
import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.ClientResponse;
import com.sun.jersey.api.client.WebResource;
import com.sun.jersey.api.client.filter.HTTPBasicAuthFilter;

// 2016 科学营注册 
public class KeXueYingBatchRegister {

    private static final String PASS = "coreapi!123";

    public static void main(String[] args) throws Exception {
        File myFile = new File("C://littlejedi/liangzhi/kexueying2.xlsx");
        //BufferedReader fis = new BufferedReader(new InputStreamReader(new FileInputStream(myFile), "UTF8"));
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        // Get iterator to all the rows in current sheet
        Iterator < Row > rowIterator = mySheet.iterator();
        
        rowIterator.next();

        boolean skip = false;
        
        boolean printUserInfoOnly = true;
        
        String previousProvince = null;
        int count = 1;
        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell province = row.getCell(0);
            province.setCellType(Cell.CELL_TYPE_STRING);
            Cell name = row.getCell(2);
            name.setCellType(Cell.CELL_TYPE_STRING);
            Cell birthday = row.getCell(5);
            birthday.setCellType(Cell.CELL_TYPE_STRING);
            Cell nationalId = row.getCell(7);
            nationalId.setCellType(Cell.CELL_TYPE_STRING);
            Cell usernameCell = row.createCell(8);
            usernameCell.setCellType(Cell.CELL_TYPE_STRING);
            Cell passwordCell = row.createCell(9);
            passwordCell.setCellType(Cell.CELL_TYPE_STRING);
            String provinceStr = province.getStringCellValue();
            if (previousProvince == null || !provinceStr.equals(previousProvince)) {
            	count = 1;
            }
            String number = String.format("%03d", count);
            String nameStr = name.getStringCellValue();
            String birthdayStr = birthday.getStringCellValue().replace("-", "");
            String nationalIdStr = nationalId.getStringCellValue().replaceAll("[^A-Za-z0-9]", "");;
            UserRegistration registration = new UserRegistration();
            final String username = "2016" + provinceStr + number;
            registration.setUsername(username);
            final String basicPassword = "stem" + birthdayStr;
            //System.out.println(basicPassword);
            registration.setBasicPassword(basicPassword);
            registration.setType(UserType.REGULAR);
            registration.setRealName(nameStr);
            registration.setPhoneNumber("12345");
            registration.setNationalId(nationalIdStr);
            usernameCell.setCellValue(username);
            passwordCell.setCellValue(basicPassword);
            Client client = Client.create();
            client.addFilter(new HTTPBasicAuthFilter("root", PASS));
            client.setConnectTimeout(60000);
            client.setReadTimeout(60000);
            WebResource webResource = null;
            ClientResponse response = null;
            System.out.println(registration.toString());
            previousProvince = provinceStr;
            count++;
            if (printUserInfoOnly) {
            	continue;
            }
            // Get user
            User user;
            try {
              doRegister(nameStr, nationalIdStr, registration, username, basicPassword, client);
            } catch (Exception e) {
            	// Retry forever
            	while (true) {
            		Thread.sleep(2000);
            		doRegister(nameStr, nationalIdStr, registration, username, basicPassword, client);
            	}
            }
            // Verify login
            webResource = client.resource("http://www.stemcloud.cn:8080/users/login");
            UserCredentials credz = new UserCredentials(username, basicPassword);
            response = webResource.type(MediaType.APPLICATION_JSON)
                    .post(ClientResponse.class, credz);
            if (response.getStatus() == 200) {
                System.out.println("Login successful for user=" + registration.toString());
            } else {
                throw new Exception("Login failed for user=" + registration.toString());
            }
         }
        
        FileOutputStream fileOut = new FileOutputStream("C://littlejedi/liangzhi/kexueying2.xlsx");  
        myWorkBook.write(fileOut);  
        fileOut.close();  
    }

	private static void doRegister(String nameStr, String nationalIdStr,
			UserRegistration registration, final String username,
			final String basicPassword, Client client) throws Exception {
		WebResource webResource;
		ClientResponse response;
		User user;
		webResource = client
		        .resource("http://www.stemcloud.cn:8080/users").path("findByUsername").queryParam("username", username);
		response = webResource.accept(MediaType.APPLICATION_JSON).get(ClientResponse.class);
		if (response.getStatus() == 404) {
		    webResource = client.resource("http://www.stemcloud.cn:8080/users/register");
		    response = webResource.type(MediaType.APPLICATION_JSON + ";charset=utf-8")
		            .post(ClientResponse.class, registration);
		    if (response.getStatus() != 200) {
		        throw new Exception("response is not 200");
		    }
		    System.out.println("Registering new user " + registration.toString());
		} else if (response.getStatus() == 200) {
		    user = response.getEntity(User.class);
		    BasicPasswordEncryptor passwordEncryptor = new BasicPasswordEncryptor();
		    String encryptedPassword = passwordEncryptor.encryptPassword(basicPassword);
		    user.setBasicPassword(encryptedPassword);
		    user.setType(UserType.REGULAR);
		    user.setRealName(nameStr);
		    user.setPhoneNumber(nationalIdStr);
		    user.setNationalId(nationalIdStr);
		    user.setUsername(user.getUsername().trim());
		    webResource = client.resource("http://www.stemcloud.cn:8080/users").path(user.getId().toString());
		    response = webResource.type(MediaType.APPLICATION_JSON)
		            .put(ClientResponse.class, user);
		    if (response.getStatus() != 200) {
		        throw new Exception("response is not 200");
		    }
		    System.out.println("Updating exising user " + registration.toString());
		}
	}


}
