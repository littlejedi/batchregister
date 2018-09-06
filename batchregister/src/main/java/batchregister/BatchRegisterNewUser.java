package batchregister;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import javax.ws.rs.core.MediaType;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jasypt.util.password.BasicPasswordEncryptor;

import com.google.common.base.Strings;
import com.liangzhi.commons.domain.EducationLevel;
import com.liangzhi.commons.domain.User;
import com.liangzhi.commons.domain.UserCredentials;
import com.liangzhi.commons.domain.UserGrade;
import com.liangzhi.commons.domain.UserRegistration;
import com.liangzhi.commons.domain.UserType;
import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.ClientResponse;
import com.sun.jersey.api.client.WebResource;
import com.sun.jersey.api.client.filter.HTTPBasicAuthFilter;

public class BatchRegisterNewUser {

    private static final String PASS = "coreapi!123";

    public static void main(String[] args) throws Exception {
        File myFile = new File("C://temp/大连八十中信息表.xlsx");
        //BufferedReader fis = new BufferedReader(new InputStreamReader(new FileInputStream(myFile), "UTF8"));
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        // Get iterator to all the rows in current sheet
        Iterator < Row > rowIterator = mySheet.iterator();
        
        rowIterator.next();
        
        //rowIterator.next();

        boolean skip = false;
        
        boolean printUserInfoOnly = false;
        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell province = row.getCell(0);
            province.setCellType(Cell.CELL_TYPE_STRING);
            Cell city = row.getCell(1);
            city.setCellType(Cell.CELL_TYPE_STRING);
            Cell district = row.getCell(2);
            district.setCellType(Cell.CELL_TYPE_STRING);
            Cell schoolName = row.getCell(3);
            schoolName.setCellType(Cell.CELL_TYPE_STRING);
            Cell name = row.getCell(4);
            name.setCellType(Cell.CELL_TYPE_STRING);
            Cell nationalId = row.getCell(5);
            nationalId.setCellType(Cell.CELL_TYPE_STRING);
            Cell phone = row.getCell(6);
            phone.setCellType(Cell.CELL_TYPE_STRING);
            Cell email = row.getCell(7);
            email.setCellType(Cell.CELL_TYPE_STRING);
            Cell password = row.getCell(8);
            password.setCellType(Cell.CELL_TYPE_STRING);
            String nameStr = name.getStringCellValue();
            String nationalIdStr = nationalId.getStringCellValue();
            String phoneStr = phone.getStringCellValue();
            String passwordStr = password.getStringCellValue();
            String provinceStr = province.getStringCellValue();
            String cityStr = city.getStringCellValue();
            String districtStr = district.getStringCellValue();
            String schoolNameStr = schoolName.getStringCellValue();
            /*if (!nationalIdStr.equals("ycjhstem050") && skip) {
                System.out.println("Skipping to ycjhstem050");
                continue;
            } else {
                skip = false;
            }*/
            if (Strings.isNullOrEmpty(nameStr) || Strings.isNullOrEmpty(nationalIdStr)) {
            	continue;
            }
            UserRegistration registration = new UserRegistration();
            final String username = row.getCell(9).getStringCellValue();
            registration.setUsername(username);
            final String basicPassword = "Stem123";
            //final String basicPassword = phoneStr;
            //System.out.println(basicPassword);
            registration.setBasicPassword(basicPassword);
            registration.setType(UserType.REGULAR);
            registration.setRealName(nameStr);
            registration.setPhoneNumber(phoneStr);
            registration.setNationalId(nationalIdStr);
            registration.setEmail(email.getStringCellValue());
            registration.setGrade(UserGrade.MIDDLE_SCHOOL);
            registration.setCurrentSchoolLevel(EducationLevel.MIDDLE_SCHOOL);
            registration.setProvince(province.getStringCellValue());
            registration.setCity(city.getStringCellValue());
            registration.setDistrict(district.getStringCellValue());
            registration.setSchoolName(schoolName.getStringCellValue());
            Client client = Client.create();
            client.addFilter(new HTTPBasicAuthFilter("root", PASS));
            client.setConnectTimeout(60000);
            client.setReadTimeout(60000);
            WebResource webResource = null;
            ClientResponse response = null;
            System.out.println(registration.toString());
            if (printUserInfoOnly) {
            	continue;
            }
            // Get user
            User user;
            try {
              doRegister(nameStr, nationalIdStr, registration, username, basicPassword, phoneStr, provinceStr, cityStr, districtStr, schoolNameStr, client);
            } catch (Exception e) {
            	// Retry forever
            	while (true) {
            		Thread.sleep(2000);
            		doRegister(nameStr, nationalIdStr, registration, username, basicPassword, phoneStr, provinceStr, cityStr, districtStr, schoolNameStr, client);
            	}
            }
            // Verify login
            webResource = client.resource("http://121.40.132.224:8080/users/login");
            UserCredentials credz = new UserCredentials(username, basicPassword);
            response = webResource.type(MediaType.APPLICATION_JSON)
                    .post(ClientResponse.class, credz);
            if (response.getStatus() == 200) {
                System.out.println("Login successful for user=" + registration.toString());
            } else {
                throw new Exception("Login failed for user=" + registration.toString());
            }
         }
    }

	private static void doRegister(String nameStr, String nationalIdStr,
			UserRegistration registration, final String username,
			final String basicPassword, final String phone, final String province, final String city, final String district, final String schoolName, Client client) throws Exception {
		WebResource webResource;
		ClientResponse response;
		User user;
		webResource = client
		        .resource("http://121.40.132.224:8080/users").path("findByUsername").queryParam("username", username);
		response = webResource.accept(MediaType.APPLICATION_JSON).get(ClientResponse.class);
		if (response.getStatus() == 404) {
		    webResource = client.resource("http://121.40.132.224:8080/users/register");
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
		    user.setProvince(province);
		    user.setCity(city);
		    user.setDistrict(district);
		    user.setSchoolName(schoolName);
		    user.setPhoneNumber(phone);
		    user.setUsername(user.getUsername().trim());
		    webResource = client.resource("http://121.40.132.224:8080/users").path(user.getId().toString());
		    response = webResource.type(MediaType.APPLICATION_JSON).put(ClientResponse.class, user);
		    if (response.getStatus() != 200) {
		        throw new Exception("response is not 200");
		    }
		    System.out.println("Updating exising user " + registration.toString());
		}
	}


}
