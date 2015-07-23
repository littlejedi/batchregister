package batchregister;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import javax.ws.rs.core.MediaType;

import junit.framework.Assert;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.liangzhi.commons.domain.User;
import com.liangzhi.commons.domain.UserRegistration;
import com.liangzhi.commons.domain.UserType;
import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.ClientResponse;
import com.sun.jersey.api.client.WebResource;

public class AnotherSimpleBatchRegister {

    public static void main(String[] args) throws Exception {
        File myFile = new File("E://liangzhi/201507223.xlsx");
        //BufferedReader fis = new BufferedReader(new InputStreamReader(new FileInputStream(myFile), "UTF8"));
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        // Get iterator to all the rows in current sheet
        Iterator < Row > rowIterator = mySheet.iterator();
        
        //rowIterator.next();

        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell name = row.getCell(0);
            Cell nationalId = row.getCell(2);
            Cell phone = row.getCell(3);
            phone.setCellType(Cell.CELL_TYPE_STRING);
            String nameStr = name.getStringCellValue();
            String nationalIdStr = nationalId.getStringCellValue();
            String phoneStr = phone.getStringCellValue();
            UserRegistration registration = new UserRegistration();
            registration.setUsername(nationalIdStr);
            registration.setBasicPassword("stem123456");
            registration.setType(UserType.REGULAR);
            registration.setRealName(nameStr);
            registration.setPhoneNumber(phoneStr);
            registration.setNationalId(nationalIdStr);
            Client client = Client.create();
            // Get user
            User user;
            WebResource webResource = client
                    .resource("http://www.stemcloud.cn:8080/users").path("findByUsername").queryParam("username", nationalIdStr);
            ClientResponse response = webResource.accept(MediaType.APPLICATION_JSON).get(ClientResponse.class);
            if (response.getStatus() == 404) {
                webResource = client.resource("http://www.stemcloud.cn:8080/users/register");
                response = webResource.type(MediaType.APPLICATION_JSON + ";charset=utf-8")
                        .post(ClientResponse.class, registration);
                if (response.getStatus() != 200) {
                    throw new Exception("response is not 200");
                }
            } else if (response.getStatus() == 200) {
                user = response.getEntity(User.class);
                user.setBasicPassword("stem123456");
                user.setType(UserType.REGULAR);
                user.setRealName(nameStr);
                user.setPhoneNumber(phoneStr);
                user.setNationalId(nationalIdStr);
                Assert.assertEquals(nationalIdStr, user.getUsername());
                webResource = client.resource("http://www.stemcloud.cn:8080/users");
                response = webResource.type(MediaType.APPLICATION_JSON)
                        .put(ClientResponse.class, user);
                if (response.getStatus() != 200) {
                    throw new Exception("response is not 200");
                }
            }
            System.out.println(registration.toString());
            }
    }

}
