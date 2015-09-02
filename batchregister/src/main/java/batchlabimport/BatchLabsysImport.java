package batchlabimport;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import javax.ws.rs.core.MediaType;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.liangzhi.commons.domain.Lab;
import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.ClientResponse;
import com.sun.jersey.api.client.WebResource;

public class BatchLabsysImport {

    public static void main(String[] args) throws Exception {
        File myFile = new File("E://liangzhi/labsys.xlsx");
        //BufferedReader fis = new BufferedReader(new InputStreamReader(new FileInputStream(myFile), "UTF8"));
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        // Get iterator to all the rows in current sheet
        Iterator < Row > rowIterator = mySheet.iterator();
        
        rowIterator.next();
        
        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {    
            Row row = rowIterator.next();
            row.getCell(11).setCellType(Cell.CELL_TYPE_STRING);
            row.getCell(16).setCellType(Cell.CELL_TYPE_STRING);
            row.getCell(17).setCellType(Cell.CELL_TYPE_STRING);
            String recSociety = row.getCell(0).getStringCellValue();
            String name = row.getCell(1).getStringCellValue();
            String institution = row.getCell(2).getStringCellValue();
            String researchArea = row.getCell(3).getStringCellValue();
            String detailedAddr = row.getCell(5).getStringCellValue();
            String description = row.getCell(6).getStringCellValue();
            String conditions = row.getCell(7).getStringCellValue();
            String openTime = row.getCell(8).getStringCellValue();
            String discipline = row.getCell(9).getStringCellValue();
            String subDiscipline = row.getCell(10).getStringCellValue();
            String capacity = row.getCell(11).getStringCellValue();
            String contactName = row.getCell(14).getStringCellValue();
            String contactTitle = row.getCell(15).getStringCellValue();
            String telephone = row.getCell(16).getStringCellValue();
            String mobile = row.getCell(17).getStringCellValue();
            String email = row.getCell(18).getStringCellValue();
            String title = row.getCell(19).getStringCellValue();
            String telephoneAndMobile = "固定电话：" + telephone + " 手机：" + mobile;
            Lab lab = new Lab();
            lab.setRecSociety(recSociety);
            lab.setName(name);
            lab.setInstitution(institution);
            lab.setResearchArea(researchArea);;
            lab.setCity("上海");
            lab.setDetailedAddress(detailedAddr);
            lab.setDescription(description);
            lab.setCondition(conditions);
            lab.setOpenTime(openTime);
            lab.setDiscipline(discipline);
            lab.setSubDiscipline(subDiscipline);
            lab.setCapacity(capacity);
            lab.setContactName(contactName);
            lab.setContactTitle(contactTitle);
            lab.setContactPhone(telephoneAndMobile);
            lab.setContactEmail(email);
            lab.setContactTitle(title);
            System.out.println(lab.toString());
            Client client = Client.create();
            WebResource webResource = client.resource("http://www.stemcloud.cn:8080/labsys/labs");
            ClientResponse response = webResource.type(MediaType.APPLICATION_JSON_TYPE).post(ClientResponse.class, lab);
            if (response.getStatus() != 200) {
                throw new Exception("Failed to create lab:" + lab.toString());
            }
        }
    }

}
