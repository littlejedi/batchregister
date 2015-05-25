package batchprojectupload;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import javax.ws.rs.core.MediaType;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;
import com.liangzhi.commons.domain.Project;
import com.liangzhi.commons.domain.ProjectAuthor;
import com.liangzhi.commons.domain.ProjectAward;
import com.liangzhi.commons.domain.ProjectCategory;
import com.liangzhi.commons.domain.ProjectFairs;
import com.liangzhi.commons.domain.ProjectView;
import com.sun.jersey.api.client.Client;
import com.sun.jersey.api.client.ClientResponse;
import com.sun.jersey.api.client.WebResource;

public class BatchProjectUpload {

    public static void main(String[] args) throws Exception {
        File myFile = new File("E://liangzhi/30届上海市赛项目摘要.xlsx");
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
        Set<Integer> set = new HashSet<Integer>();
        while (rowIterator.hasNext()) {           
            Row row = rowIterator.next();
            Cell name = row.getCell(0);
            Cell category = row.getCell(1);
            Cell abstractText = row.getCell(2);
            Cell award = row.getCell(3);
            String nameStr = name.getStringCellValue();
            String categoryStr = category.getStringCellValue();
            if (categoryStr.equals("医学与健康")) {
                categoryStr = "医学与健康学";
            }
            String abstractTextStr = abstractText.getStringCellValue();
            String awardStr = award.getStringCellValue();
            ProjectView req = new ProjectView();
            // project
            Project proj = new Project();
            proj.setName(nameStr);
            proj.setParticipatedFairId(ProjectFairs.SHIC_30.getValue());
            proj.setAbstractText(abstractTextStr);
            proj.setCategoryId(ProjectCategory.getIdByName(categoryStr));
            set.add(proj.getCategoryId());
            List<ProjectAward> awards = Lists.newArrayList();
            // award
            ProjectAward aw = new ProjectAward();
            aw.setFairId(ProjectFairs.SHIC_30.getValue());
            aw.setName(awardStr);
            awards.add(aw);
            // empty author list
            List<ProjectAuthor> authors = Lists.newArrayList();
            req.setProject(proj);
            req.setAwards(awards);
            req.setAuthors(authors);
            System.out.println(req.toString());
            Client client = Client.create();
            WebResource webResource = client.resource("http://www.stemcloud.cn:8080/projects").queryParam("index", "true");
            ClientResponse response = webResource.type(MediaType.APPLICATION_JSON)
                    .post(ClientResponse.class, req);
            if (response.getStatus() != 200) {
                System.out.println(req.toString());
                throw new Exception("Response is not 200");
            }
        }
        for (Integer i : set) {
            System.out.println(i);
        }
    }

}
