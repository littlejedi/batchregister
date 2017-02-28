package batchregister;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BatchProductCatalogProcessing {
	
	
	public static void main(String[] args) throws Exception {
		
		  File myFile = new File("C://Nick/stemcloud/catalog/test.xlsx");
	      //BufferedReader fis = new BufferedReader(new InputStreamReader(new FileInputStream(myFile), "UTF8"));
	      FileInputStream fis = new FileInputStream(myFile);

	      // Finds the workbook instance for XLSX file
	      XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

	      // Return 4th sheet from the XLSX workbook
	      XSSFSheet mySheet = myWorkBook.getSheetAt(4);

	      // Get iterator to all the rows in current sheet
	      Iterator < Row > rowIterator = mySheet.iterator();
	        
	      rowIterator.next();

	      // Traversing over each row of XLSX file
	      while (rowIterator.hasNext()) {
	    	  Row row = rowIterator.next();
	    	  Cell courseProductId = row.getCell(1);
	    	  Cell productName = row.getCell(2);
	    	  Cell materialId = row.getCell(3);
	    	  Cell materialName = row.getCell(4);
	    	  Cell materialQuantity = row.getCell(5);
	    	  Cell materialIdTwo = row.getCell(6);
	    	  Cell materialIdTwoName = row.getCell(7);
	    	  Cell materialIdThree = row.getCell(8);
	    	  Cell materialIdThreeName = row.getCell(9);
	    	  Cell materialIdFour = row.getCell(10);
	    	  Cell materialIdFourName = row.getCell(11);
	    	  courseProductId.setCellType(Cell.CELL_TYPE_STRING);
	    	  productName.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialId.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialName.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialQuantity.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialIdTwo.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialIdTwoName.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialIdThree.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialIdThreeName.setCellType(Cell.CELL_TYPE_STRING);
	    	  if (materialIdFour != null) {
	    		  materialIdFour.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  if (materialIdFourName != null) {
	    		  materialIdFourName.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  StringBuffer buffer = new StringBuffer();
	    	  buffer.append("课程来源:");
	    	  buffer.append(courseProductId);
	    	  buffer.append(", 产品名称:");
	    	  buffer.append(productName);
	    	  buffer.append(", 物料编码1:");
	    	  buffer.append(materialId);
	    	  buffer.append(", 物料名称1:");
	    	  buffer.append(materialName);
	    	  buffer.append(", 物料数量1:");
	    	  buffer.append(materialQuantity);
	    	  buffer.append(", 物料编码2:");
	    	  buffer.append(materialIdTwo);
	    	  buffer.append(", 物料名称2:");
	    	  buffer.append(materialIdTwoName);
	    	  buffer.append(", 物料编码3:");
	    	  buffer.append(materialIdThree);
	    	  buffer.append(", 物料名称3:");
	    	  buffer.append(materialIdThreeName);
	    	  buffer.append(", 物料编码4:");
	    	  buffer.append(materialIdFour);
	    	  buffer.append(", 物料名称4:");
	    	  buffer.append(materialIdFourName);
	    	  System.out.println(buffer.toString());
	      }
	}

}
