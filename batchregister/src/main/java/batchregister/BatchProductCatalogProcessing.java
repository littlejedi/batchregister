package batchregister;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Strings;
import com.google.common.collect.Lists;

import model.BomListRow;
import utils.FileExportUtils;

public class BatchProductCatalogProcessing {
	
	public static void main(String[] args) throws Exception {
		
		  int base = 11;
		
		  File myFile = new File("C://Nick/stemcloud/productcatalog.xlsx");
	      //BufferedReader fis = new BufferedReader(new InputStreamReader(new FileInputStream(myFile), "UTF8"));
	      FileInputStream fis = new FileInputStream(myFile);

	      // Finds the workbook instance for XLSX file
	      XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
	      
	      // The workbook we are writing to
	      XSSFWorkbook newWorkBook = new XSSFWorkbook();
	      XSSFSheet bomListSheet = newWorkBook.createSheet("配件BOMList");
	      FileExportUtils.insertRow(bomListSheet, "产品编号", "产品名称", "物料编号", "物料名称", "物料数量");

	      // Return 4th sheet from the XLSX workbook
	      XSSFSheet mySheet = myWorkBook.getSheetAt(0);

	      // Get iterator to all the rows in current sheet
	      Iterator < Row > rowIterator = mySheet.iterator();
	        
	      rowIterator.next();

	      // Traversing over each row of XLSX file
	      while (rowIterator.hasNext()) {
	    	  Row row = rowIterator.next();
	    	  Cell productId = row.getCell(1);
	    	  Cell productName = row.getCell(base);
	    	  Cell materialId = row.getCell(base+1);
	    	  Cell materialName = row.getCell(base+2);
	    	  Cell materialQuantity = row.getCell(base+3);
	    	  Cell materialIdTwo = row.getCell(base+4);
	    	  Cell materialIdTwoName = row.getCell(base+5);
	    	  Cell materialIdThree = row.getCell(base+6);
	    	  Cell materialIdThreeName = row.getCell(base+7);
	    	  Cell materialIdFour = row.getCell(base+8);
	    	  Cell materialIdFourName = row.getCell(base+9);
	    	  if (productId == null) {
	    		  break;
	    	  }
	    	  productId.setCellType(Cell.CELL_TYPE_STRING);
	    	  productName.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialId.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialName.setCellType(Cell.CELL_TYPE_STRING);
	    	  materialQuantity.setCellType(Cell.CELL_TYPE_STRING);
	    	  if (materialIdTwo != null) {
	    		  materialIdTwo.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  if (materialIdTwoName != null) {
	    		  materialIdTwoName.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  if (materialIdThree != null) {
	    		  materialIdThree.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  if (materialIdThreeName != null) {
	    		  materialIdThreeName.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  if (materialIdFour != null) {
	    		  materialIdFour.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  if (materialIdFourName != null) {
	    		  materialIdFourName.setCellType(Cell.CELL_TYPE_STRING);
	    	  }
	    	  final String productIdString = Strings.nullToEmpty(productId.getStringCellValue());
	    	  final String productNameString = Strings.nullToEmpty(productName.getStringCellValue());
	    	  final String materialIdString = Strings.nullToEmpty(materialId.getStringCellValue());
	    	  final String materialNameString = Strings.nullToEmpty(materialName.getStringCellValue());
	    	  final String materialQuantityString = Strings.nullToEmpty(materialQuantity.getStringCellValue());
	    	  final String materialIdTwoString = materialIdTwo != null ? Strings.nullToEmpty(materialIdTwo.getStringCellValue()) : "";
	    	  final String materialNameTwoString = materialIdTwoName != null ? Strings.nullToEmpty(materialIdTwoName.getStringCellValue()): "";
	    	  final String materialIdThreeString = materialIdThree != null ? Strings.nullToEmpty(materialIdThree.getStringCellValue()): "";
	    	  final String materialNameThreeString = materialIdThreeName != null ? Strings.nullToEmpty(materialIdThreeName.getStringCellValue()): "";
	    	  final String materialIdFourString = materialIdFour != null ? Strings.nullToEmpty(materialIdFour.getStringCellValue()): "";
	    	  final String materialNameFourString = materialIdFourName != null ? Strings.nullToEmpty(materialIdFourName.getStringCellValue()): "";
	    	  StringBuffer buffer = new StringBuffer();
	    	  buffer.append("产品编号:");
	    	  buffer.append(productIdString);
	    	  buffer.append(", 产品名称:");
	    	  buffer.append(productNameString);
	    	  buffer.append(", 物料编码1:");
	    	  buffer.append(materialIdString);
	    	  buffer.append(", 物料名称1:");
	    	  buffer.append(materialNameString);
	    	  buffer.append(", 物料数量1:");
	    	  buffer.append(materialQuantityString);
	    	  buffer.append(", 物料编码2:");
	    	  buffer.append(materialIdTwoString);
	    	  buffer.append(", 物料名称2:");
	    	  buffer.append(materialNameTwoString);
	    	  buffer.append(", 物料编码3:");
	    	  buffer.append(materialIdThreeString);
	    	  buffer.append(", 物料名称3:");
	    	  buffer.append(materialNameThreeString);
	    	  buffer.append(", 物料编码4:");
	    	  buffer.append(materialIdFourString);
	    	  buffer.append(", 物料名称4:");
	    	  buffer.append(materialNameFourString);
	    	  System.out.println(buffer.toString());
	    	  List<BomListRow> rows = Lists.newArrayList();
	    	  if (!Strings.isNullOrEmpty(materialIdString) && !Strings.isNullOrEmpty(materialNameString)) {
	    		  rows.add(new BomListRow(productIdString, productNameString, materialIdString, materialNameString, materialQuantityString));
	    	  }
              if (!Strings.isNullOrEmpty(materialIdTwoString) && !Strings.isNullOrEmpty(materialNameTwoString)) {
            	  rows.add(new BomListRow(productIdString, productNameString, materialIdTwoString, materialNameTwoString, "1"));
	    	  }
              if (!Strings.isNullOrEmpty(materialIdThreeString) && !Strings.isNullOrEmpty(materialNameThreeString)) {
            	  rows.add(new BomListRow(productIdString, productNameString, materialIdThreeString, materialNameThreeString, "1"));
	    	  }
              if (!Strings.isNullOrEmpty(materialIdFourString) && !Strings.isNullOrEmpty(materialNameFourString)) {
            	  rows.add(new BomListRow(productIdString, productNameString, materialIdFourString, materialNameFourString, "1"));
	    	  }
              for (BomListRow bomListRow : rows) {
            	  FileExportUtils.insertRow(bomListSheet, bomListRow.getProductSerialId(), bomListRow.getProductName(), bomListRow.getMaterialSerialId(), bomListRow.getMaterialName(), bomListRow.getMaterialQuantity());
              }
	      }
	      
	      FileOutputStream out = new FileOutputStream("C://Nick/stemcloud/bomlist.xlsx");
	      newWorkBook.write(out);
	      
	}

}
