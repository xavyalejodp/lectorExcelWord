package com.xdpackages.poi;

import java.io.File;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.xdpackages.config.Configuration;



public class ExcelReader {

	private Logger logger = Logger.getLogger(ExcelReader.class) ;
	
	public boolean readExcelWriteWord(String excelFileName, String wordFileName) throws Exception {
		try {
			
			 // Creating a Workbook from an Excel file (.xls or .xlsx)
	        Workbook workbook = WorkbookFactory.create(new File(excelFileName));

	        // Retrieving the number of sheets in the Workbook
	        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

	        /*
	           =============================================================
	           Iterating over all the sheets in the workbook (Multiple ways)
	           =============================================================
	        */
	        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
	        workbook.forEach(sheet -> {
	            System.out.println("=> " + sheet.getSheetName());
	        });
	        
	        /*
	           ==================================================================
	           Iterating over all the rows and columns in a Sheet (Multiple ways)
	           ==================================================================
	        */

	        // Getting the Sheet at index zero
	        Sheet sheet = workbook.getSheetAt(0);

	        // Create a DataFormatter to format and get each cell's value as String
	        DataFormatter dataFormatter = new DataFormatter();
	        
	     // 3. Or you can use Java 8 forEach loop with lambda
	        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
	        
	        WordWriter wordWriter = new WordWriter(wordFileName);
	        
	        sheet.forEach(	        		        		        
	        	row -> {
	        		int numRow = row.getRowNum();
	        
	        		if(numRow == 0) {
	        			wordWriter.setHeader(row);
	        		}else {
	        			
	        			wordWriter.setDetail(row);
	        			/*row.forEach(	     	        					
		        				cell -> {	            		        					
		        					String cellValue = dataFormatter.formatCellValue(cell);
		        					System.out.print(cellValue + "\t");
		        				});*/
	        		}
	            System.out.println();
	        });
	        
	        wordWriter.closeWord();

	        // Closing the workbook
	        workbook.close();
	        
		} catch (Exception e) {
			logger.error(e);
			throw e;
		}
		
		return false;
		
	}
}
