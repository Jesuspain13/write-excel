package es.uma.helloworld.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteMethod {
	
	public void write(int rowSelected, int cellSelected) {
		// Create a Workbook
		File f = this.searchFile();
		
		Workbook workbook = null;
		Sheet sheet = null;
		if (f != null) {
			try {
				FileInputStream fip = new FileInputStream(f); 
				workbook = WorkbookFactory.create(fip);
				sheet = workbook.getSheet("Hello world sheet");
			} catch (Exception ex) {
				System.out.println("invalid format ex");
			}
		} else {
			workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
			sheet = workbook.createSheet("Hello world sheet");
		}
		
        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Row
        
        Row headerRow = sheet.getRow(rowSelected);
        if (headerRow == null) {
        	headerRow = sheet.createRow(rowSelected);
        }
        // Create cells
        Cell cell = headerRow.getCell(cellSelected);
        if (cell == null) {
        	cell = headerRow.createCell(cellSelected);
        	cell.setCellValue("Hello world nuevo");
        } else {
        	cell.setCellValue("Hello world nuevo");
        }

        // Write the output to a file
        try {
        	FileOutputStream fileOut = null;
        	if (f != null) {
        		fileOut = new FileOutputStream(f);
        	} else {
        		fileOut = new FileOutputStream(new File("poi-generated.xlsx"));
        	}
	        workbook.write(fileOut);
	        fileOut.close();
	
	        // Closing the workbook
	        workbook.close();
	        System.out.println("Completado con exito");
        } catch(Exception ex) {
	    	ex.printStackTrace();
	    }
    }
	
	private File searchFile() {
		File f = new File("C:\\Users\\Jesus\\eclipse-workspace\\app");
		File[] matchingFiles = f.listFiles(new FilenameFilter() {
			
			public boolean accept(File dir, String name) {
				return name.startsWith("poi") && name.endsWith("xlsx");
			}
		});
		
		if (matchingFiles != null) {
			return matchingFiles[0];
		} else {
			return null;
		}
	}


}
