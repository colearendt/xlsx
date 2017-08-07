package org.cran.rexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.DateFormatConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DebugRexcel {

	/*
	 * Cannot keep existing formats when using CellBlock.
	 * Take a file with existing formats and write it back.
	 */
	public static void issue22() throws InvalidFormatException, IOException {
		System.out.println( "Testing issue22 =================================" );
		FileInputStream in = new FileInputStream(new File("../resources/issue22.xlsx") );
		Workbook wb = WorkbookFactory.create(in);
		Sheet sheet0 = wb.getSheetAt(0);
		Sheet sheet1 = wb.getSheetAt(1);
		Sheet sheet2 = wb.getSheetAt(2);
		
		Cell c00_before = sheet1.getRow(0).getCell(0);
		CellStyle cs_before = c00_before.getCellStyle();
		System.out.println("Background color before creation is: " + cs_before.getFillBackgroundColor() );
		System.out.println("Cell style is: " +  cs_before.toString() );
		
		double[] data = {1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0};
		System.out.println("On Sheet 1 ....");
		// need to create=true, because sheet is empty!
		RCellBlock cb0 = new RCellBlock(sheet0, 0, 0, 3, 3, true);  
		cb0.setMatrixData(0, 2, 0, 2, data, true, null);
		
		System.out.println("On Sheet 2 ....");
		RCellBlock cb1 = new RCellBlock(sheet1, 0, 0, 3, 3, false);
		Cell c00 = cb1.getCell(0, 0);
		CellStyle cs = c00.getCellStyle();
		System.out.println("Background color after creation is: " +  cs.getFillBackgroundColor() );
		System.out.println("Cell style is: " +  cs.toString() );
	    // good news, that the BackgroundColor in the CellBlock is maintained!
	    cb1.setMatrixData(0, 2, 0, 2, data, true, null);
		
		System.out.println("On Sheet 3 ....");
		RCellBlock cb2 = new RCellBlock(sheet2, 0, 0, 3, 3, false);
		cb2.setMatrixData(0, 2, 0, 2, data, true, null);
		
		
	    FileOutputStream out = new FileOutputStream( "/tmp/issue22_out.xlsx" );
	    wb.write(out);   
	    out.close();        
	    System.out.println("Wrote /tmp/issue22_out.xlsx");
	}
	
	
	
	/*
	 * Does not read integers correctly as Strings "12.0" instead of "12"
	 */
	public static void issue25() throws InvalidFormatException, IOException {
		
		System.out.println( "Testing issue25 =================================" );
		FileInputStream in = new FileInputStream(new File("../resources/issue25.xlsx") );
		Workbook wb = WorkbookFactory.create(in);
		Sheet sheet = wb.getSheetAt(0); 
		
		Row row = sheet.getRow(37);
		Cell cell = row.getCell(0);
		
		System.out.println( cell.getCellType() );
		
		System.out.println( cell.getNumericCellValue() );  // returns 8561807.0 !!  I want "8561807"
				
		RInterface R = new RInterface();
		String[] res = R.readColStrings(sheet, 35, 38, 0);
		for (int i=0; i<res.length; i++){
			System.out.println(res[i]);
		}        
	}
	
	
	/*
	 * Don't hardcode US Locale for datetimes in Excel.  Experiment a bit.
	 */
	public static void issue26() throws InvalidFormatException, IOException {
		
		System.out.println( "Testing issue26 =================================" );
		Workbook wb = new XSSFWorkbook();
	    CreationHelper createHelper = wb.getCreationHelper();
	    Sheet sheet = wb.createSheet("Sheet1");
	    Row row = sheet.createRow(0);
	    	    
	    // first cell
	    Cell cell0 = row.createCell(0);	    
	    CellStyle cellStyle = wb.createCellStyle();
	    cellStyle.setDataFormat(
	        createHelper.createDataFormat().getFormat("m/d/yy h:mm:ss"));	    
	    cell0.setCellValue(new Date(1385903303326L)); // 12/1/2013 08:08 AM
	    cell0.setCellStyle(cellStyle);
	    
	    
	    // second cell using another format with French locale
	    CellStyle cs2 = wb.createCellStyle();
	    String excelFormatPrefix = DateFormatConverter.getPrefixForLocale(Locale.FRENCH);
	    System.out.println("The LOCALE prefix is: " + excelFormatPrefix);
	    String excelFormatPattern = excelFormatPrefix + "dd MMM, yyyy;@";
	    System.out.println("Converted pattern in FRENCH locale is: " + DateFormatConverter.convert(Locale.FRENCH, "m/d/yy h:mm:ss"));
	    
	    DataFormat df = wb.createDataFormat();
	    cs2.setDataFormat(df.getFormat(excelFormatPattern));
	    Cell cell1 = row.createCell(1);
	    cell1.setCellValue(new Date(1385903303326L));
	    cell1.setCellStyle(cs2);
	    
	    FileOutputStream out = new FileOutputStream( "/tmp/issue26_out.xlsx" );
	    wb.write(out);   
	    out.close();        
	    System.out.println("Wrote /tmp/issue26_out.xlsx");
	}
	
	public static void issue35() throws InvalidFormatException, IOException {
		
		System.out.println( "Testing issue35 =================================" );
		FileInputStream in = new FileInputStream(new File("../resources/issue35.xlsx") );
		Workbook wb = WorkbookFactory.create(in);
		Sheet sheet = wb.getSheetAt(0); 
		
		Row row = sheet.getRow(1);
		Cell cell = row.getCell(2);
		
		System.out.println( cell.getCellType() );
		
		System.out.println( cell.getNumericCellValue() );  // returns 8561807.0 !!  I want "8561807"
				
		RInterface R = new RInterface();
		double[] res = R.readColDoubles(sheet, 1, 3, 2);
		for (int i=0; i<res.length; i++){
			System.out.println(res[i]);
		}        
		
		
	}
	
    public static void main(String[] args) throws InvalidFormatException, IOException {
    	
    	System.out.println( new File(".").getCanonicalPath() );
    	
    	//issue22();
    	//issue25();
    	//issue26();
    	issue35();
    	
    	
        
      }
    
      
}
