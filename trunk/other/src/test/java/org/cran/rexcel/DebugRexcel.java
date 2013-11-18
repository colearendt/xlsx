package org.cran.rexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DebugRexcel {

	/*
	 * Cannot keep existing formats when using CellBlock.
	 * Take a file with existing formats and write it back.
	 */
	public static void issue22() throws InvalidFormatException, IOException {
		System.out.println( "Testing issue22 =================================" );
		FileInputStream in = new FileInputStream(new File("/home/adrian/Documents/rexcel/trunk/resources/issue22.xlsx") );
		Workbook wb = WorkbookFactory.create(in);
		Sheet sheet0 = wb.getSheetAt(0);
		Sheet sheet1 = wb.getSheetAt(1);
		Sheet sheet2 = wb.getSheetAt(2);
		
		Cell c00_before = sheet1.getRow(0).getCell(0);
		CellStyle cs_before = c00_before.getCellStyle();
		System.out.println( cs_before.getFillBackgroundColor() );
		
		
		double[] data = {1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0};
//		RCellBlock cb0 = new RCellBlock(sheet0, 0, 0, 3, 3, true);
//		System.out.println("sheet0");	
//		cb0.setMatrixData(0, 2, 0, 2, data, true, null, true);
		
		RCellBlock cb1 = new RCellBlock(sheet1, 0, 0, 3, 3, true);
		Cell c00 = cb1.getCell(0, 0);
		CellStyle cs = c00.getCellStyle();
		System.out.println( cs.getFillBackgroundColor() );
		// good news, that the BackgroundColor in the CellBlock is maintained!
		cb1.setMatrixData(0, 2, 0, 2, data, true, null, true);
		
//		RCellBlock cb2 = new RCellBlock(sheet2, 0, 0, 3, 3, true);
//		cb2.setMatrixData(0, 2, 0, 2, data, true, null, true);
//		System.out.println("sheet2");	
		
		
	    FileOutputStream out = new FileOutputStream( "/tmp/issue22_out.xlsx" );
	    wb.write(out);   
	    out.close();        

	}
	
	
	
	/*
	 * Does not read integers correctly as Strings "12.0" instead of "12"
	 */
	public static void issue25() throws InvalidFormatException, IOException {
		
		System.out.println( "Testing issue25 =================================" );
		FileInputStream in = new FileInputStream(new File("/home/adrian/Documents/rexcel/trunk/resources/issue25.xlsx") );
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
	
    public static void main(String[] args) throws InvalidFormatException, IOException {
    	
    	issue22();
    	//issue25();
    	
    	//TestRInterface ti = new TestRInterface();
        //ti.issue13();
    	
    	
    	
        
      }
    
      
}
