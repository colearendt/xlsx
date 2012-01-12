package tests;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
//import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dev.RInterface;


public class TestRInterface {
  
  public static void testWrite() throws IOException{
    RInterface R = new RInterface(); // create the R interface
    R.NCOLS = 5;
    R.NROWS = 13;
    
    double[] Xdouble = new double[R.NROWS];
    double[] Xdates  = new double[R.NROWS];
    for (int i = 0; i < R.NROWS; i++){
      Xdouble[i] = 0.123   + (double) i;
      Xdates[i]  = 40170.0 + (double) i;
    }
    Xdouble[3] = Double.NaN;  
    
    // create a new file
    FileOutputStream out = new FileOutputStream("C:/Temp/junk.xlsx");
    Workbook wb = new XSSFWorkbook();    // create a new workbook
    Sheet sheet = wb.createSheet();      // create a new sheet
    
    CellStyle cs1 = wb.createCellStyle();
    DataFormat fmt1 = wb.createDataFormat();
    cs1.setDataFormat(fmt1.getFormat("m/d/yyyy"));
    
    CellStyle cs2 = wb.createCellStyle();
    cs2.setBorderBottom(CellStyle.BORDER_THICK);
    cs2.setFillForegroundColor((short) 5);    
    cs2.setFillBackgroundColor((short) 6);
    cs2.setFillPattern(CellStyle.SOLID_FOREGROUND); 
    
    String[] header = {"Number", "Date"};
    
    R.createCells(sheet, 0, 0);
    R.writeColDoubles(sheet, 0, 0, Xdouble, false);         // 1st column double
    R.writeColDoubles(sheet, 0, 1, Xdates, false, cs1);    // 2nd column date
    R.writeRowStrings(sheet, 0, 0, header, cs2);     // header with styles 
    
    
    wb.write(out);   // write the workbook, 
    out.close();        
  }
  
  public static void testRead() throws InvalidFormatException, IOException{
    RInterface R = new RInterface(); // create the R interface
 
    FileInputStream in = new FileInputStream("/home/adrian/Documents/rexcel/trunk/inst/tests/test_import.xlsx");
//    FileInputStream in = new FileInputStream("H:/user/R/Adrian/findataweb/temp/xlsx/trunk/inst/tests/test_import.xlsx");
//    FileInputStream in = new FileInputStream("S:/All/Structured Risk/NEPOOL/Incs & Decs/Apr 11/Inc Dec 15Apr11 Incs and Decs.xls");
    Workbook wb = WorkbookFactory.create(in);
    Sheet sheet = wb.getSheetAt(7); 
    
    //String[] res = R.readRowStrings(sheet, 0, 10, 0);
//    String[] res = R.readColStrings(sheet, 1, 20, 1);
    double[] res = R.readColDoubles(sheet, 1, 10, 4);
    //double[] res = R.readColDoubles(sheet, 0, 2006, 0);
    for (int i=0; i<res.length; i++){
      System.out.println(res[i]);
    }
  }
  
  public static void main(String[] args) throws IOException, InvalidFormatException {
		
    //testWrite();    
	testRead();  
      
    System.out.println("Done!");

  }
	
}
