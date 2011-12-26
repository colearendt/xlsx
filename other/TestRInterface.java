package tests;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
//import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dev.RInterface;


public class TestRInterface {
  
  public static void testWrite() throws IOException{
    RInterface R = new RInterface(); // create the R interface
    R.NCOLS = 15;
    R.NROWS = 13000;
    //R.CELL_ARRAY = new Cell[R.NROWS][R.NCOLS];
    
    double[] X = new double[R.NROWS];
    for (int i = 0; i < R.NROWS; i++){
      X[i] = 0.123 + (double) i;
    }
    
    // create a new file
    FileOutputStream out = new FileOutputStream("C:/Temp/junk.xlsx");
    Workbook wb = new XSSFWorkbook();  // create a new workbook
    Sheet sheet = wb.createSheet();      // create a new sheet
    
    R.createCells(sheet, 0, 0);
    for (int j = 0; j < R.NCOLS; j++) {  // write one column at a time
      R.writeColDoubles(sheet, 0, j, X);
    }
    
    wb.write(out);   // write the workbook, 
    out.close();        
  }
  
  public static void testRead() throws InvalidFormatException, IOException{
    RInterface R = new RInterface(); // create the R interface
 
//    FileInputStream in = new FileInputStream("C:/Temp/ModelList.xlsx");
//    FileInputStream in = new FileInputStream("H:/user/R/Adrian/findataweb/temp/xlsx/trunk/inst/tests/test_import.xlsx");
    FileInputStream in = new FileInputStream("S:/All/Structured Risk/NEPOOL/Incs & Decs/Apr 11/Inc Dec 15Apr11 Incs and Decs.xls");
    Workbook wb = WorkbookFactory.create(in);
    Sheet sheet = wb.getSheetAt(3); 
    
    //String[] res = R.readRowStrings(sheet, 0, 10, 0);
    String[] res = R.readColStrings(sheet, 1, 20, 1);
//    double[] res = R.readColDoubles(sheet, 1, 20, 1);
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
