package org.cran.rexcel;

import java.io.FileOutputStream;
import java.io.IOException;

//import java.io.InputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

public class TestRCellBlock {
  @Rule
  public TemporaryFolder folder = new TemporaryFolder();

  @Test
  public void writeXLSX() throws IOException{
    double[] Xdouble = new double[13];
    double[] Xdates  = new double[13];
    for (int i = 0; i < Xdouble.length; i++){
      Xdouble[i] = 0.123   + (double) i;
      Xdates[i]  = 40170.0 + (double) i;
    }
    Xdouble[3] = RInterface.NA_DOUBLE;
    
    // create a new file
    FileOutputStream out = new FileOutputStream( folder.newFile( "junk.xlsx" ) );
    //FileOutputStream out = new FileOutputStream( "/tmp/junk.xlsx" );
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
    
    RCellBlock block = new RCellBlock(sheet, 0, 0, Xdouble.length+1, 5, true );
    
    block.setColData(0, 1, Xdouble, false, null); // 1st column double
    block.setColData(1, 1, Xdates, false, cs1);   // 2nd column date
    block.setRowData(0, 0, header, false, cs2);   // header with styles 
    System.out.print( "Style: " + block.getCell(0, 0).getCellStyle() );

    wb.write(out);   // write the workbook, 
    out.close();        
  }
}
