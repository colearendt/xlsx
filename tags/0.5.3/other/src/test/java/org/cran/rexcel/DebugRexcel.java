package org.cran.rexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DebugRexcel {
    
    public static void main(String[] args) throws InvalidFormatException, IOException {
//        FileInputStream in = new FileInputStream(new File("C:/temp/fca3_monthly_ob_v2.xls") );
//        Workbook wb = WorkbookFactory.create(in);
//        Sheet sheet = wb.getSheetAt(0); 
//
//        RInterface R = new RInterface();
//
//        int maxRow = sheet.getLastRowNum();
//        System.out.println(maxRow);
//        
//        //String[] res = R.readRowStrings(sheet, 0, 10, 0);
//        String[] res = R.readColStrings(sheet, 2075, 2080, 0);
//        //double[] res = R.readColDoubles(sheet, 0, 2006, 0);
//        for (int i=0; i<res.length; i++){
//          System.out.println(res[i]);
//        }
//        
        TestRInterface ti = new TestRInterface();
        //ti.issue13();
        
      }
    
      
}
