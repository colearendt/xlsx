package org.cran.rexcel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class RInterface {
	
  public int NCOLS = 0;
  public int NROWS = 0;

  public static final double NA_DOUBLE = Double.NaN;
  public static final int NA_INT = -2147483648;
  public static final String NA_STRING = "NA";
  //public static final int NA_BOOL = -128;

  public static boolean isNA( double x )
  {
      return ( Double.isNaN(x) );
  }

  public static boolean isNA( int x )
  {
      return ( x == NA_INT );
  }

  public static boolean isNA( String x )
  {
      return ( x == null );
  }
  
  
   
   /*
    * Read a column of doubles from the sheet.  If the cell is not a Number show a NaN!
    */
   public double[] readColDoubles(Sheet sheet, int startRowIndex, int endRowIndex, int colIndex) {
     
     int N = endRowIndex - startRowIndex + 1;
     double[] res = new double[N];
     for (int i=0; i<N; i++) {
       Row row = sheet.getRow(startRowIndex+i);
       if (row == null) {
           res[i] = NA_DOUBLE;
           continue;
       } 
       Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
       switch(cell.getCellTypeEnum()) {
       case STRING:
         try {
           res[i] = Double.valueOf(cell.getStringCellValue());
         } catch (NumberFormatException nfe) {
           //System.out.println("NumberFormatException: " + nfe.getMessage());
           res[i] = Double.NaN; 
         }
         break;
       case NUMERIC:
         res[i] = cell.getNumericCellValue();
         break;
       case FORMULA:
         try {
           res[i] = cell.getNumericCellValue();
         } catch (IllegalStateException e) {
           res[i] = Double.NaN;  // cell.getCellFormula();
         }
         break;
       case BOOLEAN:
         res[i] = (double) (cell.getBooleanCellValue()?1:0);
         break;
       case ERROR:
         res[i] = NA_DOUBLE;
         break;
       default:
         res[i] = NA_DOUBLE;
       }
     }
     
     return res;
   }

   /*
    * Read a column of Strings from the sheet.  If the cell type is not a String, convert to String!
    */
   public String[] readColStrings(Sheet sheet, int startRowIndex, int endRowIndex, int colIndex) {
     
     int N = endRowIndex - startRowIndex + 1;
     String[] res = new String[N];
     for (int i=0; i<N; i++) {
       Row row = sheet.getRow(startRowIndex+i);
       if (row == null) {
           res[i] = "";
           continue;
       }
       Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
       switch(cell.getCellTypeEnum()) {
         case STRING:
           res[i] = cell.getStringCellValue();
           break;
         case NUMERIC:
           double val = cell.getNumericCellValue();
           // check if exact integer value, then you want it represented without any 
           // spurious .0 at the end
           if (val == Math.rint(val)) {  
        	   res[i] = Long.toString(Math.round(val));
           } else {
        	   res[i] = Double.toString(val);
           }
           break;
         case FORMULA:
           try {
             res[i] = cell.getStringCellValue();
           } catch (IllegalStateException e) {
             res[i] = NA_STRING;  // cell.getCellFormula();
           }
           break;
         case BOOLEAN:
           res[i] = Boolean.toString(cell.getBooleanCellValue());
           break;
         case BLANK:
           res[i] = "";
           break;
         case ERROR:
           res[i] = "ERROR";
           break;
         default:
           res[i] = "";
       }
       
     }
     
     return res;
   }
 
   
   /*
    * Read a row of Strings from the sheet
    */
   public String[] readRowStrings(Sheet sheet, int startColIndex, int endColIndex, int rowIndex) {
     
     int N = endColIndex - startColIndex + 1;
     String[] res = new String[N];
     Row row = sheet.getRow(rowIndex);
     for (int i=0; i<N; i++) {
       Cell cell = row.getCell(startColIndex+i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
       switch(cell.getCellTypeEnum()) {
       case STRING:
         res[i] = cell.getStringCellValue();
         break;
       case NUMERIC:
         double val = cell.getNumericCellValue();
         // check if exact integer value, then you want it represented without any 
         // spurious .0 at the end
         if (val == Math.rint(val)) {  
       	   res[i] = Long.toString(Math.round(val));
         } else {
       	   res[i] = Double.toString(val);
         }  
         break;
       case FORMULA:
         res[i] = cell.getStringCellValue();
         break;
       case BOOLEAN:
         res[i] = Boolean.toString(cell.getBooleanCellValue());
         break;
       default:
         res[i] = "";
       }
     }
     
     return res;
   }
   
}  
