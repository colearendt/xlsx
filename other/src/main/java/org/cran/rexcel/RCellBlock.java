package org.cran.rexcel;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * A rectangular block of Excel sheet cells.
 * The object serves as a wrapper of Apache POI
 * providing a way to do bulk cell value assignments and cell style modifications.
 * It also provide cell block-relative indexing of its cells.
 * 
 * @author Adrian A. Dragulescu
 * @author Alexey Stukalov
 */
public class RCellBlock {

	public int noRows;
	public int noCols;
    private Cell[][] cells;

    /**
     * Constructs continuous cell block.
     * 
     * @param sheet sheet of a cell block
     * @param startRowIndex starting row of a block in a sheet.
     * @param startColIndex starting column of a block in a sheet.
     * @param nRows numbers of rows in a block
     * @param nCols number of columns in a block
     * @param create if true, rows and cells are created as necessary (erasing 
     * 	 existing content and style).  If false the existing cells are used and 
     *   content and style is not erased.
     */
    public RCellBlock( Sheet sheet, int startRowIndex, int startColIndex, 
    	int nRows, int nCols, boolean create)
    {
        cells = new Cell[nCols][nRows];
        noRows = nRows;
        noCols = nCols;
       
		for (int i = 0; i < nRows; i++) {
			Row r = sheet.getRow(startRowIndex + i);
			if (r == null) { // row is not already there
				if (create)
					r = sheet.createRow(startRowIndex + i);
				else
					throw new RuntimeException(
							"Row "
									+ (startRowIndex + i)
									+ " doesn't exist in the sheet.  Need to call with argument create=TRUE.");
			}
			
			for (int j = 0; j < nCols; j++) {
				cells[j][i] = create ? r.createCell(startColIndex+j)
                        : r.getCell(startColIndex+j, Row.CREATE_NULL_AS_BLANK);
			}
		}
	}

    /**
     * Gets the cell at given position.
     * @param rowIndex 0-based row index relative to the block 
     * @param colIndex 0-based column index relative to the block
     * @return Cell object
     */
    public Cell getCell( int rowIndex, int colIndex )
    {
        return cells[ colIndex ][ rowIndex ];
    }

    /**
     * Gets the sheet of the cells block.
     * @return
     */
    public Sheet getSheet()
    {
        return ( cells != null & cells.length >= 1
                 & cells[0] != null & cells[0].length >= 1
                 & cells[0][0] != null
                 ? cells[0][0].getSheet() : null );
    }

    /**
     * Gets the workbook of the cells block.
     * @return
     */
    public Workbook getWorkbook()
    {
        final Sheet sheet = getSheet();
        return ( sheet != null ? sheet.getWorkbook() : null );
    }

    public boolean isXSSF()
    {
        return ( getSheet() instanceof XSSFSheet );
    }   
    
    
    /**
     * Writes a column of data to the sheet.
     * Use for numerics, Dates, DateTimes... 
     */
    public void setColData( int colIndex, int rowOffset, double[] data, boolean showNA, CellStyle style ){
        final Cell[] colCells = cells[colIndex];
        for (int i=0; i<data.length; i++) {
            if ( showNA || !RInterface.isNA(data[i])) {
                colCells[rowOffset+i].setCellValue(data[i]);
            } else {
                colCells[rowOffset+i].setCellType(Cell.CELL_TYPE_BLANK);
            }
            if ( style != null ) setCellStyle(style, rowOffset+i, colIndex);
        }
    }

    /**
     * Writes a column of integer to the sheet. 
     */
    public void setColData( int colIndex, int rowOffset, int[] data, boolean showNA, CellStyle style ){
        final Cell[] colCells = cells[colIndex];
        for (int i=0; i<data.length; i++) {
            if ( showNA || !RInterface.isNA(data[i])) {
                colCells[rowOffset+i].setCellValue(data[i]);
            } else {
                colCells[rowOffset+i].setCellType(Cell.CELL_TYPE_BLANK);
            }
            if ( style != null ) setCellStyle(style, rowOffset + i, colIndex);
        }
    }

    /**
     * Writes a column of strings to the sheet.
     */
    public void setColData( int colIndex, int rowOffset, String[] data, boolean showNA, CellStyle style ){
        final Cell[] colCells = cells[colIndex];
        for (int i=0; i<data.length; i++) {
            if ( showNA || !RInterface.isNA(data[i])) {
                colCells[rowOffset+i].setCellValue(data[i]);
            } else {
                colCells[rowOffset+i].setCellType(Cell.CELL_TYPE_BLANK);
            }
            if ( style != null ) setCellStyle(style, rowOffset+i, colIndex);
        }
    }

    /**
     * Writes a row of strings to the sheet.
     * @param rowIndex is relative to the CellBlock
     * @param colOffset is relative to the CellBlock.  If colOffset==0, start from 
     *    first column of the CellBlock.  
     */
    public void setRowData( int rowIndex, int colOffset, String[] data, boolean showNA, CellStyle style ){
        for (int i=0; i<data.length; i++) {
            if ( showNA || !RInterface.isNA(data[i])) {
                cells[colOffset+i][rowIndex].setCellValue(data[i]);
            } else {
                cells[colOffset+i][rowIndex].setCellType(Cell.CELL_TYPE_BLANK);
            }
            if ( style != null ) setCellStyle(style, rowIndex, colOffset+i);
        }
    }

    /**
     * Writes a matrix of data to the sheet.  Useful when you have a lot 
     * of data to write at once.  Use for numerics, Dates, DateTimes...
     * The index argument are all relative to the CellBlock top cell!  
     * @param newStyle, if null do nothing to the existing cell style.  If not 
     *   null, replace the cell style with this cell style.
     */
    public void setMatrixData( int startRow, int endRow, int startColumn, 
    	int endColumn, double[] data, boolean showNA, CellStyle newStyle) { 
    	
		for (int j = startColumn; j <= endColumn; j++) {
			for (int i = startRow; i <= endRow; i++) {
				if (showNA || !RInterface.isNA(data[i])) {
					cells[j][i].setCellValue(data[(endRow - startRow + 1)
							* (j - startColumn) + (i - startRow)]);
				} else {
					cells[j][i].setCellType(Cell.CELL_TYPE_BLANK);
				}
				if (newStyle != null) {
					setCellStyle(newStyle, i, j);
				}
			}
		}
	}

    /**
     * Writes a matrix of data to the sheet.  Useful when you have a lot 
     * of data to write at once.  Use for numerics, Dates, DateTimes...
     * The index argument are all relative to the CellBlock top cell!  
     */
    public void setMatrixData( int startRow, int endRow, int startColumn, 
    	int endColumn, String[] data, boolean showNA, CellStyle style ){
    	
        for (int j=startColumn; j<=endColumn; j++) {
        	for (int i=startRow; i<=endRow; i++) {
        		if ( showNA || !RInterface.isNA(data[i])) {
        			cells[j][i].setCellValue(data[(endRow-startRow+1)*(j-startColumn) + (i-startRow)]);
        		} else {
        			cells[j][i].setCellType(Cell.CELL_TYPE_BLANK);
        		}
        		if ( style != null ) setCellStyle(style, i, j);
        	}
        }
    }
    
    
    /**
     * Set the style of cells in a column. 
     */
    public void setColCellStyle( CellStyle style, int colIndex, int rowOffset, int length )
    {
        final Cell[] colCells = cells[colIndex];
        for (int i=0; i<length; i++) {
            colCells[rowOffset+i].setCellStyle( style );
        }
    }

    /**
     * Sets the style of cells in a row. 
     */
    public void setRowCellStyle( CellStyle style, int rowIndex, int colOffset, int length )
    {
        for (int i=0; i<length; i++) {
            cells[colOffset+i][rowIndex].setCellStyle( style );
        }
    }

    /**
     * Set the style of cells in a sub-block.
     */
    public void setCellStyle( CellStyle style, int[] rowIndices, int[] colIndices )
    {
        for ( int colIx : colIndices ) {
            final Cell[] colCells = cells[colIx];
            for ( int rowIx : rowIndices ) {
                colCells[rowIx].setCellStyle( style );
            }
        }
    }

    /**
     * Set the style of one cell.
     */
    private void setCellStyle( CellStyle style, int rowIx, int colIx )
    {
    	cells[colIx][rowIx].setCellStyle( style );
    }
    
    
    /**
     * Sets the style of all cells in a block
     */
    public void setCellStyle( CellStyle style )
    {
        for (Cell[] colCells : cells) {
          for (Cell cell : colCells) {
             cell.setCellStyle(style);
          }
        }
    }

    /**
     * Interface for modifying cell styles.
     * @see modifyCellStyle()
     */
    protected interface CellStyleModifier {
        /**
         * Changes the given cell style
         * @param style style to modify
         */
        public void modify( CellStyle style );
    }

    /**
     * Modifies existing cell styles of a group of cells.
     *     
     * Implementation note. The methods add new styles to the workbook.
     * It tries to avoid creating duplicate cell styles during same call,
     * but it doesn't lookup and reuse the styles that were already present
     * in the workbook before the method was called, so the duplicate styles 
     * might be generated.
     * 
     * @param rowIndices 0-based block-relative indices of rows
     * @param colIndices 0-based block-relative indices of columns
     * @param modifier cell styles modifier
     */
    protected void modifyCellStyle( int[] rowIndices, int[] colIndices, CellStyleModifier modifier )
    {
    	if ( rowIndices.length != colIndices.length ) 
    		throw new RuntimeException("Length of indRows should equal length of indCols!");
        Map<Short, CellStyle> styleMap = new HashMap<Short,CellStyle>();
        CellStyle defaultStyle = null;
        
        for ( int i=0; i<colIndices.length; i++ ) {
        	Cell cell = cells[colIndices[i]][rowIndices[i]];
        	CellStyle style = cell.getCellStyle();
        	if ( style == null ) {
        		if ( defaultStyle == null ) {
        			defaultStyle = getWorkbook().createCellStyle();
        			modifier.modify( defaultStyle );
        		}
        		cell.setCellStyle(defaultStyle);
        	}
        	else {               
        		if ( styleMap.containsKey( style.getIndex() ) ) {  // has been cached 
        			cell.setCellStyle( styleMap.get( style.getIndex() ) );
        		}
        		else {      // create a new style, and cache it
        			CellStyle modStyle = getWorkbook().createCellStyle();
        			modStyle.cloneStyleFrom( style );
        			modifier.modify(modStyle);
        			styleMap.put(style.getIndex(), modStyle);
        			cell.setCellStyle( modStyle );
        		}
        	}
        }
    }

    /**
     * Sets the fill style of a given sub-block of cells.
     * The other properties of cell styles are preserved.
     * @see modifyCellStyle() 
     */
    public void setFill( final short foreground, final short background, 
    	final short pattern, int[] rowIndices, int[] colIndices )
    {
        modifyCellStyle(rowIndices, colIndices, new CellStyleModifier() {
            public void modify( CellStyle style ) {
                style.setFillPattern( pattern );
                style.setFillForegroundColor( foreground );
                style.setFillBackgroundColor( background );
            }
        } );
    }

    
    /**
     * Sets the fill style of a given sub-block of cells, XSSF version.
     * The other properties of cell styles are preserved.
     * @see modifyCellStyle() 
     */
    public void setFill( final XSSFColor foreground, final XSSFColor background, final short pattern,
                         int[] rowIndices, int[] colIndices )
    {
        modifyCellStyle(rowIndices, colIndices, new CellStyleModifier() {
            public void modify( CellStyle style ) {
                style.setFillPattern( pattern );
                if ( style instanceof XSSFCellStyle ) {
                    XSSFCellStyle xssfStyle = (XSSFCellStyle)style;
                    xssfStyle.setFillForegroundColor( foreground );
                    xssfStyle.setFillBackgroundColor( background );
                } else throw new RuntimeException( "Current cell style doesn't support XSSF colors");
            }
        } );
    }

    
    /**
     * Sets the font of a given sub-block of cells.
     * The other properties of cell styles are preserved.
     * @see modifyCellStyle() 
     */
    public void setFont( final Font font, int[] rowIndices, int[] colIndices )
    {
        modifyCellStyle(rowIndices, colIndices, new CellStyleModifier() {
            public void modify( CellStyle style ) {
                style.setFont( font );
            }
        } );
    }

    
    /**
     * Sets the border style of a given sub-block of cells.
     * If provided border type if BORDER_NONE the corresponding border is not modified.
     * The other properties of cell styles are preserved.
     * @see modifyCellStyle() 
     */
    public void putBorder( final short borderTop, final short topBorderColor,
                           final short borderBottom, final short bottomBorderColor,
                           final short borderLeft, final short leftBorderColor,
                           final short borderRight, final short rightBorderColor,
                           int[] rowIndices, int[] colIndices )
    {
        modifyCellStyle(rowIndices, colIndices, new CellStyleModifier() {
            public void modify( CellStyle style ) {
                if ( borderTop != CellStyle.BORDER_NONE ) {
                    style.setBorderTop( borderTop );
                    style.setTopBorderColor( topBorderColor );
                }
                if ( borderBottom != CellStyle.BORDER_NONE ) {
                    style.setBorderBottom( borderBottom );
                    style.setBottomBorderColor( bottomBorderColor );
                }
                if ( borderLeft != CellStyle.BORDER_NONE ) {
                    style.setBorderLeft( borderLeft );
                    style.setLeftBorderColor( leftBorderColor );
                }
                if ( borderRight != CellStyle.BORDER_NONE ) {
                    style.setBorderRight( borderRight );
                    style.setRightBorderColor( rightBorderColor );
                }
            }
        } );
    }

    
    
    /**
     * Sets the border style of a given sub-block of cells.
     * If provided border type if BORDER_NONE the corresponding border is not modified.
     * The other properties of cell styles are preserved.
     * @see modifyCellStyle() 
     */
    public void putBorder( final short borderTop, final XSSFColor topBorderColor,
                           final short borderBottom, final XSSFColor bottomBorderColor,
                           final short borderLeft, final XSSFColor leftBorderColor,
                           final short borderRight, final XSSFColor rightBorderColor,
                           int[] rowIndices, int[] colIndices )
    {
        modifyCellStyle(rowIndices, colIndices, new CellStyleModifier() {
            public void modify( CellStyle style ) {
                if ( style instanceof XSSFCellStyle ) {
                    XSSFCellStyle xssfStyle = (XSSFCellStyle)style;
                    if ( borderTop != CellStyle.BORDER_NONE ) {
                        xssfStyle.setBorderTop( borderTop );
                        xssfStyle.setTopBorderColor( topBorderColor );
                    }
                    if ( borderBottom != CellStyle.BORDER_NONE ) {
                        xssfStyle.setBorderBottom( borderBottom );
                        xssfStyle.setBottomBorderColor( bottomBorderColor );
                    }
                    if ( borderLeft != CellStyle.BORDER_NONE ) {
                        xssfStyle.setBorderLeft( borderLeft );
                        xssfStyle.setLeftBorderColor( leftBorderColor );
                    }
                    if ( borderRight != CellStyle.BORDER_NONE ) {
                        xssfStyle.setBorderRight( borderRight );
                        xssfStyle.setRightBorderColor( rightBorderColor );
                    }
                } 
                else throw new RuntimeException( 
                		"Current cell style doesn't support XSSF colors");
            }
        } );
    }
    
    
}