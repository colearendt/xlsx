\name{CellStyle}
\alias{CellStyle}
\alias{createCellStyle}
\alias{setCellStyle}
\alias{getCellStyle}
\title{Functions to manipulate cells.}
\description{
Create and set cell styles. 
}
\usage{

CellStyle(wb, hAlign=NULL, vAlign=NULL, borderPosition=NULL,
  borderPen="BORDER_NONE", borderColor=NULL, fillBackgroundColor=NULL,
  fillForegroundColor=NULL, fillPattern=NULL, font=NULL,
  dataFormat=NULL, wrapText=FALSE)

setCellStyle(cell, cellStyle)

getCellStyle(cell)

}
\arguments{
  \item{wb}{a workbook object as returned by
    \code{\link{createWorkbook}} or \code{\link{loadWorkbook}}.}
  
  \item{hAlign}{a character specifiying the horizontal alignment.  Valid
    values are "ALIGN_CENTER", "ALIGN_CENTER_SELECTION", "ALIGN_FILL", 
    "ALIGN_GENERAL", "ALIGN_JUSTIFY", "ALIGN_LEFT", "ALIGN_RIGHT".}
			
  \item{vAlign}{a character specifiying the vertical alignment.  Valid
    values are "VERTICAL_BOTTOM", "VERTICAL_CENTER", "VERTICAL_JUSTIFY", 
    "VERTICAL_TOP".}
		       
  \item{borderPosition}{a character specifiying the border position.
    Valid values are "BOTTOM", "LEFT", "TOP", "RIGHT".  You can pass a
    vector with any combination of these four values, if you want to
    have  multiple borders (e.g. pass all four if you want to box your cell.)}
  
  \item{borderPen}{a character specifiying the border pen.  Valid values
    are "BORDER_DASH_DOT", "BORDER_DASH_DOT_DOT1", "BORDER_DASHED",
    "BORDER_DOTTED", "BORDER_DOUBLE", "BORDER_HAIR", "BORDER_MEDIUM",
    "BORDER_MEDIUM_DASH_DOT1", "BORDER_MEDIUM_DASH_DOT_DOT1", 
    "BORDER_MEDIUM_DASHED", "BORDER_NONE", "BORDER_SLANTED_DASH_DOT1",
    "BORDER_THICK", "BORDER_THIN".}
			
  \item{borderColor}{a character specifiying the border color.  Any
    color names as returned by \code{\link[grDevices]{colors}} can be
    used.}
  
  \item{fillBackgroundColor}{a character specifiying the fill background
    color.  Any color names as returned by
    \code{\link[grDevices]{colors}} can be used.}
  
  \item{fillForegroundColor}{a character specifiying the fill foreground
    color.  Any color names as returned by
    \code{\link[grDevices]{colors}} can be used.}
  
  \item{fillPattern}{a character specifiying the fill pattern.  Valid
    values are "BRICKS1", "DIAMONDS1", "FINE_DOTS", "LEAST_DOTS1",
    "LESS_DOTS1", "NO_FILL", "SOLID_FOREGROUND", "SPARSE_DOTS",
    "SQUARES1", "THICK_BACKWARD_DIAG", "THICK_FORWARD_DIAG",
    "THICK_HORZ_BANDS", "THICK_VERT_BANDS", "THIN_BACKWARD_DIAG1",
    "THIN_FORWARD_DIAG1", "THIN_HORZ_BANDS1", "THIN_VERT_BANDS1".}
  
  \item{font}{a \code{\link{Font}} object.}
  
  \item{dataFormat}{a character specifying the data format.  See
    Details.}

  \item{cell}{a \code{\link{Cell}} object.}

  \item{cellStyle}{a \code{CellStyle} object as returned by
    \code{createCellStyle}.}

  \item{wrapText}{a boolean if the text of the cells should be wrapped
    or not.}
  
}

\value{
  
  \code{createCellStyle} creates a CellStyle object.

  \code{getCellStyle} returns the associated CellStyle object for the
    given Cell.
    
}
\details{
  \code{setCellStyle} sets the \code{CellStyle} to one \code{Cell}
  object.

  Specifying the \code{dataFormat} argument allows you to format the
  cell.  For example, "#,##0.00" corresponds to using a comma separator
  for powers of 1000 with two decimal places, "m/d/yyyy" can be used to
  format dates and is the equivalent of \R's MM/DD/YYYY format.  To
  format datetimes use "m/d/yyyy h:mm:ss;@".  To show negative values in
  red within parantheses with two decimals and commas after power of
  1000 use "#,##0.00_);[Red](#,##0.00)".  I am not aware of an official
  way to discover these strings.  I find them out by recording a macro
  that formats a specific cell and then checking out the resulting VBA
  code.  From there you can read the \code{dataFormat} code.
  
  You need to have a \code{Workbook} object to attach  a
  \code{CellStyle} object to it.

  Since OS X 10.5 Apple dropped support for AWT on the main thread,  
  so essentially you cannot use any graphics classes in R on OS X 10.5  
  since R is single-threaded. (verbatim from Simon Urbanek).  This
  implies that setting colors on Mac will not work as is!  A set of
  about 50 basic colors are still available please see the javadocs. 

  For Excel 95/2000/XP/2003 the choice of colors is limited.  See
  \code{xlsx:::.INDEXED_COLORS} for the list of available colors.  
  
  Unspecified values for arguments are taken from the system locale.  
  
}
\author{Adrian Dragulescu}
\examples{
\dontrun{
  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")

  rows  <- createRow(sheet, rowIndex=1)    

  cell.1 <- createCell(rows, colIndex=1)[[1,1]]     
  setCellValue(cell.1, "Hello R!")

  cellStyle1 <- createCellStyle(wb, borderPosition="RIGHT",
    borderPen="BORDER_DASHED", fillBackgroundColor="yellow",
    fillForegroundColor="tomato", fillPattern="BIG_SPOTS")
  setCellStyle(cell.1, cellStyle1)  

  # you need to save the workbook now if you want to see this art
}
}

