#' Functions to manipulate cells.
#' 
#' Functions to manipulate cells.
#' 
#' \code{setCellValue} writes the content of an R variable into the cell.
#' \code{Date} and \code{POSIXct} objects are passed in as numerical values.
#' To format them as dates in Excel see \code{\link{CellStyle}}.
#' 
#' These functions are not vectorized and should be used only for small
#' spreadsheets.  Use \code{CellBlock} functionality to efficiently read/write
#' parts of a spreadsheet.
#' 
#' @param row a list of row objects. See \code{Row}.
#' @param colIndex a numeric vector specifying the index of columns.
#' @param simplify a logical value.  If \code{TRUE}, the result will be
#' unlisted.
#' @param value an R variable of length one.
#' @param richTextString a logical value indicating if the value should be
#' inserted into the Excel cell as rich text.
#' @param showNA a logical value.  If \code{TRUE} the cell will contain the
#' "#N/A" value, if \code{FALSE} they will be skipped.  The default value was
#' chosen to remain compatible with previous versions of the function.
#' @param keepFormulas a logical value.  If \code{TRUE} the formulas will be
#' returned as characters instead of being explicitly evaluated.
#' @param encoding A character value to set the encoding, for example "UTF-8".
#' @param cell a \code{Cell} object.
#' @return
#' 
#' \code{createCell} creates a matrix of lists, each element of the list being
#' a java object reference to an object of type Cell representing an empty
#' cell.  The dimnames of this matrix are taken from the names of the rows and
#' the \code{colIndex} variable.
#' 
#' \code{getCells} returns a list of java object references for all the cells
#' in the row if \code{colIndex} is \code{NULL}.  If you want to extract only a
#' specific columns, set \code{colIndex} to the column indices you are
#' interested.
#' 
#' \code{getCellValue} returns the value in the cell as an R object.  Type
#' conversions are done behind the scene.  This function is not vectorized.
#' @author Adrian Dragulescu
#' @seealso To format cells, see \code{\link{CellStyle}}.  For rows see
#' \code{\link{Row}}, for sheets see \code{\link{Sheet}}.
#' @examples
#' 
#' 
#' file <- system.file("tests", "test_import.xlsx", package = "xlsx")
#' 
#' wb <- loadWorkbook(file)  
#' sheets <- getSheets(wb)
#' 
#' sheet <- sheets[['mixedTypes']]      # get second sheet
#' rows  <- getRows(sheet)   # get all the rows
#' 
#' cells <- getCells(rows)   # returns all non empty cells
#' 
#' values <- lapply(cells, getCellValue) # extract the values
#' 
#' # write the months of the year in the first column of the spreadsheet
#' ind <- paste(2:13, ".2", sep="")
#' mapply(setCellValue, cells[ind], month.name)
#' 
#' ####################################################################
#' # make a new workbook with one sheet and 5x5 cells
#' wb <- createWorkbook()
#' sheet <- createSheet(wb, "Sheet1")
#' rows  <- createRow(sheet, rowIndex=1:5)
#' cells <- createCell(rows, colIndex=1:5) 
#' 
#' # populate the first column with Dates
#' days <- seq(as.Date("2013-01-01"), by="1 day", length.out=5)
#' mapply(setCellValue, cells[,1], days)
#' 
#'  
#' 
#' @name Cell
NULL

######################################################################
# Create some cells in the row object.  "cell" is the index of columns. 
# You can pass in a list of rows.
# Return a matrix of lists.  Each element is a cell object.
#' @export
#' @rdname Cell
createCell <- function(row, colIndex=1:5)
{
  cells <- matrix(list(), nrow=length(row), ncol=length(colIndex),
    dimnames=list(names(row), colIndex))
  
  for (ir in seq_along(row))
    for (ic in seq_along(colIndex))
      cells[[ir,ic]] <- .jcall(row[[ir]], "Lorg/apache/poi/ss/usermodel/Cell;",
        "createCell", as.integer(colIndex[ic]-1))
    
  cells
}

######################################################################
# Get the cells for a list of rows.  Users who want basic things only
# don't need to use this function. 
# 
#' @export
#' @rdname Cell
getCells <- function(row, colIndex=NULL, simplify=TRUE)
{
  nC  <- length(colIndex)
  if (!is.null(colIndex))
    colIx <- as.integer(colIndex-1)     # ugly, have to do it here
 
  res <- row
  for (ir in seq_along(row)){
    if (is.null(colIndex)){                           # get all columns
      minColIx <- .jcall(row[[ir]], "T", "getFirstCellNum")   # 0-based
      maxColIx <- .jcall(row[[ir]], "T", "getLastCellNum")-1  # 0-based
      # actual col index; if the row is empty do nothing
      colIx <- if (minColIx < 0) numeric(0) else seq.int(minColIx, maxColIx) 
    }
    nC <- length(colIx)
    rowCells <- vector("list", length=nC)
    namesCells <- vector("character", length=nC)
    for (ic in seq_along(rowCells)){
      aux <- .jcall(row[[ir]], "Lorg/apache/poi/ss/usermodel/Cell;",
        "getCell", colIx[ic])
      if (!is.null(aux)){
        rowCells[[ic]] <- aux
        namesCells[ic] <- .jcall(aux, "I", "getColumnIndex")+1
      }
    }
    names(rowCells) <- namesCells  # need namesCells if spreadsheet is ragged
    res[[ir]] <- rowCells
  }

  if (simplify)
    res <- unlist(res)
  
  res
}


######################################################################
# Only one cell and one value.  
# You vectorize outside this function if you want.
# 
#    Date    = .jnew("java/text/SimpleDateFormat",
#      "yyyy-MM-dd")$parse(as.character(value)),     # does not format it!
#
#' @export
#' @rdname Cell
setCellValue <- function(cell, value, richTextString=FALSE, showNA=TRUE)
{
  if (is.na(value)) {
    if (showNA) {
      return(invisible(.jcall(cell, "V", "setCellErrorValue", .jbyte(42))))
    } else { return(invisible()) }
  }
  
  value <- switch(class(value)[1],
    integer = as.numeric(value),
    numeric = value,              
    Date    = as.numeric(value) + 25569,             # add Excel origin
    POSIXct = as.numeric(value)/86400 + 25569,               
    as.character(value))  # for factors and other types

  if (richTextString)
    value <- .jnew("org/apache/poi/sf/usermodel/RichTextString",
      as.character(value))  # do I need to convert to as.character ?!!
  
  invisible(.jcall(cell, "V", "setCellValue", value))
}


######################################################################
# get cell value. ONE cell only
# Not happy with the case when you have formulas.  Still not general
#   enough.  We'll see how many things still not work.
#
#' @export
#' @rdname Cell
getCellValue <- function(cell, keepFormulas=FALSE, encoding="unknown")
{
  cellType <- .jcall(cell, "I", "getCellType") + 1
  value <- switch(cellType,
    .jcall(cell, "D", "getNumericCellValue"),        # 1 = numeric
                  
    {strVal <- .jcall(.jcall(cell,                   # 2 = string
      "Lorg/apache/poi/ss/usermodel/RichTextString;",
      "getRichStringCellValue"), "S", "toString");
     if (encoding=="unknown") {strVal} else {Encoding(strVal) <- encoding; strVal}
   },  
                  
    ifelse(keepFormulas, .jcall(cell, "S", "getCellFormula"),   # if a formula
      tryCatch(.jcall(cell, "D", "getNumericCellValue"),        # try to extract  
        error=function(e){                                      # contents  
          tryCatch(.jcall(cell, "S", "getStringCellValue"),     
            error=function(e){
              tryCatch(.jcall(cell, "Z", "getBooleanCellValue"),
                error=function(e)e,
                finally=NA)
            }, finally=NA)
        }, finally=NA)
    ),
                  
    NA,                                              # blank cell
                  
    .jcall(cell, "Z", "getBooleanCellValue"),        # boolean
                  
    NA, #ifelse(keepErrors, .jcall(cell, "B", "getErrorCellValue"), NA), # error
                  
    "Error"                                          # catch all
  ) 

  value
}
