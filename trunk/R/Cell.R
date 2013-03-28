######################################################################
# Create some cells in the row object.  "cell" is the index of columns. 
# You can pass in a list of rows.
# Return a matrix of lists.  Each element is a cell object.
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
      colIx    <- seq.int(minColIx, maxColIx)        # actual col index
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






















