# One sheet extraction.  Similar to read.csv. 
#
#
read.xlsx <- function(file, sheetIndex, sheetName=NULL,
  rowIndex=NULL, startRow=NULL, endRow=NULL, colIndex=NULL,
  as.data.frame=TRUE, header=TRUE, colClasses=NA,
  keepFormulas=FALSE, encoding="unknown", ...)
{
  if (is.null(sheetName) & missing(sheetIndex))
    stop("Please provide a sheet name OR a sheet index.")

  wb <- loadWorkbook(file)
  sheets <- getSheets(wb)
  sheet  <- if (is.null(sheetName)) {
    sheets[[sheetIndex]]
  } else {
    sheets[[sheetName]]
  }
  
  if (is.null(sheet))
    stop("Cannot find the sheet you requested in the file!")

  rowIndex <- if (is.null(rowIndex)) {
    if (is.null(startRow))
      startRow <- .jcall(sheet, "I", "getFirstRowNum") + 1
    if (is.null(endRow)) 
      endRow <- .jcall(sheet, "I", "getLastRowNum") + 1
    startRow:endRow
  } else rowIndex
  
  rows  <- getRows(sheet, rowIndex)
  if (length(rows)==0)
      return(NULL)             # exit early
  
  cells <- getCells(rows, colIndex)
  res <- lapply(cells, getCellValue, keepFormulas=keepFormulas,
                encoding=encoding)

  if (as.data.frame) {
    # need to use the index from the names because of empty cells
    ind <- lapply(strsplit(names(res), "\\."), as.numeric) 
    namesIndM <- do.call(rbind, ind)
    
    row.names <- sort(as.numeric(unique(namesIndM[,1])))
    col.names <- paste("V", sort(unique(namesIndM[,2])), sep="")  
    col.names <- sort(unique(namesIndM[,2]))
    cols <- length(col.names)

    VV <- matrix(list(NA), nrow=length(row.names), ncol=cols,
      dimnames=list(row.names, col.names))
    # you need indM for empty rows/columns when indM != namesIndM
    indM <- apply(namesIndM, 2, function(x){as.numeric(as.factor(x))})
    VV[indM] <- res 
    
    if (header){  # first row of cells that you want
      colnames(VV) <- VV[1,]
      VV <- VV[-1,,drop=FALSE]
    }
    
    res <- vector("list", length=cols)
    names(res) <- colnames(VV)
    for (ic in seq_len(cols)) {
      aux   <- unlist(VV[,ic], use.names=FALSE)
      nonNA <- which(!is.na(aux)) 
      if (length(nonNA)>0) {  # not a NA column in the middle of data
        ind <- min(nonNA)
        if (class(aux[ind])=="numeric") {
          # test first not NA cell if it's a date/datetime
          dateUtil <- .jnew("org/apache/poi/ss/usermodel/DateUtil")
          cell <- cells[[paste(row.names[ind + header], ".", col.names[ic], sep = "")]]
          isDatetime <- dateUtil$isCellDateFormatted(cell)
          if (isDatetime){
            if (identical(aux, round(aux))){ # you have dates
              aux <- as.Date(aux-25569, origin="1970-01-01") 
            } else {  # Excel does not know timezones?!
              aux <- as.POSIXct((aux-25569)*86400, tz="GMT",
                                origin="1970-01-01")
            }
          }
        }
      }
      if (!is.na(colClasses[ic]))
        suppressWarnings(class(aux) <- colClasses[ic])  # if it gets specified
      res[[ic]] <- aux
    }

    res <- data.frame(res, ...)
  }

  res
}



