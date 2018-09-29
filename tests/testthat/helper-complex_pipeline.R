
test_complex_read.xlsx <- function(type="xlsx") {
  filename <- paste("test_import.", type, sep="")
  file <- test_ref(filename)
  
  ##Load test_import
  wb <- loadWorkbook(file)
  
  ## getSheets
  sheets <- getSheets(wb)
  
  ## Get the second sheet with mixedTypes ...
  sheet  <- sheets[[2]]
  rows   <- getRows(sheet)
  cells  <- getCells(rows)
  values <- lapply(cells, getCellValue)
  
  ## extract cell [5,2] and see if == 'Apr'
  expect_equal(values[["5.2"]],"Apr")
  
  orig <- getOption("stringsAsFactors")
  options(stringsAsFactors=FALSE)
  
  ##Read data in second sheet ...
  res <- read.xlsx(file, 2)
  
  if (type=='xls') {
    typelist <- list('Date','character','numeric','logical','character','POSIXct')
  } else {
    typelist <- list('Date','character','numeric','logical','numeric','POSIXct')
  }
    
  
  invisible(mapply(expect_is
         ,res
         ,typelist
         ))
  
  options(stringsAsFactors=orig)
  
  ## Some cells are errors because of wrong formulas
  if (type=='xls')
    expect_match(res$Double[7:8],'java\\.lang\\.IllegalStateException')
  
  ## Test high level import keeping formulas...
  res <- read.xlsx(file, 2, keepFormulas=TRUE)
  
  expect_equal(as.character(res$Double),
               c('-1.123','1.123','10^1.5','SQRT(2)','LOG(2)'
                 ,'10^210+1.12','1/0','-1/0','E2+1','E10+1'
                 ,'E11+1','E12+1')
  )
  
  
  #Force conversion of 5th column to numeric.
  colClasses <- rep(NA, length=6); colClasses[5] <- "numeric"
  res <- read.xlsx(file, 2, colClasses=colClasses)
  expect_is(res[,5],'numeric')
  
  ##Test you can import sheet oneColumn... 
  res <- read.xlsx(file, "oneColumn", keepFormulas=TRUE)
  expect_equal(ncol(res),1)
  
  ##Check that you can import String formulas ...
  res <- read.xlsx(file, "formulas", keepFormulas=FALSE)
  expect_equal(as.character(res$date),c('2010-1','2010-2','2010-3'))
  
  ## Test you can read #N/A's ... 
  res <- read.xlsx(file, "NAs")
  expect_equal(which(is.na(res)), c(6,28,29,59,69))
  
  ## read ragged data
  res  <- read.xlsx(file, sheetName="ragged", colIndex=1:4)  
  
  ## read rowIndex with read.xlsx
  # reported bug in 0.5.1, fixed on 6/20/2013
  res <- read.xlsx(file, sheetName="all", colIndex=3:6, rowIndex=3:7)
  expect_equal(res,data.frame(mon=c('Jan','Feb','Mar','Apr')
                              , day=c(1,2,3,4)
                              , year=c(2000,2001,2002,2003)
                              , date=as.Date(c('1999-01-01','2000-01-01','2001-01-01','2002-01-01'))
                              )
               )
}

test_complex_read.xlsx2 <- function(type='xlsx') {
  filename <- paste0("test_import.", type)
  file <- test_ref(filename)
  
  res <- read.xlsx2(file, sheetName="mixedTypes", stringsAsFactors=FALSE)
  expect_equal(as.character(lapply(res,typeof)),rep('character',6))
  
  res <- read.xlsx2(file, sheetName="mixedTypes", colClasses=c(
    "numeric", "character", rep("numeric", 4))
    , stringsAsFactors=FALSE)
  expect_equal(as.character(lapply(res,typeof)),c('double','character',rep('double',4)))
  
  
  res <- read.xlsx2(file, sheetName="mixedTypes", startRow=2, endRow=4
                    , stringsAsFactors=FALSE
                    , header = FALSE)
  expect_equal(nrow(res),3)
  
  
  res <- read.xlsx2(file, sheetName="all", startRow=3, stringsAsFactors=FALSE)
  expect_equal(nrow(res),10)
  expect_equal(ncol(res),8)
  
  ## read more columns than on the spreadsheet
  res <- read.xlsx2(file
                    , sheetName="all"
                    , header=FALSE
                    , startRow=3
                    , endRow=6
                    , colIndex=3:14
                    , stringsAsFactors=FALSE)
  expect_equal(nrow(res),4)
  expect_equal(ncol(res),12)
  
  ## pass in some colClasses
  res <- read.xlsx2(file, sheetName="all", startRow=3, colIndex=3:10,
                    colClasses=c("character", rep("numeric", 2), "Date", "character",
                                 "numeric", "numeric", "POSIXct"))
  expect_is(res[,4],'Date')
  expect_is(res[,8],'POSIXct')
  
  ## read non contiguos blocks
  res <- read.xlsx2(file, sheetName="all", startRow=3,
                    colIndex=c(3,4,6,8,9,10))
  expect_equal(ncol(res), 6)
  expect_equal(nrow(res),10)
  
  ## read ragged data
  res <- read.xlsx2(file, sheetName="ragged", stringsAsFactors=FALSE)
  
  res_tmp <- data.frame(Field1=c('','A1','','','')
                    ,Field2=c('B2','','B4','','')
                    ,Field3=c('','','','C5','')
                    ,Value=c('','7','8','9','')
                    , stringsAsFactors=FALSE)
  expect_identical(res,res_tmp)
}

test_complex_cell <- function(type='xlsx') {
  filename <- paste("test_import.", type, sep="")
  file <- test_ref(filename)
  
  wb <- loadWorkbook(file)
  sheets <- getSheets(wb)
  sheet  <- sheets[['deletedFields']]
  
  ## extract only some rows (say 5) ...
  rows   <- getRows(sheet, rowIndex=1:5)
  cells  <- getCells(rows)
  res <- lapply(cells, getCellValue)
  rr <- unique(sapply(strsplit(names(res), "\\."), "[[", 1))
  expect_identical(rr, c("1","2","3","4", "5"))
  
  ## extract only some columns (say 4) ...
  rows   <- getRows(sheet)
  cells  <- getCells(rows, colIndex=1:4)
  res <- lapply(cells, getCellValue)
  cols <- unique(sapply(strsplit(names(res), "\\."), "[[", 2))
  expect_identical(cols, c('1','2','3','4'))
}

test_basic_import <- function(type="xlsx") {
  ## Testing low level import ...
  fname  <- paste0("test_import.", type)
  file   <- test_ref(fname)
  wb     <- loadWorkbook(file)
  sheets <- getSheets(wb)
  
  ## readColumns on all sheet
  sheet <- sheets[["all"]]
  res <- readColumns(sheet, startColumn=3, endColumn=10, startRow=3,
                     endRow=7)
  expect_equal(nrow(res),4)
  expect_equal(ncol(res),8)
  
  ## readColumns for formulas and NAs
  sheet <- sheets[["NAs"]]
  res <- readColumns(sheet, 1, 6, 1,  colClasses=c("Date", "character",
                                                   "integer", rep("numeric", 2),  "POSIXct"))
  expect_equal(which(is.nan(res[,5])),c(7,8,11,12))
  
  ## readColumns for ragged sheets
  sheet <- sheets[["ragged"]]
  res <- readColumns(sheet, 1, 4, 1
                     ,  colClasses=c(rep("character", 3),"numeric")
                     , stringsAsFactors=FALSE)
  expect_equal(res[1,1],'')
  
  ## readRows
  sheet <- sheets[["all"]]
  res <- readRows(sheet, startRow=3, endRow=7, startColumn=2, endColumn=15)
  
  expect_equal(ncol(res),14)
  expect_equal(nrow(res),5)
}

test_basic_export  <- function(type="xlsx") {
  outfile <- test_tmp(paste0("test_export.", type))
  
  wb <- createWorkbook(type=type)
  
  wrapper_cellStyle(wb)
  wrapper_comment(wb)
  wrapper_dataFormat(wb)
  wrapper_range(wb)
  wrapper_otherEffect(wb)
  wrapper_picture(wb)
  wrapper_addDataFrame(wb)
  wrapper_cellBlock(wb)
  
  saveWorkbook(wb, outfile)
}

test_complex_export <- function(type="xlsx") {
  
  ##Testing high level export ...  
  x <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
                  date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
                  bool=ifelse(1:10 %% 2, TRUE, FALSE))
  
  file <- test_tmp(paste0("test_highlevel_export.", type))
  ## write an xlsx file with char, int, double, date, bool columns ...
  write.xlsx(x, file, sheetName="writexlsx")
  
  write.xlsx2(x, file, sheetName="writexlsx2", append=TRUE, row.names=FALSE) 
  
  ## test the append argument by adding another sheet ...
  write.xlsx(USArrests, file, sheetName="usarrests", append=TRUE)
  
  ## test writing/reading data.frames with NA values ...
  file <- test_tmp(paste0("test_writeread_NA.", type))
  
  x <- data.frame(matrix(c(1.0, 2.0, 3.0, NA), 2, 2))
  write.xlsx(x, file, row.names=FALSE)
  xx <- read.xlsx(file, 1)
  expect_identical(x,xx)
}

test_complex_cellBlock <- function(type='xlsx') {
  outfile <- test_tmp(paste0("test_cellBlock.", type))
  
  wb <- createWorkbook(type=type)
  
  sheet  <- createSheet(wb, sheetName="CellBlock")
  
  ## Add a cell block to sheet CellBlock
  cb <- CellBlock(sheet, 7, 3, 50, 60)
  CB.setColData(cb, 1:50, 1)    # set a column
  CB.setRowData(cb, 1:50, 1)     # set a row
  
  # add a matrix, and style it
  cs <- CellStyle(wb) + DataFormat("#,##0.00")
  x  <- matrix(rnorm(40*45), nrow=40)
  CB.setMatrixData(cb, x, 10, 4, cellStyle=cs)  
  
  # highlight the negative numbers in red 
  fill <- Fill(foregroundColor = "red", backgroundColor="red")
  ind  <- which(x < 0, arr.ind=TRUE)
  CB.setFill(cb, fill, ind[,1]+9, ind[,2]+3)  # note the indices offset
  
  # set the border on the top row of the Cell Block
  border <-  Border(color="blue", position=c("TOP", "BOTTOM"),
                    pen=c("BORDER_THIN", "BORDER_THICK"))
  CB.setBorder(cb, border, 1:50, 1)
  
  saveWorkbook(wb, outfile)  
} 

test_basicFunctions <- function(type='xlsx') {
  ## Testing basic workbook functions
  ## Create an empty workbook ...
  wb <- createWorkbook(type=type)
  
  ## create a sheet called 'Sheet1'
  sheet1 <- createSheet(wb, sheetName="Sheet1")
  
  ## create another sheet called 'Sheet2'
  sheet2 <- createSheet(wb, sheetName="Sheet2")
  
  ## get sheets
  sheets <- getSheets(wb)
  expect_equal(length(sheets),2)
  
  ## remove sheet named 'Sheet2'
  removeSheet(wb, sheetName="Sheet2")
  sheets <- getSheets(wb)  
  expect_equal(length(sheets),1)
  
  ##  add rows 6:10 on Sheet1
  rows <- createRow(sheet1, 6:10)
  expect_equal(length(rows), 5)
  
  ##  remove rows 1:10 on Sheet1 of test_import.xlsx
  filename <- paste0("test_import.", type)
  file <- test_ref(filename)
  wb <- loadWorkbook(file)  
  sheets <- getSheets(wb)
  sheet <- sheets[[1]]  
  rows  <- getRows(sheet)           # get all the rows
  removeRow(sheet, rows[1:10])
  rows  <- getRows(sheet)           # get all the rows
  expect_equal(length(rows),41)
}

test_addOnExistingWorkbook <- function(ext="xlsx") {
  ##Test adding a df to an existing workbook
  
  src <- test_ref(paste0("test_import.", ext))
  wb  <- loadWorkbook( src )
  sheets <- getSheets( wb )
  
  dat <- data.frame(a=LETTERS, b=1:26)
  
  s <- sheets$mixedTypes
  addDataFrame(dat, s, startColumn=20, startRow=5)
  
  expect_equal(
    getCellValue(
      getCells(getRows(s,25),colIndex=21)[[1]]
      )
    , 'T'
  )
}
