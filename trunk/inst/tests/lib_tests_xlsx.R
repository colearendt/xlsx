# test the package
# 
# test.cellStyles
# test.comments
# test.dataFormats
# test.otherEffects
# test.picture
# test.ranges
#
# .main_highlevel_export
# .main_lowlevel_export
# .main





#####################################################################
# Test Borders, Fonts, Colors, etc. 
# 
test.cellStyles <- function(wb)
{
  cat("Testing cell styles ...\n")
  sheet  <- createSheet(wb, sheetName="cellStyles")
  rows   <- createRow(sheet, rowIndex=1:12)         
  cells  <- createCell(rows, colIndex=1:8)      

  mapply(setCellValue, cells[,1], month.name)

  cat("  Check borders of different colors.\n")
  setCellValue(cells[[2,2]], paste("<-- Thick red bottom border,",
                                   "thin blue top border."))
  borders <- Border(color=c("red","blue"), position=c("BOTTOM", "TOP"),
                    pen=c("BORDER_THICK", "BORDER_THIN"))
  cs1 <- CellStyle(wb) + borders
  setCellStyle(cells[[2,1]], cs1)   

  
  cat("  Check fills.\n")
  setCellValue(cells[[4,2]], "<-- Solid lavender fill.")
  cs2 <- CellStyle(wb) + Fill(backgroundColor="lavender",
    foregroundColor="lavender", pattern="SOLID_FOREGROUND")
  setCellStyle(cells[[4,1]], cs2) 

    
  cat("  Check fonts.\n") 
  setCellValue(cells[[6,2]], "<-- Courier New, Italicised, in orange, size 20 and bold")
  font <- Font(wb, heightInPoints=20, isBold=TRUE, isItalic=TRUE,
    name="Courier New", color="orange")
  cs3 <- CellStyle(wb) + font
  setCellStyle(cells[[6,1]], cs3)   

  
  cat("  Check alignment.\n")
  setCellValue(cells[[8,2]], "<-- Right aligned")
  cs4 <- CellStyle(wb) + Alignment(h="ALIGN_RIGHT")
  setCellStyle(cells[[8,1]], cs4)


  cat("  Check dataFormat.\n")
  setCellValue(cells[[10,1]], -12345.6789)
  setCellValue(cells[[10,2]], "<-- Format -12345.6789 in accounting style.")
  cs5 <- CellStyle(wb) + DataFormat("#,##0.00_);[Red](#,##0.00)")
  setCellStyle(cells[[10,1]], cs5)

  
  cat("  Autosize first, second column.\n") 
  autoSizeColumn(sheet, 1)
  autoSizeColumn(sheet, 2)
  setCellValue(cells[[1,4]], "First and second columns are autosized.")

  cat("Done.\n")
}


#####################################################################
# Test comments
# 
test.comments <- function(wb)
{
  cat("Testing comments ... ")
  sheet <- createSheet(wb, "Comments")
  rows  <- createRow(sheet, rowIndex=1:10)      # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns

  cell1 <- cells[[1,1]]
  setCellValue(cell1, 1)   # add value 1 to cell A1

  createCellComment(cell1, "Cogito", author="Descartes")

  comment <- getCellComment(cell1)
  stopifnot(comment$getAuthor()=="Descartes")
  stopifnot(comment$getString()$toString()=="Cogito")

  cat("Done.\n")
}


#####################################################################
# Test dataFormats
# 
test.dataFormats <- function(wb)
{
  cat("Testing dataFormats ... ")

  # create a test data.frame
  data <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
    date=seq(as.Date("1999-01-01"), by="1 year", length.out=10),
    bool=ifelse(1:10 %% 2, TRUE, FALSE), log=log(1:10),
    rnorm=10000*rnorm(10),
    datetime=seq(as.POSIXct("2011-11-06 00:00:00", tz="GMT"), by="1 hour", length.out=10))

  sheet <- createSheet(wb, "dataFormats")
  rows  <- createRow(sheet, rowIndex=1:10)       # 10 rows
  cells <- createCell(rows, colIndex=1:10)       # 10 columns

  # or do them all by looping over columns
  for (ic in 1:ncol(data))
    mapply(setCellValue, cells[,ic], data[,ic]) 

  setCellValue(cells[[1,10]], 'format "log" column with two decimals')
  cellStyle1 <- CellStyle(wb) + DataFormat("#,##0.00")
  lapply(cells[,6], setCellStyle, cellStyle1)

  setCellValue(cells[[2,10]], 'format date column')
  cellStyle2 <- CellStyle(wb) + DataFormat("m/d/yyyy")
  lapply(cells[,4], setCellStyle, cellStyle2)

  setCellValue(cells[[3,10]], paste('format datetime column (tz=GMT only),',
    'should start from 2011-11-06 00:00:00 with hour increments.'))
  cellStyle3 <- CellStyle(wb) + DataFormat("m/d/yyyy h:mm:ss;@")
  #cellStyle2$getDataFormat()
  lapply(cells[,8], setCellStyle, cellStyle3)

  setCellValue(cells[[4,10]], 
    'format "rnorm" column with two decimals, comma separator, red')
  cellStyle4 <- CellStyle(wb) + DataFormat("#,##0.00_);[Red](#,##0.00)")
  lapply(cells[,7], setCellStyle, cellStyle4)

  cat("Done.\n")
}


#####################################################################
# Test other effects
# 
test.otherEffects <- function(wb)
{
  cat("Testing other effects ... \n")

  sheet1 <- createSheet(wb, "otherEffects1")
  rows   <- createRow(sheet1, 1:10)              # 10 rows
  cells  <- createCell(rows, colIndex=1:8)       # 8 columns

  cat("  merge cells\n")
  setCellValue(cells[[1,1]], "<-- a title that spans 3 columns")
  addMergedRegion(sheet1, 1, 1, 1, 3)

  cat("  set column width\n")
  setColumnWidth(sheet1, 1, 25)
  setCellValue(cells[[5,1]], paste("<-- the width of this column",
    "is 20 characters wide."))
  
  cat("  set zoom\n")
  setCellValue(cells[[3,1]], "<-- the zoom on this sheet is 2:1.")
  setZoom(sheet1, 200, 100)

  sheet2 <- createSheet(wb, "otherEffects2")
  rows  <- createRow(sheet2, 1:10)              # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns
  #createFreezePane(sheet2, 1, 1, 1, 1)
  createFreezePane(sheet2, 5, 5, 8, 8)
  setCellValue(cells[[3,3]], "<-- a freeze pane")

  cat("  add hyperlinks to a cell\n")
  address <- "http://poi.apache.org/"
  setCellValue(cells[[1,1]], "click me!")  
  addHyperlink(cells[[1,1]], address)
  
  sheet3 <- createSheet(wb, "otherEffects3")
  rows  <- createRow(sheet3, 1:10)              # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns
  createSplitPane(sheet3, 2000, 2000, 1, 1, "PANE_LOWER_LEFT")
  setCellValue(cells[[3,3]], "<-- a split pane")
  
  cat("Done.\n")
}

  
#####################################################################
# Test pictures
# 
test.picture <- function(wb)
{
  cat("Test embedding an R picture ...\n")

  cat("Add log_plot.jpeg to a new xlsx...")
  picname <- system.file("tests", "log_plot.jpeg", package="xlsx")
  sheet <- createSheet(wb, "picture")

  addPicture(picname, sheet)

  xlsx:::.write_block(wb, sheet, iris)
  cat("Done.\n")  
}

  
#####################################################################
# Test Ranges
# 
test.ranges <- function(wb)
{
  cat("Testing ranges ... ")
  sheets <- getSheets(wb)
  sheet <- sheets[["dataFormats"]]
  
  cat("  make a new range")
  firstCell <- sheet$getRow(2L)$getCell(2L)
  lastCell  <- sheet$getRow(6L)$getCell(5L)
  rangeName <- "Test2"
  createRange(rangeName, firstCell, lastCell)
  
  ranges <- getRanges(wb)
  range <- ranges[[1]]
  res <- readRange(range, sheet, colClasses="numeric")

  cat("Done.\n")
}


#####################################################################
# Test imports
# 
.main_highlevel_import <- function(ext="xlsx", dev=TRUE)
{
  fname <- paste("test_import.", ext, sep="")
  if (dev) {
    file <- paste(SOURCEDIR, "rexcel/trunk/inst/tests/", fname, sep="")
  } else {
    file <- system.file("tests", fname, package = "xlsx")
  }

  cat("Testing high level import read.xls\n")
  cat("  read data from mixedTypes\n")
  orig <- getOption("stringsAsFactors")
  options(stringsAsFactors=FALSE)
  res <- read.xlsx(file, sheetName="mixedTypes")
  stopifnot(class(res[,1])=="Date")
  stopifnot(class(res[,2])=="character")
  stopifnot(class(res[,3])=="numeric")
  stopifnot(class(res[,4])=="logical")
  stopifnot(inherits(res[,6], "POSIXct"))
  options(stringsAsFactors=orig)

  cat("  import keeping formulas\n")
  res <- read.xlsx(file, sheetName="mixedTypes", keepFormulas=TRUE)
  stopifnot(res$Double[4]=="SQRT(2)")
  
  cat("  import with colClasses\n")
  cat("  force conversion of boolean column to numeric\n")
  colClasses <- rep(NA, length=6); colClasses[4] <- "numeric"
  res <- read.xlsx(file, sheetName="mixedTypes", colClasses=colClasses)
  stopifnot(class(res[,4])=="numeric")

  cat("Test you can read sheet oneColumn\n")
  res <- read.xlsx(file, sheetName="oneColumn", keepFormulas=TRUE)
  stopifnot(ncol(res)==1)

  cat("Test you can read string formulas ... \n")
  res <- read.xlsx(file, "formulas", keepFormulas=FALSE)
  stopifnot(res[1,3]=="2010-1") 

  cat("Test you can read #N/A's ... \n")
  res <- read.xlsx(file, "NAs")
  stopifnot(all.equal(which(is.na(res)), c(6,28,29,59,69))) 
  
}

#####################################################################
# Test highlevel export
# 
.main_highlevel_export <- function(ext="xlsx")
{
  cat("Testing high level export ... \n")  
  x <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
    date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
    bool=ifelse(1:10 %% 2, TRUE, FALSE))

  file <- paste(OUTDIR, "test_highlevel_export.", ext, sep="")
  cat("  write an xlsx file with char, int, double, date, bool columns ...\n")
  write.xlsx(x, file)

  cat("  test the append argument by adding another sheet ... \n")
  file <- paste(OUTDIR, "test_highlevel_export.", ext, sep="")
  write.xlsx(USArrests, file, sheetName="usarrests", append=TRUE)
  cat("Wrote file ", file, "\n\n")

  cat("  test writing/reading data.frames with NA values ... \n") 
  file <- paste(OUTDIR, "test_writeread_NA.", ext, sep="")
  x <- data.frame(matrix(c(1.0, 2.0, 3.0, NA), 2, 2))
  write.xlsx(x, file, row.names=FALSE)
  xx <- read.xlsx(file, 1)
  if (!identical(x,xx)) 
    stop("Fix me!")

  cat("Done.\n")
}


#####################################################################
#
.main_lowlevel_export <- function(ext="xlsx")
{
  outfile <- paste(OUTDIR, "test_export.", ext, sep="")
  if (file.exists(outfile)) unlink(outfile)
   
  wb <- createWorkbook(type=ext)

  test.cellStyles(wb)
  test.comments(wb)
  test.dataFormats(wb)
  test.ranges(wb)
  test.otherEffects(wb)
  test.picture(wb)
  
  saveWorkbook(wb, outfile)
  cat("Wrote file", outfile, "\n\n")
}

#####################################################################
# Speed Test export
# Ubuntu desktop, 85s & 3s.
#
.main_speedtest_export <- function(ext="xlsx")
{
  cat("Speed test export ... \n")  

  file <- paste(OUTDIR, "test_exportSpeed.", ext, sep="")
  x <- expand.grid(ind=1:60, letters=letters, months=month.abb)
  x <- cbind(x, val=runif(nrow(x)))
  cat("  writing a data.frame with dim", nrow(x), "x", ncol(x), "\n")
  cat("  timing write.xlsx:", system.time(write.xlsx(x, file)), "\n")   # 99s
  cat("  timing write.xlsx2:", system.time(write.xlsx2(x, file)), "\n") #  9s
  cat("  wrote file ", file, "\n")
  
  cat("Done.\n")
}



#####################################################################
#####################################################################
#
.main <- function()
{
  # best viewed with M-x hs-minor-mode

  require(xlsx)

  if (.Platform$OS.type == "windows") {
    SOURCEDIR <- "C:/google/"
    OUTDIR <<- "C:/temp/"
  } else {
    SOURCEDIR <- "/home/adrian/Documents/"
    OUTDIR <<- "/tmp/"
  }
  thisFile <- paste(SOURCEDIR, "rexcel/trunk/inst/tests/",
    "lib_tests_xlsx.R", sep="")
  source(thisFile)

  
  .main_lowlevel_export(ext="xlsx")  
  .main_highlevel_export(ext="xlsx")
  .main_speedtest_export(ext="xlsx")
  
  .main_lowlevel_export(ext="xls")  
  .main_highlevel_export(ext="xls")
  .main_speedtest_export(ext="xls")

  
  .main_highlevel_import(ext="xlsx")
  #.main_lowlevel_import(ext="xlsx")

  
  .main_import(ext="xls")

  
}



##   cat("Test memory ...\n")
##   file <- paste(outdir, "test_exportMemory.xlsx", sep="")
##   x <- expand.grid(ind=1:1000, letters=letters, months=month.abb)
##   cat("Writing object size: ", object.size(x), " uses all Java heap space\n")
##   (time <- system.time(write.xlsx2(x, file)))
##   cat("Wrote file ", file, "\n\n")






## #####################################################################
## # Test basic Java POI functionality - a bit redundant now
## # 
## test.basicJavaPOI <- function(wb)
## {
##   cat("Create an empty workbook ...\n") 
##   wb <- .jnew("org/apache/poi/xssf/usermodel/XSSFWorkbook")
##   #if (class(wb)=="jobjRef") cat("OK.\n")
##   #.jmethods(wb)

##   cat("Create a sheet called 'Sheet2' ...\n")
##   sheet2 <- .jcall(wb, "Lorg/apache/poi/ss/usermodel/Sheet;",
##     "createSheet", "Sheet2")
##   #if (.jinstanceof(sheet2, "org.apache.poi.ss.usermodel.Sheet"))
##   #  cat("OK.\n")  
##   #.jmethods(sheet2)

##   cat("Create row 1 ...\n")
##   row <- .jcall(sheet2, "Lorg/apache/poi/ss/usermodel/Row;",
##     "createRow", as.integer(0))
##   #.jmethods(row)

##   cat("Create cell [1,1] ...\n")
##   cell <- .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
##     "createCell", as.integer(0))

##   cat("Put a value in cell [1,1] ... ")
##   cell$setCellValue(1.23)

##   cat("Add cell [1,2] and put a numeric value ... \n")
##   .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
##     "createCell", as.integer(1))$setCellValue(3.1415)

##   cat("Add cell [1,3] and put a stringvalue ... \n")
##   .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
##     "createCell", as.integer(2))$setCellValue("A string")

##   cat("Add cell [1,4] and put a boolean value ... \n")
##   .jcall(row, "Lorg/apache/poi/ss/usermodel/Cell;",
##     "createCell", as.integer(3))$setCellValue(TRUE)

##   cat("Done.\n")
## }









