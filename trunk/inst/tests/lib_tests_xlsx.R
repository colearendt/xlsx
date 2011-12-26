# test the package
# 
# test.cellStyles
# test.comments


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
  font <- Font(wb, fontHeightInPoints=20, isBold=TRUE, isItalic=TRUE,
    fontName="Courier New", color="orange")
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
    date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
    bool=ifelse(1:10 %% 2, TRUE, FALSE), log=log(1:10),
    rnorm=10000*rnorm(10),
    datetime=seq(as.POSIXct("2011-11-06 00:00:00"), by="1 hour", length.out=10))

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

  setCellValue(cells[[3,10]], 'format datetime column, around DST')
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
# Test Ranges
# 
test.ranges <- function()
{
  cat("Testing ranges ... ")
  filename <- paste("test_import.", type, sep="")

  cat("  get named ranges from test_import.xlsx ")
  file <- system.file("tests", filename, package = "xlsx")
  wb <- loadWorkbook(file)
  sheets <- getSheets(wb)
  sheet <- sheets[["deletedFields"]]
  
  ranges <- getRanges(wb)
  range <- ranges[[1]]
  res <- readRange(range, sheet, colClasses="numeric")

  cat("  make a new range")
  firstCell <- sheet$getRow(14L)$getCell(4L)
  lastCell  <- sheet$getRow(20L)$getCell(7L)
  rangeName <- "Test2"
  createRange(rangeName, firstCell, lastCell)
  
  if (length(getRanges(wb)) != 2){
    cat("STOP!!! Range not created!")
  } else {
    cat("OK\n")
  }
  
}


#####################################################################
#####################################################################
#
.main <- function()
{
  # best viewed with M-x hs-minor-mode

  require(xlsx)
  
  DIR <- "H:/user/R/Adrian/"
  DIR <- "/home/adrian/Documents/"
  thisFile <- paste(DIR, "findataweb/temp/xlsx/trunk/inst/tests/",
    "lib_tests_xlsx.R", sep="")
  source(thisFile)
 
  outfile <- "/tmp/test_export.xlsx"
  if (file.exists(outfile)) unlink(outfile)
   
  wb <- createWorkbook(type=gsub(".*\\.(.*)", "\\1", outfile))

  test.cellStyles(wb)
  test.comments(wb)
  test.dataFormats(wb)
  test.ranges()
  
  saveWorkbook(wb, outfile)
  cat("Wrote file", outfile, "\n\n")

  

}









