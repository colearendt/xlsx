wrapper_addDataFrame <- function(wb) {
  #Testing addDataFrame ...

  #custom styles
  sheet1 <- createSheet(wb, sheetName="addDataFrame1")
  data0 <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
                      date=seq(as.Date("1999-01-01"), by="1 year", length.out=10),
                      bool=ifelse(1:10 %% 2, TRUE, FALSE), log=log(1:10),
                      rnorm=10000*rnorm(10),
                      datetime=seq(as.POSIXct("2011-11-06 00:00:00", tz="GMT"), by="1 hour",
                                   length.out=10))
  cs1 <- CellStyle(wb) + Font(wb, isItalic=TRUE)
  cs2 <- CellStyle(wb) + Font(wb, color="blue")
  cs3 <- CellStyle(wb) + Font(wb, isBold=TRUE) + Border()
  addDataFrame(data0, sheet1, startRow=3, startColumn=2, colnamesStyle=cs3,
               rownamesStyle=cs1, colStyle=list(`2`=cs2, `3`=cs2))

  # NA treatment, with defaults
  sheet2 <- createSheet(wb, sheetName="addDataFrame2")
  data <- data.frame(mon=month.abb
                     , index=1:12
                     , double=seq(1.23, by=1,length.out=12)
                     , stringsAsFactors=FALSE)
  data$mon[3:4] <- NA
  data$mon[12] <- "defaults, showNA=FALSE"
  data$index[c(1, 7, 11)] <- NA
  data$double[3:4] <- NA
  addDataFrame(data, sheet2, row.names=FALSE)

  # NA treatment, with showNA=TRUE
  sheet3 <- createSheet(wb, sheetName="addDataFrame3")
  data$mon[12] <- "showNA=TRUE, characterNA=NotAvailable"
  addDataFrame(data, sheet3, row.names=FALSE, showNA=TRUE,
               characterNA="NotAvailable")
  row <- getRows(sheet3, 1)
  cells <- getCells(row)
  c1.10 <- createCell(row, 10)
  setCellValue(c1.10[[1,1]], "rbind and cbind some df with addDataFrame")

  ## stack another data.frame on a sheet
  addDataFrame(data0, sheet3, startRow=17, startColumn=5)

  ## put another data.frame on a sheet side by side
  addDataFrame(data0, sheet3, startRow=17, startColumn=17)
}

wrapper_cellStyle <- function(wb) {
  ## Testing cell styles ...
  sheet  <- createSheet(wb, sheetName="cellStyles")
  rows   <- createRow(sheet, rowIndex=1:12)
  cells  <- createCell(rows, colIndex=1:8)

  mapply(setCellValue, cells[,1], month.name)

  ## Check borders of different colors.
  setCellValue(cells[[2,2]], paste("<-- Thick red bottom border,",
                                   "thin blue top border."))
  borders <- Border(color=c("red","blue"), position=c("BOTTOM", "TOP"),
                    pen=c("BORDER_THICK", "BORDER_THIN"))
  cs1 <- CellStyle(wb) + borders
  setCellStyle(cells[[2,1]], cs1)


  ## Check fills.
  setCellValue(cells[[4,2]], "<-- Solid lavender fill.")
  cs2 <- CellStyle(wb) + Fill(backgroundColor="lavender",
                              foregroundColor="lavender", pattern="SOLID_FOREGROUND")
  setCellStyle(cells[[4,1]], cs2)


  ## Check fonts.
  setCellValue(cells[[6,2]], "<-- Courier New, Italicised, in orange, size 20 and bold")
  font <- Font(wb, heightInPoints=20, isBold=TRUE, isItalic=TRUE,
               name="Courier New", color="orange")
  cs3 <- CellStyle(wb) + font
  setCellStyle(cells[[6,1]], cs3)


  ## Check alignment.
  setCellValue(cells[[8,2]], "<-- Right aligned")
  cs4 <- CellStyle(wb) + Alignment(h="ALIGN_RIGHT")
  setCellStyle(cells[[8,1]], cs4)


  ## Check dataFormat.
  setCellValue(cells[[10,1]], -12345.6789)
  setCellValue(cells[[10,2]], "<-- Format -12345.6789 in accounting style.")
  cs5 <- CellStyle(wb) + DataFormat("#,##0.00_);[Red](#,##0.00)")
  setCellStyle(cells[[10,1]], cs5)


  ## Autosize first, second column.
  autoSizeColumn(sheet, 1)
  autoSizeColumn(sheet, 2)
  setCellValue(cells[[1,4]], "First and second columns are autosized.")
}

wrapper_cellBlock <- function(wb) {
  ## Testing the CellBlock functionality ...
  sheet  <- createSheet(wb, sheetName="CellBlock")

  ## Add a cell block to sheet CellBlock
  cb <- CellBlock(sheet, 7, 3, 1000, 60)
  CB.setColData(cb, 1:100, 1)    # set a column
  CB.setRowData(cb, 1:50, 1)     # set a row

  # add a matrix, and style it
  cs <- CellStyle(wb) + DataFormat("#,##0.00")
  x  <- matrix(rnorm(900*45), nrow=900)
  CB.setMatrixData(cb, x, 10, 4, cellStyle=cs)

  # highlight the negative numbers in red
  fill <- Fill(foregroundColor = "red", backgroundColor="red")
  ind  <- which(x < 0, arr.ind=TRUE)
  CB.setFill(cb, fill, ind[,1]+9, ind[,2]+3)  # note the indices offset

  # set the border on the top row of the Cell Block
  border <-  Border(color="blue", position=c("TOP", "BOTTOM"),
                    pen=c("BORDER_THIN", "BORDER_THICK"))
  CB.setBorder(cb, border, 1:1000, 1)

  ## Modify the cell styles of existing cells on sheet dataFormats\n")
  sheets <- getSheets(wb)
  sheet  <- sheets[["dataFormats"]]
  cb <- CellBlock(sheet, 1, 1, 5, 5, create=FALSE)
  font <- Font(wb, color="red", isItalic=TRUE)
  CB.setBorder(cb, border, 1:5, 1)
  ind <- expand.grid(1:5, 1:5)
  CB.setFont(cb, font, ind[,1], ind[,2])

}

wrapper_comment <- function(wb) {
  ## Testing comments ...
  sheet <- createSheet(wb, "Comments")
  rows  <- createRow(sheet, rowIndex=1:10)      # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns

  cell1 <- cells[[1,1]]
  setCellValue(cell1, 1)   # add value 1 to cell A1

  createCellComment(cell1, "Cogito", author="Descartes")

  comment <- getCellComment(cell1)

  expect_identical(comment$getAuthor(),"Descartes")
  expect_identical(comment$getString()$toString(),'Cogito')
}

wrapper_dataFormat <- function(wb) {
  ## Testing dataFormats

  # create a test data.frame
  data <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
                     date=seq(as.Date("1999-01-01"), by="1 year", length.out=10),
                     bool=ifelse(1:10 %% 2, TRUE, FALSE), log=log(1:10),
                     rnorm=10000*rnorm(10),
                     datetime=seq(as.POSIXct("2011-11-06 00:00:00", tz="GMT"), by="1 hour",
                                  length.out=10))

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
  lapply(cells[,8], setCellStyle, cellStyle3)

  setCellValue(cells[[4,10]],
               'format "rnorm" column with two decimals, comma separator, red')
  cellStyle4 <- CellStyle(wb) + DataFormat("#,##0.00_);[Red](#,##0.00)")
  lapply(cells[,7], setCellStyle, cellStyle4)
}

wrapper_otherEffect <- function(wb) {
  ## Testing other effects ...

  sheet1 <- createSheet(wb, "otherEffects1")
  rows   <- createRow(sheet1, 1:10)              # 10 rows
  cells  <- createCell(rows, colIndex=1:8)       # 8 columns

  ## merge cells
  setCellValue(cells[[1,1]], "<-- a title that spans 3 columns")
  addMergedRegion(sheet1, 1, 1, 1, 3)

  ## set column width
  setColumnWidth(sheet1, 1, 25)
  setCellValue(cells[[5,1]], paste("<-- the width of this column",
                                   "is 20 characters wide."))

  ## set zoom
  setCellValue(cells[[3,1]], "<-- the zoom on this sheet is 2:1.")
  setZoom(sheet1, 200, 100)

  sheet2 <- createSheet(wb, "otherEffects2")
  rows  <- createRow(sheet2, 1:10)              # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns
  createFreezePane(sheet2, 5, 5, 8, 8)
  setCellValue(cells[[3,3]], "<-- a freeze pane")

  ## add hyperlinks to a cell
  address <- "https://poi.apache.org/"
  setCellValue(cells[[1,1]], "click me!")
  addHyperlink(cells[[1,1]], address)

  sheet3 <- createSheet(wb, "otherEffects3")
  rows  <- createRow(sheet3, 1:10)              # 10 rows
  cells <- createCell(rows, colIndex=1:8)       # 8 columns
  createSplitPane(sheet3, 2000, 2000, 1, 1, "PANE_LOWER_LEFT")
  setCellValue(cells[[3,3]], "<-- a split pane")
}

wrapper_picture <- function(wb) {
  ## Test embedding an R picture

  ## Add log_plot.jpeg to a new xlsx..
  picname <- test_ref('log_plot.jpeg')
  sheet <- createSheet(wb, "picture")

  addPicture(picname, sheet)

  addDataFrame(x = iris, sheet)
}

wrapper_range <- function(wb) {
  ## Testing ranges
  sheets <- getSheets(wb)
  sheet <- sheets[["dataFormats"]]

  ## make a new range
  firstCell <- sheet$getRow(2L)$getCell(2L)
  lastCell  <- sheet$getRow(6L)$getCell(5L)
  rangeName <- "Test2"
  createRange(rangeName, firstCell, lastCell)

  ranges <- getRanges(wb)
  range <- ranges[[1]]
  res <- readRange(range, sheet, colClasses="numeric")
}

## TO DO
wrapper_evalFormulasOnOpen <- function() {
  ## NOT PRESENTLY AVAILABLE
  filename <- "C:/Temp/formulaRevalueOnOpen.xlsx"
  wb <- loadWorkbook(filename)
  sheets <- getSheets(wb)

  sheet <- sheets[[1]]
  rows <- getRows(sheet)
  cells <- getCells(rows)

  setCellValue(cells[["2.1"]], 2)

  #wb$getCreationHelper()$createFormulaEvaluator()$evaluateAll()

  wb$setForceFormulaRecalculation(TRUE)

  saveWorkbook(wb, "C:/temp/junk.xlsx")

}
