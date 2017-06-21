context('addDataFrame')

test_that('can customize the format of datetimes in the output', {
  ## issue #26
  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")
  
  oldOpt <- options()
  
  format1 <- 'm/d/yyyy'
  format2 <- "dd MMM, yyyy"
  
  days <- seq(as.Date("2013-01-01"), by="1 day", length.out=5)
  
  options(xlsx.date.format=format1)
  addDataFrame(data.frame(days=days), sheet, startColumn=1,
               col.names=FALSE, row.names=FALSE)
  

  options(xlsx.date.format=format2)
  addDataFrame(data.frame(days=days), sheet, startColumn=2,
               col.names=FALSE, row.names=FALSE)
  
  options(oldOpt)
  
  f <- test_tmp("issue26_out.xlsx")
  saveWorkbook(wb, file=f)
  
  wb <- loadWorkbook(f)
  sheets <- getSheets(wb)
  r <- getRows(sheets[[1]])
  cell1 <- getCells(r,1)
  cell2 <- getCells(r,2)
  
  expect_identical(unique(as.character(lapply(cell1,read_data_format))),format1)
  
  expect_identical(unique(as.character(lapply(cell2,read_data_format))),format2)
})

test_that('does not fail with zero column data.frames', {
  ## issue #31
  wb <- createWorkbook()
  sheet <- createSheet(wb)
  df <- data.frame(x = 1:5)[,FALSE] #Data frame with some rows, no columns.
  res <- try(addDataFrame(df, sheet))
  
  expect_is(res,'CellBlock')
})