context('addDataFrame')

test_that('can customize the format of datetimes in the output', {
  ## issue #26
  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")
  
  days <- seq(as.Date("2013-01-01"), by="1 day", length.out=5)
  # use the default
  addDataFrame(data.frame(days=days), sheet, startColumn=1,
               col.names=FALSE, row.names=FALSE)
  
  # change the options temporarily
  oldOpt <- options()
  options(xlsx.date.format="dd MMM, yyyy")
  addDataFrame(data.frame(days=days), sheet, startColumn=2,
               col.names=FALSE, row.names=FALSE)
  options(oldOpt)
  
  
  saveWorkbook(wb, file=test_tmp("issue26_out.xlsx"))
  expect_true(TRUE)
})

test_that('does not fail with zero column data.frames', {
  ## issue #31
  wb <- createWorkbook()
  sheet <- createSheet(wb)
  df <- data.frame(x = 1:5)[,FALSE] #Data frame with some rows, no columns.
  res <- try(addDataFrame(df, sheet))
  
  expect_is(res,'CellBlock')
})