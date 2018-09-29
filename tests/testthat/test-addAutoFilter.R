context('addAutoFilter')

test_that('works in simple example', {
  ## issue #47
  hours <- seq(as.POSIXct("2011-01-01 01:00:00", tz="GMT"),
               as.POSIXct("2011-01-01 10:00:00", tz="GMT"), by="1 hour")
  data <- data.frame(x=1:10, type=rep(c("A", "B"), 5), datetime=hours)
  
  
  wb <- createWorkbook(type="xlsx")
  sheet  <- createSheet(wb, sheetName="Sheet1")
  addDataFrame(data, sheet, startRow=3, startColumn=2)
  af <- addAutoFilter(sheet, "C3:E3")
  saveWorkbook(wb, test_tmp("issue47.xlsx"))
  
  expect_true(.jinstanceof(af,'org/apache/poi/xssf/usermodel/XSSFAutoFilter'))
})
