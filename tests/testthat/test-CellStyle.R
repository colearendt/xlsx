context('CellStyle')

test_that('Font + Fill did not set the font', {
  ## issue #41
  wb <- createWorkbook()
  cs <- CellStyle(wb) +
    Font(wb, heightInPoints=25, isBold=TRUE, isItalic=TRUE, color="red", name="Arial") + 
    Fill(backgroundColor="lavender", foregroundColor="lavender", pattern="SOLID_FOREGROUND") +
    Alignment(h="ALIGN_RIGHT")
  
  expect_false(is.null(cs$font$ref))
})

context('Font')

test_that('color black set properly', {
  ## https://bz.apache.org/bugzilla/show_bug.cgi?id=51236
  ## fixed in 3.14, apparently
  skip('Not presently working')
  ## issue #21
  wb <- createWorkbook()
  tmp <- Font(wb, color="black")
  
  tmp2 <- Font(wb, color=NULL)
  
  expect_match(capture_output(print(tmp$ref)),'.*<main:color.*rgb=\\\\"000000\\\\".*')
})

context('rowHeight')

test_that('set properly', {
  ## issue #43
  wb <- createWorkbook()
  sheet  <- createSheet(wb, sheetName="Sheet1")
  rows  <- createRow(sheet, rowIndex=1:5)
  cells <- createCell(rows, colIndex=1:5) 
  setRowHeight( rows, multiplier=3)
  
  saveWorkbook(wb, test_tmp("issue43.xlsx"))
})