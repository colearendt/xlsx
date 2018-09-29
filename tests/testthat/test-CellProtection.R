context('Cell Protection')

test_that('locks cells appropriately', {
  ## issue #70

  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")
  rows <- createRow(sheet, rowIndex=1)
  cell.1 <- createCell(rows, colIndex=1)[[1,1]]
  cell.2 <- createCell(rows, colIndex=2)[[1,1]]
  
  setCellValue(cell.1, "just some text")
  setCellValue(cell.2, 'more text')
  
  cellStyle1 <- CellStyle(wb) +
    Font(wb, heightInPoints=20, isBold=TRUE, isItalic=TRUE,
         name="Courier New", color="green") +
    Alignment(h="ALIGN_RIGHT") +
    CellProtection(locked = T) #the cell protection part
  
  cellStyle2 <- cellStyle1 + CellProtection(locked=FALSE)
  
  setCellStyle(cell.1, cellStyle1)
  setCellStyle(cell.2, cellStyle2)
  
  expect_true(.jcall(getCellStyle(cell.1),'Z','getLocked'))
  expect_false(.jcall(getCellStyle(cell.2),'Z','getLocked'))
  
  saveWorkbook(wb, test_tmp('issue70.xlsx'))
})


test_that('tests correctly', {
  expect_true(is.CellProtection(CellProtection()))
  
  wb <- createWorkbook()
  expect_false(is.CellProtection(CellStyle(wb)))
})
