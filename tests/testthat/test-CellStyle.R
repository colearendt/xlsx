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

test_that('CellStyle + Font overwrites font', {
  ## not checking attributes at present... because there is not a nice way to do so
  wb <- createWorkbook()
  
  f1 <- Font(wb, heightInPoints=12, isBold=TRUE, isItalic=FALSE, color='black', name='Arial')
  cs <- CellStyle(wb) + f1
  
  f2 <- Font(wb, heightInPoints=10, isBold=FALSE, isItalic=TRUE, color='red', name='Calibri')
  cs2 <- cs + f2
  
  expect_identical(cs$font,f1)
  expect_identical(cs2$font,f2)
})

test_that('CellStyle + DataFormat overwrites DataFormat', {
  wb <- createWorkbook()
  
  df1 <- DataFormat('mm/dd/yyyy')
  cs <- CellStyle(wb) + df1
  
  df2 <- DataFormat('0.0')
  cs2 <- cs + df2
  
  expect_identical(df1,cs$dataFormat)
  expect_identical(df2,cs2$dataFormat)
})

test_that('CellStyle + Fill overwrites Fill', {
  wb <- createWorkbook()
  
  f1 <- Fill('red','red','SOLID_FOREGROUND')
  cs <- CellStyle(wb) + f1
  
  f2 <- Fill('black','black','SOLID_FOREGROUND')
  cs2 <- cs + f2
  
  expect_identical(f1,cs$fill)
  expect_identical(f2,cs2$fill)
})

test_that('CellStyle + Border overwrites Border', {
  wb <- createWorkbook()
  
  b1 <- Border(color='black',position=c('BOTTOM','TOP'),pen=c('BORDER_THIN','BORDER_THICK'))
  cs <- CellStyle(wb) + b1
  
  b2 <- Border(color='black', position=c('LEFT','RIGHT'),pen=c('BORDER_THICK','BORDER_THIN'))
  cs2 <- cs + b2
  
  expect_identical(b1,cs$border)
  expect_identical(b2,cs2$border)
})

test_that('CellStyle + Alignment overwrites Alignment', {
  wb <- createWorkbook()
  
  a1 <- Alignment(horizontal='ALIGN_CENTER', vertical = 'VERTICAL_TOP', wrapText = FALSE)
  cs <- CellStyle(wb) + a1
  
  a2 <- Alignment(horizontal='ALIGN_LEFT', vertical='VERTICAL_CENTER', wrapText=TRUE)
  cs2 <- CellStyle(wb) + a2
  
  expect_identical(a1, cs$alignment)
  expect_identical(a2, cs2$alignment)
})

test_that('CellStyle + CellProtection overwrites CellProtection', {
  wb <- createWorkbook()
  
  cp1 <- CellProtection(locked=FALSE,hidden=TRUE)
  cs <- CellStyle(wb) + cp1
  
  cp2 <- CellProtection(locked=TRUE,hidden=FALSE)
  cs2 <- CellStyle(wb) + cp2
  
  expect_identical(cp1, cs$cellProtection)
  expect_identical(cp2, cs2$cellProtection)
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
  
  expect_identical(as.numeric(lapply(rows,.jcall,'T','getHeight')),rep(900,5))
  
  saveWorkbook(wb, test_tmp("issue43.xlsx"))
})
