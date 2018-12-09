context('Workbook')

test_that('fails on incorrect type', {
  expect_error(createWorkbook('mine'),'Unknown format')
})

test_that("save a workbook", {
  wb <- createWorkbook()
  tf <- test_tmp("test_save.xlsx")
  
  saveWorkbook(wb, tf)
  wb_read <- loadWorkbook(tf)
})

test_that('password protect workbook', {
  wb <- createWorkbook()
  s <- createSheet(wb,'test123')
  addDataFrame(iris,s)
  
  filename <- test_tmp('password_test.xlsx')
  expect_null(saveWorkbook(wb,file=filename,password='test'))
  
  wb2 <- loadWorkbook(filename, password='test')
  
  expect_identical(names(getSheets(wb2))
                   , names(getSheets(wb)))
})

context('getSheets')

test_that('returns null with no sheets', {
  wb <- createWorkbook()
  
  expect_null(suppressMessages(getSheets(wb)))
  expect_output(getSheets(wb),'no sheets')
})



