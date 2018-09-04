context('Workbook')

test_that('fails on incorrect type', {
  expect_error(createWorkbook('mine'),'Unknown format')
})

context('getSheets')

test_that('returns null with no sheets', {
  wb <- createWorkbook()
  
  expect_null(suppressMessages(getSheets(wb)))
  expect_output(getSheets(wb),'no sheets')
})
