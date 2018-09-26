context('Workbook')

test_that('fails on incorrect type', {
  expect_error(createWorkbook('mine'),'Unknown format')
})

test_that("saves a workbook", {
  wb <- createWorkbook()
  tf <- fs::file_temp("save", ext = ".xlsx")
  
  saveWorkbook(wb, tf)
  wb_read <- loadWorkbook(tf)
})

test_that("saves a workbook with password", {
  skip("password currently broken")
  wb <- createWorkbook()
  pwd <- "password"

  tf <- fs::file_temp("save-pass", ext = ".xlsx")
  saveWorkbook(wb, tf, password = pwd)
  wb_read <- loadWorkbook(tf, password = pwd)
})

context('getSheets')

test_that('returns null with no sheets', {
  wb <- createWorkbook()
  
  expect_null(suppressMessages(getSheets(wb)))
  expect_output(getSheets(wb),'no sheets')
})


