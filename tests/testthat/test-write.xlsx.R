context('write.xlsx')

test_that('DateTime classes work', {
  ## issue #45
  hours <- seq(as.POSIXct("2011-01-01 01:00:00", tz="GMT"),
               as.POSIXct("2011-01-01 10:00:00", tz="GMT"), by="1 hour")
  df <- data.frame(x=1:10, datetime=hours)
  
  write.xlsx(df, test_tmp("issue45.xlsx"))
  
  df.in <- read.xlsx2(test_tmp("issue45.xlsx"), sheetIndex=1,
                      colClasses=c("numeric", "numeric", "POSIXct"))
  df.in$datetime <- round(df.in$datetime)
  
  expect_identical(as.numeric(df.in$datetime),as.numeric(hours))
})

test_that('works in pipeline', {
  test_complex_export('xls')
  test_complex_export('xlsx')
})

context('write.xlsx2')

test_that('write password protected workbook succeeds', {
  ## issue #49
  
  x <- data.frame(values=c(1,2,3),stringsAsFactors=FALSE)
  filename <- test_tmp('issue49.xlsx')
  
  ## write
  write.xlsx2(x, filename, password='test', row.names=FALSE)
  
  ## read
  r <- read.xlsx2(filename, sheetIndex = 1, password='test'
                  , stringsAsFactors=FALSE
                  , colClasses = 'numeric')
  
  expect_identical(x, r)
})

context('low-level interface')

test_that('works in pipeline', {
  test_basic_export('xls')
  test_basic_export('xlsx')
})

test_that('password protecting workbook works', {
  wb <- createWorkbook()
  s <- createSheet(wb,'test123')
  addDataFrame(iris,s)
  
  filename <- test_tmp('password_test.xlsx')
  expect_null(saveWorkbook(wb,file=filename,password='test'))
  
  wb2 <- loadWorkbook(filename, password='test')
  
  expect_identical(names(getSheets(wb2))
                   , names(getSheets(wb)))
})