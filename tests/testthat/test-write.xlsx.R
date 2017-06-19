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
  skip('Fails on UNIX at present')
  x <- data.frame(values=c(1,2,3))
  filename <- test_tmp('issue49.xlsx')
  
  ## write
  write.xlsx2(x, filename, password='test')
  
  ## read
  r <- read.xlsx2(filename, sheetIndex = 1, password='test')
  
  expect_identical(df, r)
})

context('low-level interface')

test_that('works in pipeline', {
  test_basic_export('xls')
  test_basic_export('xlsx')
})