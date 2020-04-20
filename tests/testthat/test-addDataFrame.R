context('addDataFrame')

test_that('can customize the format of datetimes in the output', {
  ## issue #26
  wb <- createWorkbook()
  sheet <- createSheet(wb, "Sheet1")

  oldOpt <- options()

  format1 <- 'm/d/yyyy'
  format2 <- "dd MMM, yyyy"

  days <- seq(as.Date("2013-01-01"), by="1 day", length.out=5)

  options(xlsx.date.format=format1)
  addDataFrame(data.frame(days=days), sheet, startColumn=1,
               col.names=FALSE, row.names=FALSE)


  options(xlsx.date.format=format2)
  addDataFrame(data.frame(days=days), sheet, startColumn=2,
               col.names=FALSE, row.names=FALSE)

  options(oldOpt)

  f <- test_tmp("issue26_out.xlsx")
  saveWorkbook(wb, file=f)

  wb <- loadWorkbook(f)
  sheets <- getSheets(wb)
  r <- getRows(sheets[[1]])
  cell1 <- getCells(r,1)
  cell2 <- getCells(r,2)

  expect_identical(unique(as.character(lapply(cell1,read_data_format))),format1)

  expect_identical(unique(as.character(lapply(cell2,read_data_format))),format2)
})

test_that('does not fail with zero column data.frames', {
  ## issue #31
  wb <- createWorkbook()
  sheet <- createSheet(wb)
  df <- data.frame(x = 1:5)[,FALSE] #Data frame with some rows, no columns.
  res <- try(addDataFrame(df, sheet))

  expect_is(res,'CellBlock')
})

test_that('works with tibble', {
  ## issue #73
  skip('not presently working')

  skip_if_not_installed('tibble')
  d <- as_tibble(iris)

  wb <- createWorkbook()
  s <- createSheet(wb,'test')

  addDataFrame(d,s,row.names=FALSE)

  f <- test_tmp('issue73.xlsx')
  saveWorkbook(wb,f)

  r <- read.xlsx2(f,1,stringsAsFactors=FALSE)

  expect_identical(r[c(1,150),1],c('5.1','5.9'))
  expect_identical(r[c(1,75),2],c('3.5','2.9'))
})

test_that("works with characterNA = NA", {
  wb <- createWorkbook()
  s1 <- createSheet(wb)

  addDataFrame(
    data.frame(a = c(NA,"hi"), b = c("ho","hum"), c = c(NA_character_, NA_character_)),
    s1,
    characterNA = NA
    )

  read_dat <- readRows(s1, 1,3, 1)
  alt_dat <- readColumns(s1, 1,4, 1)

  expect_identical(read_dat[1,4], "c")
  expect_identical(read_dat[2,2], "")
  expect_identical(read_dat[2,4], "")
})
