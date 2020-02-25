context('write.xlsx')

test_that("can write to home directory", {
  # issue #108
  skip("broken at present")
  skip_on_cran()
  dat <- data.frame(1:10)
  filepath <- "~/__xx__xlsx__test_homedir.xlsx"

  # how can you protect such an operation?
  expect_null(write.xlsx(dat, file = filepath))
  read.xlsx(filepath)
  unlink(filepath)
})

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

test_that("works with a tibble", {
  local_tbl <- tibble::tibble(
    id = c(1,2,3),
    name = c("one", "two", "three"),
    again = c("do", "re", "mi")
    )

  tmp_xlsx <- tempfile(pattern = "tibble", fileext = ".xlsx")

  write.xlsx(local_tbl, file = tmp_xlsx)

  read_data <- read.xlsx(tmp_xlsx, 1)

  unlink(tmp_xlsx)

  skip("have not devised a test yet")
})

test_that("output data matches input (for a data.frame)", {
  local_df <- data.frame(id = c(1,2,3), name = c("one", "two", "three"))

  tmp_xlsx <- tempfile(pattern = "output", fileext = ".xlsx")

  expect_equal(write.xlsx(local_df, tmp_xlsx), local_df)

  unlink(tmp_xlsx)
})

test_that("output data is a data.frame (for a matrix)", {
  local_matrix <- matrix(data = c(1,2,3,4,5,6), nrow = 2)

  tmp_xlsx <- tempfile(pattern = "matrix_output", fileext = ".xlsx")

  expect_is(write.xlsx(local_matrix, tmp_xlsx), "data.frame")

  unlink(tmp_xlsx)
})

test_that("output data is a data.frame (for a column)", {
  local_c <- c(1,2,3,4,5,6)

  tmp_xlsx <- tempfile(pattern = "matrix_output", fileext = ".xlsx")

  expect_is(write.xlsx(local_c, tmp_xlsx), "data.frame")

  unlink(tmp_xlsx)
})

test_that("works with a matrix", {
  local_matrix <- matrix(data = c(1,2,3,4,5,6), nrow = 2)

  tmp_xlsx <- tempfile(pattern = "matrix", fileext = ".xlsx")

  write.xlsx(local_matrix, tmp_xlsx)

  return_matrix <- read.xlsx(tmp_xlsx, 1)

  unlink(tmp_xlsx)
  skip("broken - need to do a weird test for rownames")
})

test_that("works with a column", {
  local_c <- c(1,2,3,4,5,6)

  tmp_xlsx <- tempfile(pattern = "column", fileext = ".xlsx")

  write.xlsx(local_c, tmp_xlsx)

  return_c <- read.xlsx(tmp_xlsx, 1)

  skip("broken - need to do a weird test for rownames")
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
  skip_on_os(c("linux"))
  test_basic_export('xls')
  test_basic_export('xlsx')
})
