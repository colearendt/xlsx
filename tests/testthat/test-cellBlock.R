context('setMatrixData')

test_that('accepts character matrix', {
  # issue #19
  
  wb <- createWorkbook()
  sheet <- createSheet(wb)
  mtx <- matrix(c("hello", "world", "check1", "check2"), ncol=2)
  cb <- CellBlock(sheet, 1, 1, 2, 2)
  CB.setMatrixData(cb, mtx, 1, 1)
  fileName <- test_tmp("issue19.xlsx")
  saveWorkbook(wb, fileName)
  
  expect_true(TRUE)
})

test_that('preserves existing formats', {
  ## issue #22
  
  fileIn  <- test_ref("issue22.xlsx")
  fileOut <- test_tmp("issue22_out.xlsx")
  
  wb <- loadWorkbook(fileIn)
  sheets <- getSheets(wb)
  
  mat <- matrix(1:9, 3, 3)
  for (sheet in sheets) {
    if (sheet$getSheetName() == "Sheet1" ){
      # need to create the rows for Sheet1 as it is empty!  
      cb <- CellBlock(sheet, 1, 1, 3, 3, create = TRUE)   
    } else {
      cb <- CellBlock(sheet, 1, 1, 3, 3, create = FALSE) 
    }  
    CB.setMatrixData(cb, mat, 1, 1)
  }
  saveWorkbook(wb, fileOut)
  
  expect_true(TRUE)
})


test_that('formating applied to whole columns should not get lost in cell block formatting', {
  ## issue #32
  
  #using formatting on entire columns (A to L)
  wb <- loadWorkbook(test_ref("issue32_bad.xlsx"))
  sh <- getSheets(wb)
  s1 <- sh[[1]]
  
  cb <- CellBlock(sheet = s1, startRow = 1, startCol = 1,
                  noRows = 11, noColumns = 50, create = FALSE)
  # some of the formatting is lost (col I to L, rows 1 to 11)
  saveWorkbook(wb, test_tmp("issue32_format_lost.xlsx") )
  
  # the formatting is kept
  # using cells formatting only (columns A to L rows 1 to 10,000)
  wb <- loadWorkbook(test_ref("issue32_good.xlsx"))
  sh <- getSheets(wb)
  s1 <- sh[[1]]
  
  cb <- CellBlock(sheet = s1, startRow = 1, startCol = 1,
                  noRows = 11, noColumns = 50, create = FALSE)
  saveWorkbook(wb, test_tmp("issue32_format_kept.xlsx"))
  
  expect_true(TRUE)
})