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
  
  r <- read.xlsx2(fileName,1,header=FALSE,stringsAsFactors=FALSE)
  names(r) <- NULL
  r <- as.matrix(r)
  attr(r,'dimnames') <- NULL
  
  expect_identical(
    as.matrix(r)
    , mtx
  )
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
  
  wb <- loadWorkbook(fileOut)
  sheets <- getSheets(wb)
  
  cell1 <- getCells(getRows(sheets[[1]]))
  cell2 <- getCells(getRows(sheets[[2]]))
  cell3 <- getCells(getRows(sheets[[3]]))
  
  expect_null(unlist(unique(lapply(cell1,read_fill_foreground))))
  expect_identical(as.character(unique(lapply(cell2,read_fill_foreground))),'ffff00')
  expect_identical(as.character(lapply(cell3,read_fill_foreground))
                   , c('ffff00','NULL','ffff00'
                       ,'NULL','ffff00','NULL'
                       ,'NULL','NULL','ffff00')
                   )
})


test_that('formating applied to whole columns should not get lost in cell block formatting', {
  ## issue #32
  
  skip('not presently addressed')
  
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
