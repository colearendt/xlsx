context('basic functions')

test_that('work in pipeline', {
  test_basicFunctions('xls')
  test_basicFunctions('xlsx')
})

test_that('add onto existing workbook', {
  test_addOnExistingWorkbook('xls')
  test_addOnExistingWorkbook('xlsx')
})

context('Memory Crash')

test_that('crashes with memory volume', {
  skip('do not worry at present')
  orig <- getOption('java.parameters')
  
  options(java.parameters="-Xmx4096m")  
  require(xlsx)
  
  N <- 500    # crash with N=5000, at sheet 20
  
  x <- as.data.frame(matrix(1:N, nrow=N, ncol=59))
  
  wb <- createWorkbook()
  for (k in 1:100) {
    cat("On sheet", k, "\n")
    sheetName <- paste("Sheet", k, sep="")
    sheet <- createSheet(wb, sheetName)
    
    addDataFrame(x, sheet)
  }
  
  saveWorkbook(wb, test_tmp("junk.xlsx"))
  
  options(java.parameters = orig)
})


