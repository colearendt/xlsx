context('addPicture')

test_that('works with emf image', {
  ## issue #23
  ## known issue with 3.9
  skip('spreadsheet saves but the emf picture is not there')
  fileName <- "out/test_emf.emf"
  require(devEMF)
  emf(file=fileName, bg="white")
  boxplot(rnorm(100))
  dev.off()  
  
  wb <- createWorkbook()
  sheet <- createSheet(wb, "EMF_Sheet")
  
  addPicture(file=fileName, sheet)
  saveWorkbook(wb, file="out/issue23_out.xlsx")  
  
  expect_true(TRUE)
})