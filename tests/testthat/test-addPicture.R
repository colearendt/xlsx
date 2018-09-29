context('addPicture')

test_that('works with basic image', {

  f <- test_ref('log_plot.jpeg')
  wb <- createWorkbook()
  sheet <- createSheet(wb, "Image")
  
  pic <- addPicture(file=f, sheet)
  saveWorkbook(wb, file=test_tmp("addPicture.xlsx"))
  
  expect_true(.jinstanceof(pic,'org/apache/poi/xssf/usermodel/XSSFPicture'))
})
