context('cell')

test_that('works in complex pipeline', {
  test_complex_cell('xlsx')
  test_complex_cell('xls')
})

context('setCellValue')

test_that(
  'NA can be output as empty', {
    ## issue #6
    tfile <- tempfile(fileext=".xlsx")
    
    wb <- createWorkbook()
    sheet <- createSheet(wb)
    rows <- createRow(sheet, rowIndex=1:5)
    cells <- createCell(rows)
    mapply(setCellValue, cells[,1], c(1,2,NA,4,5))
    setCellValue(cells[[2,2]], "Hello")
    mapply(setCellValue, cells[,3], c(1,2,NA,4,5), showNA=FALSE)
    setCellValue(cells[[3,3]], NA, showNA=FALSE)  
    saveWorkbook(wb, tfile)
    
    expect_identical(
      read.xlsx2(tfile,1,header=FALSE,stringsAsFactors=FALSE)
      , data.frame(
        X1=c('1','2','ERROR','4','5')
        ,X2=c('','Hello','','','')
        ,X3=c('1','2','','4','5')
        ,X4=c('','','','','')
        ,X5=c('','','','','')
        , stringsAsFactors=FALSE
      )
    )
    
    aux <- data.frame(x=c(1,2,NA,4,5))
    write.xlsx(aux, file=tfile, showNA=FALSE)
    
    expect_identical(
      read.xlsx(tfile,1,colIndex=2, stringsAsFactors=FALSE)
      , data.frame(
        x=c(1,2,NA,4,5)
        , stringsAsFactors=FALSE
      )
    )
  }
)
