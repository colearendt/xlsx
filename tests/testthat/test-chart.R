context('chart')

test_that('basic example works', {
  ## issue #64
  skip('not currently supported')
  
  wb <- createWorkbook()
  sheet <- createSheet(wb, "mtcars")
  
  
  x <- 1:nrow(mtcars)
  y <- sort(mtcars$mpg)
  
  jDrawing <- sheet$createDrawingPatriarch();
  anchor <- structure(c(0, 0, 0, 0, 0, 5, 10, 5),
                      names=c('dx1', 'dx2', 'dy1', 'dy2', 'col1', 'row1', 'col2', 'row2'))
  jAnchor <- jDrawing$createAnchor(as.integer(anchor[1]),
                                   as.integer(anchor[2]),
                                   as.integer(anchor[3]),
                                   as.integer(anchor[4]),
                                   as.integer(anchor[5]),
                                   as.integer(anchor[6]),
                                   as.integer(anchor[7]),
                                   as.integer(anchor[8]))
  
  jChart <- jDrawing$createChart(jAnchor)
  
  # make the legend 
  jLegend <- jChart$getOrCreateLegend()
  jLegend$setPosition(J('org.apache.poi.ss.usermodel.charts.LegendPosition')$TOP_RIGHT)
  
  
  # set up the data
  jX <- .jcall('org.apache.poi.ss.usermodel.charts.DataSources',
               "Lorg/apache/poi/ss/usermodel/charts/ChartDataSource;",
               "fromArray",
               .jcast(.jarray(x), '[Ljava/lang/Object;'))
  
  jY <- .jcall('org.apache.poi.ss.usermodel.charts.DataSources',
               "Lorg/apache/poi/ss/usermodel/charts/ChartDataSource;",
               "fromArray",
               .jcast(.jarray(y), '[Ljava/lang/Object;'))
  
  jData <- jChart$getChartDataFactory()$createScatterChartData()
  jData$addSerie(jX, jY)
  
  # now you can make axes
  jXAxis <- jChart$createValueAxis(J('org.apache.poi.ss.usermodel.charts.AxisPosition')$BOTTOM)
  jYAxis <- jChart$createValueAxis(J('org.apache.poi.ss.usermodel.charts.AxisPosition')$LEFT)
  
  # plot the chart
  #  [8] "public void org.apache.poi.xssf.usermodel.XSSFChart.plot(org.apache.poi.ss.usermodel.charts.ChartData,org.apache.poi.ss.usermodel.charts.ChartAxis[])"
  
  jAxes <- .jarray(c(jXAxis, jYAxis)) 
  
  # NOT WORKING YET!
  .jcall(jChart,
         "V",
         "plot",
         jData,
         .jcast(jAxes, "[Lorg/apache/poi/ss/usermodel/charts/ChartAxis;"))
  
  .jcall(jChart,
         "V",
         "plot",
         jData,
         jAxes)
  
  
  
  jChart$plot(jData, .jcast(.jarray(c(jXAxis, jYAxis)), '[Lorg/apache/poi/ss/usermodel/charts/ChartAxis;'))
  
  
  
  #addDataFrame(mtcars, sheet, row.names=FALSE)
  
  saveWorkbook(wb, test_tmp("issue64.xlsx"))  
})
