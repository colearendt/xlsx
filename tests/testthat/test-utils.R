context(".onLoad")

test_java_version <- function(version){
  testthat::with_mock(
    `rJava::.jcall` = function(...){return(version)}
    ,xlsx:::.onLoad('.','xlsx')
  )
}
test_that("version comparison works", {
  test_java_version("1.7.0")
  test_java_version("1.9.0")
  test_java_version("7.0.0")
  test_java_version("10.0.2")
  test_java_version("14.0.4")
  test_java_version("100.0.4")
})
