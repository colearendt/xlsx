context(".onLoad")

test_java_version <- function(version){
  testthat::with_mock(
    `rJava::.jcall` = function(...){return(version)}
    ,xlsx:::.onLoad('.','xlsx')
  )
}
test_that("version comparison works", {
  skip("no longer checking version")
  # the previous version comparison failed in some locales
  expect_error(test_java_version("0.99.0"))
  expect_error(test_java_version("1.0.0"))
  expect_error(test_java_version("1.4.0"))
  expect_error(test_java_version("1.4.99"))
  test_java_version("1.5.0")
  test_java_version("1.6.0")
  test_java_version("1.7.0")
  test_java_version("1.9.0")
  test_java_version("7.0.0")
  test_java_version("10.0.2")
  test_java_version("14.0.4")
  test_java_version("100.0.4")
  test_java_version("1.8.0_121")
})


test_that("set_java_tmp_dir works", {
  new_tmp_dir <- tempfile("java_tmp_dir_")
  dir.create(new_tmp_dir, recursive = TRUE)

  current_tmp_dir <- get_java_tmp_dir()

  expect_equal(set_java_tmp_dir(new_tmp_dir), current_tmp_dir)
  expect_equal(get_java_tmp_dir(), new_tmp_dir)

  # set back
  set_java_tmp_dir(current_tmp_dir)

})
