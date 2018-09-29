
test_ref <- function(file) {
  rprojroot::is_testthat$find_file(paste0('ref/',file))
}

test_tmp <- function(file) {
  tmp_folder <- rprojroot::is_testthat$find_file('tmp')
  if (!file.exists(tmp_folder))
    dir.create(tmp_folder)
  rprojroot::is_testthat$find_file(paste0('tmp/',file))
}

remove_tmp <- function() {
  tmp_folder <- rprojroot::is_testthat$find_file('tmp')
  if (file.exists(tmp_folder))
    unlink(tmp_folder, recursive=TRUE)
}
