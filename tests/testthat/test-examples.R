context('test examples')

test_that('examples succeed', {
  test_examples(path=rprojroot::is_testthat$find_file('../../man/'))
})

