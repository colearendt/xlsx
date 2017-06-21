context('test examples')

test_that('examples succeed', {
  test_examples(path=rprojroot::is_git_root$find_file('man/'))
})

