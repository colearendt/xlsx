library("devtools")

print('RUNNING CHECK')
revdep_check(bioconductor = TRUE)

print('ABOUT TO SAVE SUMMARY')
revdep_check_save_summary()

print('ABOUT TO PRINT PROBLEMS')
revdep_check_print_problems()
