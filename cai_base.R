## Set your libpaths as required

library(broom)
library(openxlsx)
# Here we use 'tidy' instead of 'lm' as the output is easier to manipulate
# change 'lm1' to be the name of your model
trimmed = tidy(lm1)[,-c(3:5)]
trimmed<-data.frame(trimmed)
# create an empty list for loop outputted df to go into
output <- list()
# 'selected_vars1' must be a list of your predictors
# here we make 1 table per predictor by looping a filer(where) function
for (i in selected_vars1) {
  tmp <- filter(trimmed, grepl(i, term, fixed = TRUE))
  output[[i]] <- tmp
}
# create a blank workbook
wb <- createWorkbook()
# create 1 sheet per variable, each sheet is named after a variable.
lapply(seq_along(output), function(i){
  addWorksheet(wb=wb, sheetName = names(output[i]))
  writeData(wb, sheet = i, output[i][-length(output[[i]])])
})
#Save Workbook
saveWorkbook(wb, "test.xlsx", overwrite = TRUE)
