library(broom)
library(openxlsx)
library(tidyverse) 


dfAll <- data.frame(var1 = rnorm(mean = 0, sd = 1, 1000)
                    , var2 = rnorm(mean = 17, sd = 4, 1000)
                    , var3 = gl(10, 100,1000)
                    , targ = sample(2,1000, prob = c(0.8,0.2), replace = TRUE)-1)
dfAll2 <- dfAll
dfAll2$var1 <- ifelse(dfAll2$var1 < -1,-1,dfAll2$var1)
dfAll2$var2 <- ifelse(dfAll2$var2 > 25,25,dfAll2$var2)

i <- sapply(dfAll, is.factor)
p <- sapply(dfAll, is.integer)



dfAll[i] <- lapply(dfAll[i], as.character)
dfAll[p] <- lapply(dfAll[p], as.numeric)


i <- sapply(dfAll2, is.factor)
p <- sapply(dfAll2, is.integer)



dfAll2[i] <- lapply(dfAll2[i], as.character)
dfAll2[p] <- lapply(dfAll2[p], as.numeric)

#dfAll$var1 <- dfAll$var1[sample(length(dfAll$var1), 25)]<-NA

Mined_Data = dfAll
Refined_Data = dfAll2

lm1 = lm(targ ~ ., data = Refined_Data)
Target = "targ"

summary(lm1)


# Make a list of predictors
names <- names(Refined_Data)
selected_vars <- as.list(setdiff(names(Refined_Data),Target))

# Naming the "rownames" column of the coefficients so that we can use the column properly
model_coeffs <- coef(lm1)
model_coeffs <- data.frame(model_coeffs)
model_coeffs <- cbind(Row.Names = rownames(model_coeffs), model_coeffs)
rownames(model_coeffs) <- NULL
names(model_coeffs) <- c("Level","Value")
# Create an empty list for loop output to go into

output <- list()
# Here we make 1 table per predictor by looping a filer(where) function
for (i in selected_vars) {
  tmp <- filter(model_coeffs, grepl(i, Level, fixed = TRUE))
  output[[i]] <- tmp
}

names(Refined_Data) <- paste(names(Refined_Data),"b", sep = "_")
dfJoin <- cbind(Mined_Data,Refined_Data)


#character <- names(Mined_Data[, sapply(Mined_Data, class) == 'character'])
character <- names(Mined_Data %>% 
                     select_if(is.character))
character <- character[character != Target]


outputChar <- list()
for (i in character) {
  p <- paste(i,"b", sep = "_")
  #tmp <- dplyr::filter(dfJoin, i == p)
  #tmp1 <- dplyr::filter(dfJoin, grepl(i, p, fixed = TRUE))
  tmp2 <- unique(dplyr::select(dfJoin,i,p))
  names(tmp2) <- c("Original", "Model")
  
  Label <- c("Categorical")
  
  tmp3 <- cbind(tmp2,Label)
  
  #if (dplyr::filter(tmp3, grepl(Original, Model, fixed = TRUE))){tmp3$Equal <- "Equal"}
  #if (tmp3$Original == tmp3$Model){tmp3$Equal <- "Equal"}
  #else{tmp3$Equal <- "Not equal"}
  
  
  tmp3$Equal <- as.numeric(tmp3$Original == tmp3$Model)
  tmp3$Equal[tmp3$Equal==1] <- "Equal"
  tmp3$Equal[tmp3$Equal==0] <- "Not equal"
  
  tmp3[order(tmp3$Equal),]
  
  
  cat('Comparing', i, 'to', p, '\n')
  
  outputChar[[i]] <- tmp3
}

################################################
numeric <- names(Mined_Data %>% 
                   select_if(is.numeric))
numeric <- numeric[numeric != Target]

for (i in numeric) {
  p <- paste(i,"b", sep = "_")
  
  tmp1 <- dfJoin %>% dplyr::filter(dfJoin[i] == dfJoin[p]) # equal cases
  tmp2 <- unique(dplyr::select(tmp1,i,p)) # unique cases
  
  tmp3 <- tmp2[order(tmp2[i]),] # sort cases
  names(tmp3) <- c("Original", "Model") # name columns
  
  tmp4 <- head(tmp3, n = 1L); tmp5 <- tail(tmp3, n= 1L);  # min,med,max
  if (nrow(tmp3) > 0){tmp6 <- tmp3[(nrow(tmp3)/2),]} 
  else {tmp6 = c("blank","blank")} # deals with error when table empty 
  
  tmp7 <- rbind(tmp4,tmp6,tmp5) # append
  
  Label <- c("min", "med", "max"); Equal <- "Equal"; tmp8 <- cbind(tmp7,Label,Equal); # Label rows
  
  cat('Grabbing where', i, 'is equal to', p, '\n')
  
  #### - Repeat for non-equal levels
  tmpne1 <- dfJoin %>% dplyr::filter(dfJoin[i] != dfJoin[p]) 
  tmpne2 <- unique(dplyr::select(tmpne1,i,p)) 
  
  tmpne3 <- tmpne2[order(tmpne2[i]),] # sort cases
  names(tmpne3) <- c("Original", "Model") 
  
  tmpne4 <- head(tmpne3, n = 1L); tmpne5 <- tail(tmpne3, n= 1L); 
  if (nrow(tmpne3) > 0){tmpne6 <- tmpne3[(nrow(tmpne3)/2),]} 
  else {tmpne6 = c("blank","blank")} # deals with error when table empty
  
  tmpne7 <- rbind(tmpne4,tmpne6,tmpne5) 
  
  Label <- c("min", "med", "max"); Equal <- "Not equal"; tmpne8 <- cbind(tmpne7,Label,Equal); 
  
  cat('Grabbing where', i, 'is not equal to', p, '\n')
  
  # Choose how many tables to output
  if (nrow(tmpne1) < 1){outputChar[[i]] <- tmp8}
  else if (nrow(tmp1) < 1){outputChar[[i]] <- tmpne8}
  else{outputChar[[i]] <- rbind(tmp8,tmpne8)}
}

# Make tables the same order for loops
outputChar <- outputChar[order(names(outputChar))]
output <- output[order(names(output))]


wb <- createWorkbook()
# Create 1 sheet per variable, each sheet is named after a variable.
lapply(seq_along(outputChar), function(i){
  addWorksheet(wb=wb, sheetName = names(outputChar[i])) # Create worksheets
  
  writeData(wb, sheet = i, "Transformations",startRow = 1) # Table title
  mergeCells(wb, i, cols = 1:4, rows = 1:2)
  addStyle(wb, i, createStyle(halign = "center", fontSize = 15, textDecoration = "bold"), rows = 1:2, cols = 1:2) # formats merged cells
  
  writeData(wb, sheet = i, outputChar[[i]],startRow = 3) # Transformation data
  addStyle(wb, i, createStyle(halign = "center", textDecoration = "bold"), rows = 3, cols = 1:4) # Format column headings
  
})


# Create 1 sheet per variable, also add room for DS comment.
lapply(seq_along(output), function(i){
  
  writeData(wb, sheet = i, "Encoding",startRow = 1, startCol = 7) # Table title
  mergeCells(wb, i, cols = 7:8, rows = 1:2)
  addStyle(wb, i, createStyle(halign = "center", fontSize = 15, textDecoration = "bold"), rows = 1:2, cols = 7:8) # formats merged cells
  
  
  writeData(wb, sheet = i, output[[i]],startRow = 3, startCol = 7)
  addStyle(wb, i, createStyle(halign = "center", textDecoration = "bold"), rows = 3, cols = 7:8) # Format column headings
  
  # Data Scientist comment 
  writeData(wb, sheet = i, "Data Scientist Comment",startRow = 1, startCol = 11) # Table title
  mergeCells(wb, i, cols = 11:12, rows = 1:2)
  addStyle(wb, i, createStyle(halign = "center", fontSize = 10, textDecoration = "bold", wrapText = TRUE), rows = 1:2, cols = 11:12) # formats merged cells
  
  mergeCells(wb, i, cols = 11:12, rows = 3:4)
  addStyle(wb, i, createStyle(halign = "center", fontSize = 10, textDecoration = "bold"), rows = 3:4, cols = 11:12) # formats merged cells
})


#Save Workbook
saveWorkbook(wb, "C:/Dunc/test.xlsx", overwrite = TRUE)
