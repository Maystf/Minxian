library(magrittr)
library(tidyverse)
library(readxl)
library(xlsx)
library(glue)


####### Preparation, the parameters that need manually input ######

# the directory of your original excels 
datdir_file <- 'C:/Users/Hu/Desktop/county/'
# what kind of data structure do you want
ID <- 2L
# If you assign with `1L` to parameter `ID`, that means you want a 
# single result excel (variables are defined with uniform names)
# if you assign with `2L`, this script will export the result
# excels province by province

# delete the area that shiqu, shixiaqu, jiaoqu
area_delete <- read_xlsx("C:/Users/Hu/Desktop/delete-area.xlsx") %>% pull(1L)

# the directory to save your result
datdir_result <- 'C:/Users/Hu/Desktop/result/'

######## Automatically Process ########

# This function is the core process of this script 
mod_gather <- function(x){
    subset(x, select = colnames(x)[-1L]) %>% 
        gather(., year, value, colnames(.)[-1L]) 
}

file <- dir(datdir_file)

filecd <- dir(datdir_file) %>% paste0(datdir_file, .)

county_str <- str_extract_all(file, pattern = "^[:alpha:]+") %>% 
     unlist() %>% unique()

file_index <- vector(mode = "list", length = length(county_str))
for (i in 1:length(county_str)) {
    file_index[[i]] <- str_which(file, pattern = glue("^{county_str[i]}"))
}

# create a empty lists, a list save one province's data
county <- vector(mode = "list", length = length(filecd))

for (i in 1:length(file)) {
    # import our excel data
    countable <- read_xls(filecd[i], col_names = TRUE, na = "") %>% 
        `[`(-c((nrow(.)-2):nrow(.)), )
    
    # rename the columns
    names(countable)[2] <- c("area")
    # deleter area
    countable %<>% filter(! area %in% area_delete)
    
    tablerow <- pull(countable, 1L)
    # the name of our variables
    varname <- pull(countable, 1L) %>% na.omit()
    # the number of our variables
    var_n <- length(varname)
    
    if(var_n == 1L){ # only one variable in a excel 
        
        print(file[i])
        
        # countable[ , 2] %<>% factor( ) %>% fct_inorder(ordered = TRUE)
        
        countable %<>% mod_gather()
        # again rename the columns
        names(countable) <- c("area", "year", varname[1L])

        county[[i]] <- countable %<>% arrange(area, year)
        
    } else if(var_n > 1L) { # there are numbers of variable existed in a excel
        
        county_sub <- vector(mode = "list", length = var_n)        
        
        rowindex_start <- match(varname, tablerow)
        
        rowindex_end <- c(rowindex_start[-1] - 1L, nrow(countable)) 
        
        for(k in 1:var_n) {
            
            county_sub[[k]] <- countable[rowindex_start[k]:rowindex_end[k], ]
            
            # county_sub[[k]][ , 2] %<>% factor( ) %>% fct_inorder(ordered = TRUE)
            
            }
            
        county_sub %<>% lapply(mod_gather) 
        
        for (j in 1:var_n) names(county_sub[[j]]) <- c("area", "year", varname[j])
                
        # combine the variables
        county_sub %<>% reduce(full_join, by = c("area", "year")) %>% arrange(area, year)
        
        county[[i]] <- county_sub
        
    } else glue("{file[i]}, Error! Please verify the structure of your imported excel.") %>% print()
}


if(ID == 1L){
    
    reduce(county, bind_rows) %>% 
        write_excel_csv(county, paste0(datdir_result, "result.xlsx"), na = "NA")
    
} else if(ID == 2L) {

    for ( i in 1:length(county_str)) {
        
        county %>% `[`(file_index[[i]]) %>% 
            reduce(full_join, by = c("area", "year")) %>% 
                write_excel_csv(paste0(datdir_result, county_str[i], ".xls"), na = "NA")
        
    }
}      
    
