library(magrittr)
library(tidyverse)
library(readxl)
library(xlsx)
library(glue)

# custom function
# Note that this function is the core tidy process of this script
mod_gather <- function(x){
    subset(x, select = colnames(x)[-1L]) %>%
        gather(., year, value, colnames(.)[-1L])
}

##### Preparation, the parameters that need manually input ######
setwd('D:/R/Minxian')
# the directory of your original excels
datdir_file <- file.path(getwd(), "county")
# the directory to save our result
datdir_result <- file.path(getwd(), "result")
if(!dir.exists(datdir_result)) dir.create(datdir_result)

# what kind of data structure are you want? ====
ID <- 2L
## If you assign ID with `1L`,
## that means you want a single result excel (variables are defined with uniform names)
## If you assign ID with `2L`,
## this script will export the result excels province by province

# delete the areas we doesn't want ====
area_delete <- read_xlsx("delete-area.xlsx") %>% pull(1L)

######## Automatically Process ########

# filename
file <- dir(datdir_file)
# path + filename
filecd <- paste(datdir_file, file, sep = "/")
# names of provinces
county_str <- str_extract_all(file, pattern = "^[:alpha:]+") %>%
     unlist() %>% unique()
# which files are blong to one province
file_index <- vector(mode = "list", length = length(county_str))
for (i in 1:length(county_str)) {
    file_index[[i]] <- str_which(file, pattern = glue("^{county_str[i]}"))
}
# create a empty lists, a list save one province's data
county <- vector(mode = "list", length = length(file))

for (i in 1:length(file)) { # loop in file

    # import our excel data
    countable <- read_xls(filecd[i], col_names = TRUE, na = "") %>%
        `[`(-c((nrow(.)-2):nrow(.)), )
    # rename the secound column (which are county name)for columns joining laterly
    names(countable)[2] <- c("area")
    # the row index of excel, we form it for spliting data frame
    # that there aren't just one variable in a excel
    tablerow <- pull(countable, 1L)
    # the name of our variables
    varname <- pull(countable, 1L) %>% na.omit()
    # the number of variables in a excel
    var_n <- length(varname)

    if(var_n == 1L){ # only one variable in a excel

        print(file[i])
        # delete the areas taht we do not want
        countable %<>% filter(! area %in% area_delete)
        # convert the data type of area from character to factor
        # doing this switch, we could order the areas in result excel
        # tidy the data frame using our custom gather function
        countable %<>% mod_gather()
        # name columns with the variable name
        names(countable) <- c("area", "year", varname[1L])
        # according to the original excel
        countable$area %<>% factor( ) %>% fct_inorder()
        # re-order the rows, firstly by the levels of county, and then by calendar
        county[[i]] <-  arrange(countable, area, year)

    } else if(var_n > 1L) { # there are numbers of variable existed in a excel

        county_sub <- vector(mode = "list", length = var_n)

        rowindex_start <- match(varname, tablerow)

        rowindex_end <- c(rowindex_start[-1] - 1L, nrow(countable))

        for(k in 1:var_n) {

            county_sub[[k]] <- countable[rowindex_start[k]:rowindex_end[k], ] %>%
                filter(!area %in% area_delete)

            county_sub[[k]]$area %<>% factor( ) %>% fct_inorder()

            }

        # tidy the data frame using our custom gather function
        county_sub %<>% lapply(mod_gather)
        # name columns with the variable name
        for (j in 1:var_n) names(county_sub[[j]]) <- c("area", "year", varname[j])
        # combine the variables
        county_sub %<>% reduce(full_join, by = c("area", "year"))
        # re-order the counties in one excel
        county[[i]] <-  arrange(county_sub, area, year)

    } else glue("{file[i]}, Error! Please verify the structure of your imported excel.") %>% print()
}

names(county) <- str_sub(file, start = 1L, end = -5L)

county_pro <- vector(mode = "list", length = length(county_str)) %>%
    set_names(county_str)

for ( i in 1:length(county_pro)) {

    # order the county by the levels of excel (had joined variable)
    order_excel <- county[file_index[[i]]] %>% map("area") %>%
        lapply(unique)

    order_length <- map_int(order_excel, length)

    order_excel <- which.max(order_length) %>% names()

    factor_order <- county[[order_excel]] %>% pull(area) %>% unique()

    # combine the excels which are belongs to one provinve to one excel
    county_pro[[i]] <- county[file_index[[i]]] %>%
         reduce(full_join, by = c("area", "year"))

    county_pro[[i]]$area %<>% factor(levels = factor_order)

    county_pro[[i]] %<>% arrange(area, year)

}


if(ID == 1L){

    bind_rows(county_pro) %>%
    write_excel_csv(paste0(datdir_result, "result.xlsx"), na = "NA")

} else if(ID == 2L) {

    for (i in 1:length(county_str)) {

        write_excel_csv(county_pro[[i]], paste0(datdir_result, "/", county_str[i], ".xls"), na = "NA")

    }
}

