library(R6)
library(data.table)
library(stringr)

conductor <- R6Class("conductor",
  public = base::list(
    dt = NULL,
    available_filters = NULL,
    # Initialize the connection when the R6 object is created
    initialize = function(data) {
      data <-janitor::clean_names(data)
      if(data.table::is.data.table(data)==FALSE){
        self$dt <- data.table::as.data.table(data)  
      } else{
        self$dt <-data
      }

      self$available_filter <-c('in', 'like', 'between')

      base::message("Data table initalized")
    },

    add_leading_character = function(
      columns, 
      character_length=10, 
      column_type = 'str', 
      create_new_column=FALSE, 
      new_column_name='padded_column'
    ) {
      base::stopifnot(base::is.numeric(character_length))
      if(column_type == 'str'){pad_format <-base::paste0('%0',character_length,'s')}
      if(column_type == 'int'){pad_format <-base::paste0('%0',character_length,'d')}

      if(create_new_column==TRUE){
        self$dt[, (new_column_name):=base::lapply(.SD, base::sprintf, fmt=pad_format), .SDcols = columns]
      }else{
        self$dt[, (columns):=base::lapply(.SD, base::sprintf, fmt=pad_format), .SDcols = columns]
      }
    },

   clean_column_text = function(
    columns_to_transform, 
    transformations_vector=NULL, 
    remove_special_characters = FALSE,
    change_case = NULL
  ){
    base::stopifnot(
      base::is.atomic(columns_to_transform),
      base::is.logical(remove_special_characters)
     )

    case_map = base::list(
      'upper' = toupper,
      'lower' = tolower,
      'proper' = str_to_title
    )

    if(base::is.null(transformations_list)==FALSE){
      self$dt[, (columns_to_transform):=base::lapply(.SD, stringr::str_replace_all, transformations_vector), .SDcols = columns_to_transform]
    }

    if (remove_special_characters==TRUE){
      self$dt[, (columns_to_transform):=base::lapply(.SD, base::gsub,pattern='[^A-Za-z]+', replacement=''), .SDcols = columns_to_transform]
    }

    if(base::is.null(change_case)==FALSE){
      self$dt[, (columns_to_transform):=base::lapply(.SD, case_map[[change_case]]), .SDcols = columns_to_transform]
    }
  },

  create_flag = function(
    lookup_column,
    flag_mapping,
    flag_column_name = 'flag_column',
    value_in_mapping = 'Y',
    value_not_in_mapping = 'N'
  ){
    base::stopifnot(
      base::is.character(flag_column_name),
      base::is.character(lookup_column),
      base::is.character(value_in_mapping),
      base::is.character(value_not_in_mapping),
      base::is.atomic(flag_mapping)
    )
    lookup_list <-base::tolower(flag_mapping)
    self$dt[, flag_column:=base::lapply(.SD, function(x) base::ifelse(base::tolower(x) %in% flag_mapping, value_in_mapping, value_not_in_mapping)), .SDcols = lookup_column]
  },

  concatenate_by_group = function(
    concat_column, 
    groupby_column, 
    separator = ' | '
  ){
    self$dt[, .(group_concat = base::paste(base::get(concat_column), collapse =separator)), by = base::get(groupby_column)]
  },

  drop_rows = function(
    column_to_filter,
      filter, 
      filter_mode ='in', 
      inplace=FALSE
  ){# Drop rows that meet the criteria
    filter_mode <-base::tolower(filter_mode)
    
    base::stopifnot(
      base::is.logical(inplace),
      filter_mode %in% self$available_filters
    )

    if(filter_mode == 'in' && inplace==TRUE){
      self$dt <-self$dt[!get(column_to_filter) %in% filter]
    }

    if(filter_mode=='in' && inplace==FALSE){
      return(self$dt[!get(column_to_filter) %in% filter])
    }

    if(filter_mode == 'like' && inplace==TRUE){
      self$dt <-self$dt[!get(column_to_filter) %like% filter]
    }

    if(filter_mode == 'like' && inplace==FALSE){
      return(self$dt[!get(column_to_filter) %like% filter])
    }

    if(filter_mode == 'between' && inplace==TRUE){
      self$dt <-self$dt[!get(column_to_filter) %between% filter]
    }

    if(filter_mode == 'between' && inplace==FALSE){
      return(self$dt[!get(column_to_filter) %between% filter])
    }
  },

  filter_rows = function(
     column_to_filter,
      filter, 
      filter_mode ='in', 
      inplace=FALSE
  ){ # Select rows that meet the criteria
    filter_mode <-base::tolower(filter_mode)

    base::stopifnot(
      base::is.logical(inplace),
      filter_mode %in% self$available_filters
    )
    
    if(filter_mode == 'in' && inplace==TRUE){
       self$dt <-self$dt[base::get(column_to_filter) %in% filter ]
    }

    if(filter_mode=='in' && inplace==FALSE){
      return(self$dt[base::get(column_to_filter) %in% filter ])
    }

    if(filter_mode=='like' && inplace==TRUE){
      self$dt <-self$dt[base::get(column_to_filter) %like% filter ]
    }

    if(filter_mode == 'like' && inplace==FALSE){
      return(self$dt[base::get(column_to_filter) %like% filter ])
    }

    if(filter_mode == 'between' && inplace==TRUE){
      self$dt <-self$dt[base::get(column_to_filter) %between% filter ]
    }

    if(filter_mode == 'between' && inplace==FALSE){
      return(self$dt[base::get(column_to_filter) %between% filter ])
    }
  },

  select_columns = function(columns_to_select, inplace=FALSE){
      if(inplace==TRUE){
        self$dt <-self$dt[, ..columns_to_select]
      }else{
        return(self$dt[, ..columns_to_select])
      }
  },
 # Function to remove columns

  convert_to_dataframe = function(inplace=TRUE){
    if(inplace==TRUE){
      data.table::setDF(self$dt)
    }else{
      return(base::as.data.frame(self$dt))
    }
  },


  export_to_excel = function(workbook_path, output_list, header_style=NULL, body_style=NULL, ...){
    base::stopifnot(
      base::is.list(output_list)
    )

    options("openxlsx.dateFormat" = "yyyy-mm-dd")


    if(is.na(header_style)==TRUE){ # Default header style
      header_style <-createStyle(
        fgFill = "#000000", 
        halign = "CENTER", 
        textDecoration = "Bold",
        border = "Bottom", 
        fontColour = "white"
      )
    }

    workbook_path <-base::file.path(workbook_path)
    wb <-openxlsx::createWorkbook()

    for (i in length(output_list)){
      openxlsx::addWorksheet(wb, names(output_list[i]))
      openxlsx::writeData(
        wb, 
        sheet = i, 
        output_list[i], 
        rowNames = TRUE, 
        headerStyle = header_style,
        ...
      )
    }

    openxlsx::saveWorkbook(wb, workbok_path, overwrite=TRUE)
  }
  )
)

data_transformer <-function(data){
  cond <-conductor$new(data)
  return(cond)
}

