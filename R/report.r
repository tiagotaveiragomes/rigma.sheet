library(R6)
library(openxlsx)
library(googledrive)

#' Spreadsheet report
#'
#' @export

Report <- R6Class("Report",

  public=list(

    #' @field wb openxlsx.Workbook object
    wb=NULL,

    #' @field path the directory where the report is to be generated
    path=NULL,

    #' @field gdrive the directory where the report is to be generated
    gdrive=NULL,

    #' @field name the name of the report file
    name=NULL,

    #' @field title Human friendly version of the report name
    title=NULL,

    #' @field author Author that wrote the report
    author=NULL,

    #' @field sheets List that hols sheets object that show in the report
    sheets=NULL,

    #' @field ts Timestamp when the report was created
    timestamp=NULL,

    #' @field description Description of the contents of the report
    description=NULL,

    #' @field config.filename.ts Whether to add the timestamp to the report file name
    config.filename.ts=NULL,

    #' @description Initialize the report
    initialize = function(name, path, gdrive=NULL, title, description, author,
      config.filename.ts=TRUE
    ){
      self$path <- path
      self$name <- name
      self$title <- title
      self$author <- author
      self$gdrive <- gdrive
      self$sheets <- list()
      self$description <- description
      self$timestamp <- format(Sys.time(), "%Y%m%d_%H%M")

      self$config.filename.ts <- config.filename.ts

      self$wb <- createWorkbook(
        creator=self$author,
        title=self$title,
        subject=self$description
      )
    },

    #' @description Add a sheet object to the report
    sheet.add = function(sheet){
      if(!sheet$name %in% names(self$sheets)){
        addWorksheet(self$wb, sheet$name)
      }
      self$sheets[[sheet$name]] <- sheet$data
      writeData(self$wb, sheet$name, sheet$data,
        colNames=FALSE,
        rowNames=FALSE
      )

      for (fmt in sheet$formatting){
        op <- self[[fmt$op]]
        rlang::invoke(op, fmt$args)
      }
    },

    #' @description Save the report to disk
    save = function(){
      saveWorkbook(self$wb, self$uri.local, overwrite=TRUE)
      if(!is.null(self$gdrive)){
        print(self$uri.local)
        print(self$uri.drive)
        drive_upload(
          self$uri.local,
          self$uri.drive,
          overwrite=TRUE
        )
      }
    },

    sheet.pane.freeze = function(name, row, col){
      freezePane(self$wb, name, row, col)
    },

    sheet.col.width = function(name, cols, width){
      setColWidths(self$wb, name, cols, width)
    },

    sheet.add.style = function(name, style, rows, cols, expand, stack){
      addStyle(self$wb, name, style, rows, cols, expand, stack)
    },

    sheet.merge.header = function(name, col, row){
      mergeCells(self$wb, name, col, row)
    },

    sheet.conditional.formatting = function(name, rows, cols, rule, style){
      conditionalFormatting(self$wb, name, cols, rows, rule, style)
    }
  ),
  active=list(

    #' @description The current report extension
    extension = function(value){
      ".xlsx"
    },

    #' @description The filename for this report
    filename = function(value){
      ts <- ifelse(self$config.filename.ts, self$timestamp, "")
      paste0(c(self$name, ts, self$extension), collapse="")
    },

    #' @description The full uri where the report is stored
    uri.local = function(value){
      paste0(c(self$path, self$filename), collapse="/")
    },

    #' @description The full google driv euri where the report is stored
    uri.drive = function(value){
      paste0(c(self$gdrive, self$filename), collapse="/")
    }
  )
)