library(R6)
library(openxlsx)

#' Spreadsheet sheet object
#'
#' @export

Sheet <- R6Class("Sheet",

  public=list(

    #' @field name Name of the sheet
    name=NULL,

    #' @field title Title of the sheet
    title=NULL,

    #' @field data Data to write into the sheet
    data=NULL,

    #' @field sheet The actual spreadsheet object
    sheet=NULL,

    config=NULL,

    rules=NULL,

    styles=NULL,

    formatting=NULL,

    header=NULL,

    groups=NULL,

    sections=NULL,

    char.spacer=" ",

    char.width=10,

    size.ratio=7.5,

    coord.wildcard.any = "any",

    #' @description Initialize the sheet
    initialize = function(name, title){
      self$name <- name
      self$title <- title
      self$formatting <- list()
    },

    build.row.spacer = function(data){
      self$build.spacer(ncol(data))
    },

    build.spacer = function(n){
      rep(self$char.spacer, n)
    },

    add.format.op = function(op){
      self$formatting[[length(self$formatting)+1]] <- op
    },

    #' @description Creates a population characteristics table
    pop.char = function(sections, groups, header,
      config=NULL, styles=NULL, rules=NULL
    ){

      self$rules <- self$config.merge(rules, self$default.rules)
      self$styles <- self$config.merge(styles, self$default.styles)
      self$config <- self$config.merge(config, self$default.config)

      self$groups <- groups
      self$header <- header
      self$header$name <- "header"
      self$sections <- sections

      row <- 1
      col <- 1

      header <- self$pop.char.header(row, col)

      col <- header$col
      row <- header$row
      data <- header$data

      for (section in self$sections){
        if (self$config$section.spacer.rows >= 1){
          for (i in 1:self$config$section.spacer.rows){
            data <- rbind(data, self$build.row.spacer(data))
          }
        }
        row <- row + 1
        res <- self$pop.char.section(row, col, section)
        col <- res$col
        row <- res$row
        data <- rbind(data, res$data)
      }

      self$data <- data
      self$build.format.col.width()
      self$build.format.freeze.panes()
      self$build.format.header()
    },

    build.format.header = function(){
      start_col <- 2
      merge_step <- (ncol(self$data) - 1) / length(self$groups)
      for (group in self$groups){
        end_col <- start_col + merge_step - 1
        self$add.format.op(list(
            op="sheet.merge.header",
            args=list(
              name=self$name,
              col=start_col:end_col,
              row=1
            )
          )
        )
        start_col <- end_col + 1
      }
    },

    build.format.freeze.panes = function(){
      first.active.row <- 1
      first.active.col <- 1
      if(self$config$freeze.header == TRUE){
        first.active.row <- 2 + length(self$header$attrs)
      }
      if(self$config$freeze.attrs == TRUE){
        first.active.col <- 2
      }
      self$add.format.op(list(
          op="sheet.pane.freeze",
          args=list(
            name=self$name,
            row=first.active.row,
            col=first.active.col
          )
        )
      )
    },

    build.format.cell = function(row, col, styles){
      for (style in styles) {
        self$add.format.op(list(
            op="sheet.add.style",
            args=list(
              name=self$name,
              style=style,
              rows=row,
              cols=col,
              expand=TRUE,
              stack=TRUE
            )
          )
        )
      }
    },

    build.rule.cell = function(row, col, rules){
      for (rule in rules) {
        self$add.format.op(list(
            op="sheet.conditional.formatting",
            args=list(
              name=self$name,
              rows=row,
              cols=col,
              rule=rule$rule,
              style=rule$style
            )
          )
        )
      }
    },

    build.format.col.width = function(){
      fun <- function(col) max(nchar(col))
      col.max.chars <- apply(self$data, 2, fun)
      for (i in 1:length(col.max.chars)){
        chars <- col.max.chars[i]
        self$add.format.op(list(
          op="sheet.col.width",
          args=list(
            name=self$name,
            cols=c(i),
            width=round(self$char.width * chars / self$size.ratio)
          )
        ))
      }
    },

    build.coord.cell = function(s=NULL, g=NULL, a=NULL, v=NULL){
      s <- ifelse(is.null(s), self$coord.wildcard.any, s$name)
      g <- ifelse(is.null(g), self$coord.wildcard.any, g$name)
      a <- ifelse(is.null(a), self$coord.wildcard.any, a$name)
      v <- ifelse(is.null(v), self$coord.wildcard.any, v$name)
      paste0(c("s", s, "g", g, "a", a, "v", v), collapse=".")
    },

    process.cell.coords = function(row, col, s=NULL, g=NULL, a=NULL, v=NULL){
      coord <- self$build.coord.cell(s, g, a, v)
      if (coord %in% names(self$styles)){
        self$build.format.cell(
          row=row, col=col, styles=self$styles[[coord]]
        )
      }
      if (coord %in% names(self$rules)){
        self$build.rule.cell(
          row=row, col=col, rules=self$rules[[coord]]
        )
      }
    },

    process.cell.static = function(row, col, coord){
      if (coord %in% names(self$styles)){
        self$build.format.cell(
          row=row, col=col, styles=self$styles[[coord]]
        )
      }
      if (coord %in% names(self$rules)){
        self$build.rule.cell(
          row=row, col=col, rules=self$rules[[coord]]
        )
      }
    },

    pop.char.header = function(row, col){
      data <- c(self$header$title)
      self$process.cell.static(row, col, "header.main")

      for (group in self$groups) {
        data <- c(
          data,
          group$title,
          self$build.spacer(length(self$header$attrs)-1)
        )
        for(label in self$header$labels) {
          col <- col + 1
          self$process.cell.static(row, col, "header.group")
        }
      }
      row <- row + 1
      for (attr in self$header$attrs){
        col <- 1
        row.data <- c(attr$title)
        self$process.cell.static(row, col, "header.attr")
        for (group in self$groups) {
          for (i in 1:length(attr$stats)) {
            col <- col + 1
            stat <- attr$stats[[i]]
            label <- self$header$labels[[i]]
            row.data <- c(
              row.data,
              stat$fun(group$data, attr$name)
            )

            self$process.cell.coords(row, col)
            self$process.cell.coords(row, col, self$header)
            self$process.cell.coords(row, col, NULL, group)
            self$process.cell.coords(row, col, NULL, NULL, attr)
            self$process.cell.coords(row, col, NULL, NULL, NULL, label)
            self$process.cell.coords(row, col, self$header, group, attr, label)
          }
        }
        data <- rbind(data, row.data)
        row <- row + 1
      }
      list(data=data, row=row, col=1)
    },

    pop.char.section = function(row, col, section){
      data <- c(section$title)
      self$process.cell.static(row, col, "header.section")
      for (group in self$groups) {
        for (label in section$labels) {
          col <- col + 1
          data <- c(data, label$title)
          self$process.cell.static(row, col, "header.value")
        }
      }
      row <- row + 1
      for (attr in section$attrs){
        col <- 1
        row.data <- c(attr$title)
        self$process.cell.static(row, col, "header.attr")

        for (group in self$groups) {

          for (i in 1:length(attr$stats)) {
            col <- col + 1
            stat <- attr$stats[[i]]
            label <- section$labels[[i]]
            row.data <- c(
              row.data,
              stat$fun(group$data, attr$name)
            )

            self$process.cell.coords(row, col)
            self$process.cell.coords(row, col, section)
            self$process.cell.coords(row, col, NULL, group)
            self$process.cell.coords(row, col, NULL, NULL, attr)
            self$process.cell.coords(row, col, NULL, NULL, NULL, label)
            self$process.cell.coords(row, col, section, group, attr, label)
          }
        }
        data <- rbind(data, row.data)
        row <- row + 1
      }
      list(data=data, row=row, col=1)
    },

    config.merge = function(extra, base){
      if (is.null(extra)){
        return(base)
      }
      for (name in names(base)){
        if(!name %in% names(extra)){
          extra[[name]] <- base[[name]]
        }
      }
      return(extra)
    }
  ),
  active=list(
    default.config = function(){
      list(
        section.spacer.rows=1,
        freeze.header=TRUE,
        freeze.attrs=TRUE
      )
    },
    default.rules = function(){
      list(
        s.any.g.any.a.any.v.any=list(
          list(
            rule=rigma.sheet::rules.is.na,
            style=rigma.sheet::styles.color.error
          ),
          list(
            rule=rigma.sheet::rules.is.nan,
            style=rigma.sheet::styles.color.error
          )
        )
      )
    },
    default.styles = function(){
      list(
        header.main=list(rigma.sheet::styles.title.left),
        header.attr=list(rigma.sheet::styles.align.left),
        header.group=list(rigma.sheet::styles.title.right),
        header.section=list(
          rigma.sheet::styles.title.left,
          rigma.sheet::styles.border.bottom.thin
        ),
        header.value=list(
          rigma.sheet::styles.border.bottom.thin,
          rigma.sheet::styles.align.right
        ),
        s.any.g.any.a.any.v.any=list(rigma.sheet::styles.align.right)
      )
    }
  )
)