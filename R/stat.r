library(R6)

#' Handy Statistics
#'
#' @export

Stat <- R6Class("Stat",

  public=list(
    digits=NA,
    nsmall=NA,
    initialize = function(digits=2, nsmall=2){
      self$digits <- digits
      self$nsmall <- nsmall
    },
    preety = function(n){
      format(n, digits=self$digits, nsmall=self$nsmall)
    },
    avg = function(df, col){
      self$preety(mean(df[,col], na.rm=TRUE))
    },
    std = function(df, col){
      self$preety(sd(df[,col], na.rm=TRUE))
    },
    p50 = function(df, col){
      self$preety(median(df[,col], na.rm=TRUE))
    },
    iqr = function(df, col){
      self$preety(IQR(df[,col], na.rm=TRUE))
    },
    prev = function(n, t){
      self$preety(n / t * 100)
    },
    count = function(df, col){
      format(sum(!is.na(df[,col])), digits=0, nsmall=0)
    }
  )
)