
library(openxlsx)

#' @export
styles.aling.center <- createStyle(
  halign="center",
  valign="center"
)

#' @export
styles.align.left <- createStyle(
  halign="left",
  valign="center"
)

#' @export
styles.align.right <- createStyle(
  halign="right",
  valign="center"
)

#' @export
styles.border.bottom.thin <- createStyle(
  border="bottom",
  borderStyle="thin"
)

#' @export
styles.title.left <- createStyle(
  halign="left",
  valign="center",
  textDecoration="bold"
)

#' @export
styles.title.right <- createStyle(
  halign="right",
  valign="center",
  textDecoration="bold"
)

#' @export
styles.title.center <- createStyle(
  halign="center",
  valign="center",
  textDecoration="bold"
)

#' @export
styles.color.gray <- createStyle(
  fontColour="#666666"
)

#' @export
styles.color.error <- createStyle(
  fontColour="#FF0000"
)