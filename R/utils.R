#' Returns NA rather than NaN if x is a vector of NA
#'
#' @param x
#'
#' @return
#'
#' @examples
mean_or_na <- function(x) {
  if (all(is.na(x)))
    return(NA)
  mean(x, na.rm = TRUE)
}
