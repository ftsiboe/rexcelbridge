#' Build a DTN ProphetX AIHIST RTD formula
#'
#' Constructs a single-cell Excel formula string for the DTN ProphetX
#' \code{AIHIST} RTD function. This can be passed to \code{rb_eval_single()}
#' to evaluate historical data directly from Excel
#' via COM automation.
#'
#' @param symbol Character. ProphetX instrument symbol (e.g., \code{"BEANS.20254.B"}).
#' @param time_scale Character. Time scale such as \code{"Daily"}, \code{"Weekly"}, or \code{"Monthly"}.
#' @param date Date or string convertible to Date. The Excel serial number is computed relative to 1899-12-30.
#' @param field Character. Field(s) to request, such as \code{"Open"}, \code{"High"},
#'   \code{"Low"}, \code{"Close"}, or \code{"Description"}.
#'
#' @return A character string containing a valid Excel formula of the form:
#' \preformatted{
#' =IFERROR(RTD("prophetx.rtdserver","","AIHIST",symbol,time_scale,"1",date,"",field,"XD"),0)
#' }
#'
#' @details 
#' - The returned string is not evaluated in R; it must be written into an Excel
#'   cell using \code{rb_eval_single()}.
#' - The \code{date} argument is converted to an Excel serial (days since 1899-12-30).
#' - Wrapping with \code{IFERROR(...,0)} ensures Excel returns \code{0} if the RTD call fails.
#' 
#' @export
dtn_prophetX_formula <- function(symbol, time_scale, date, field) {
  sprintf(
    '=IFERROR(RTD("prophetx.rtdserver","","AIHIST",%s,%s,"1",%d,"",%s,"XD"),0)',
    rb_excel_quote(symbol),
    rb_excel_quote(time_scale),
    as.integer(as.Date(date) - as.Date("1899-12-30")),
    rb_excel_quote(field)
  )
}
