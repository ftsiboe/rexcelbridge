#' Build a Bloomberg BDS Excel formula (bulk/descriptor data)
#'
#' Constructs a Bloomberg \code{BDS} formula string for bulk fields (returns a spilled table).
#' Use with \code{\link{rb_eval_single}} to read the result.
#'
#' @param security Character scalar, e.g., \code{"IBM US Equity"} or an index/ticker.
#' @param field Character scalar bulk field (e.g., \code{"INDX_MEMBERS"}).
#' @param overrides Optional named vector/list of overrides.
#'
#' @return A character string like:
#' \preformatted{
#' =BDS("SPX Index","INDX_MEMBERS")
#' }
#' @export
bloomberg_bds_formula <- function(security, field, overrides = NULL) {
  ov_txt <- ""
  if (!is.null(overrides) && length(overrides)) {
    if (is.null(names(overrides)) || any(!nzchar(names(overrides)))) {
      stop("`overrides` must be a *named* vector/list: e.g., c(Currency='USD').")
    }
    kv <- unlist(mapply(function(k, v) c(rb_excel_quote(k), rb_excel_quote(v)),
                        names(overrides), as.character(overrides), SIMPLIFY = FALSE),
                 use.names = FALSE)
    ov_txt <- paste0(",", paste(kv, collapse = ","))
  }
  sprintf('=BDS(%s,%s%s)',
          rb_excel_quote(security),
          rb_excel_quote(field),
          ov_txt)
}
