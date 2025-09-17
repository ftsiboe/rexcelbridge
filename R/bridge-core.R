
#' Start an Excel COM session (Windows)
#'
#' Launch a hidden (by default) Excel Application via COM and configure display alerts
#' and calculation mode (Automatic).
#'
#' @param visible Logical; show the Excel window. Default FALSE.
#' @return A COM handle to the Excel Application object.
#' @import RDCOMClient
#' @export
rb_start_excel <- function(visible = FALSE) {
  xl <- RDCOMClient::COMCreate("Excel.Application")
  xl[["Visible"]]       <- visible
  xl[["DisplayAlerts"]] <- FALSE
  xl[["Calculation"]]   <- -4105  # xlCalculationAutomatic
  xl
}

#' Ensure an Excel add-in is loaded (best effort)
#'
#' Tries to locate and enable an add-in by name pattern (case-insensitive).
#' Useful before evaluating formulas that depend on a vendor add-in.
#'
#' @param xl Excel Application COM handle from rb_start_excel().
#' @param pattern Character regex used to match the add-in name (e.g., "ProphetX", "Bloomberg").
#' @return Invisibly TRUE (no error if not found).
#' @export
rb_ensure_addin <- function(xl, pattern) {
  addins <- xl[["AddIns"]]
  n <- addins[["Count"]]
  for (i in seq_len(n)) {
    ai <- addins[["Item"]](i)
    nm <- ai[["Name"]]
    if (!is.null(nm) && grepl(pattern, nm, ignore.case = TRUE)) {
      if (!isTRUE(ai[["Installed"]])) ai[["Installed"]] <- TRUE
    }
  }
  invisible(TRUE)
}

#' Quote a value for use in an Excel formula
#'
#' Escapes embedded quotes and wraps non-numeric values in double quotes.
#'
#' @param x Scalar to embed in an Excel formula.
#' @return Character scalar safe to paste into a formula.
#' @export
rb_excel_quote <- function(x) {
  if (is.na(x)) return('""')
  if (is.numeric(x)) return(as.character(x))
  x <- gsub('"', '""', trimws(x), fixed = TRUE)
  paste0('"', x, '"')
}

# ---------- readiness helpers --------------------------------------------------

#' Build a readiness predicate
#'
#' Returns a function(value) that decides whether a cell is "ready".
#' By default, treats blanks, Excel errors, and vendor "wait" tokens as NOT ready.
#'
#' @param wait_tokens Character vector of tokens to be treated as "not ready"
#'   (e.g., c("Wait","Loading...")). Case-insensitive.
#' @param treat_errors_as_ready Logical; if TRUE, Excel errors stop the wait.
#' @return A function f(x) -> TRUE if ready, FALSE otherwise.
#' @export
rb_ready_predicate <- function(wait_tokens = c("Wait", "Loading...", "N/A"),
                               treat_errors_as_ready = FALSE) {
  pat <- if (length(wait_tokens)) paste0("^\\s*(", paste(wait_tokens, collapse = "|"), ")\\s*$") else NULL
  function(x) {
    if (is.null(x) || (is.character(x) && trimws(x) == "")) return(FALSE)
    if (!treat_errors_as_ready && is.character(x) && grepl("^#.+!$", x)) return(FALSE)
    if (!is.null(pat) && is.character(x) && grepl(pat, x, ignore.case = TRUE)) return(FALSE)
    TRUE
  }
}

#' Evaluate one or more single-cell Excel formulas
#'
#' Writes each formula into successive rows of column A in a new workbook,
#' polls until values are "ready" per a predicate, and returns the results.
#'
#' @param formulas Character vector of Excel formulas (e.g., "=SUM(1,2)",
#'   '=RTD("prophetx.rtdserver","","AIHIST",...)', '=BDP("IBM US Equity","PX_LAST")').
#' @param timeout_sec Max seconds to wait. Default 120.
#' @param visible Show the Excel window. Default FALSE.
#' @param ready_fn A readiness predicate built by rb_ready_predicate(). Default uses
#'   c("Wait","Loading...","N/A") as transient tokens.
#' @return Data frame with columns: row, formula, result, ready.
#' @export
rb_eval_single <- function(formulas,
                           timeout_sec = 120,
                           visible = FALSE,
                           ready_fn = rb_ready_predicate()) {
  stopifnot(length(formulas) >= 1)
  xl <- wb <- sh <- NULL
  on.exit({
    if (!is.null(wb)) try(wb$Close(FALSE), silent = TRUE)
    if (!is.null(xl)) try(xl$Quit(),        silent = TRUE)
    invisible(gc())
  }, add = TRUE)
  
  xl <- rb_start_excel(visible = visible)
  
  wb <- xl[["Workbooks"]]$Add()
  sh <- wb[["ActiveSheet"]]
  
  for (i in seq_along(formulas)) {
    cell <- sh[["Cells"]]$Item(i, 1)
    cell[["Formula"]] <- formulas[i]
  }
  
  try(xl$CalculateFullRebuild(), silent = TRUE)
  
  t0 <- Sys.time()
  ready <- rep(FALSE, length(formulas))
  vals  <- vector("list", length(formulas))
  sleep <- 0.1; max_sleep <- 0.75
  
  repeat {
    try(xl$Calculate(), silent = TRUE)
    for (i in seq_along(formulas)) {
      if (!ready[i]) {
        v <- sh[["Cells"]]$Item(i, 1)[["Value2"]]
        if (ready_fn(v)) { vals[[i]] <- v; ready[i] <- TRUE }
      }
    }
    if (all(ready)) break
    if (as.numeric(difftime(Sys.time(), t0, units = "secs")) > timeout_sec) break
    Sys.sleep(sleep); sleep <- min(max_sleep, sleep * 1.5)
  }
  
  # ---- NULL-safe materialization (fixes test failure) ----
  # map NULL -> NA, then create a simple atomic vector (no I(), no AsIs)
  vals_fixed <- lapply(vals, function(x) if (is.null(x)) NA else x)
  
  # try to keep numeric if everything numeric-ish; else character
  to_vec <- function(xs) {
    flat <- vapply(xs, function(z) if (is.null(z)) NA else z, FUN.VALUE = numeric(1), USE.NAMES = FALSE)
    # If coercion above failed (because not numeric), re-coerce to character
    if (all(is.na(flat)) && any(!vapply(xs, is.null, logical(1)))) {
      return(vapply(xs, function(z) if (is.null(z)) NA_character_ else as.character(z),
                    FUN.VALUE = character(1), USE.NAMES = FALSE))
    }
    flat
  }
  result_vec <- to_vec(vals_fixed)
  
  data.frame(
    row     = seq_along(formulas),
    formula = formulas,
    result  = result_vec,
    ready   = ready,
    stringsAsFactors = FALSE
  )
}




