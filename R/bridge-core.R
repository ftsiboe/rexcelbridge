
#' Start an Excel COM session (Windows)
#'
#' Launch a hidden (by default) Excel Application via COM and configure display alerts
#' and calculation mode (Automatic).
#'
#' @param visible Logical; show the Excel window. Default FALSE.
#' @return A COM handle to the Excel Application object.
#' @import RDCOMClient
#' @family excel bridge
#' @export
rb_start_excel <- function(visible = FALSE) {
  xl <- COMCreate("Excel.Application")
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
#' Defensive against environments where `AddIns` is missing or inaccessible.
#'
#' @param xl Excel Application COM handle from [rb_start_excel()].
#' @param pattern Character regex used to match the add-in name
#'   (e.g., "ProphetX", "Bloomberg").
#' @return Invisibly returns TRUE (never errors if not found).
#' @family excel bridge
#' @export
rb_ensure_addin <- function(xl, pattern) {
  addins <- try(xl[["AddIns"]], silent = TRUE)
  if (inherits(addins, "try-error") || is.null(addins)) return(invisible(TRUE))
  
  n <- try(addins[["Count"]], silent = TRUE)
  if (inherits(n, "try-error") || is.null(n) || !is.finite(n)) return(invisible(TRUE))
  
  for (i in seq_len(n)) {
    ai <- try(addins[["Item"]](i), silent = TRUE)
    if (inherits(ai, "try-error") || is.null(ai)) next
    nm <- try(ai[["Name"]], silent = TRUE)
    if (inherits(nm, "try-error") || is.null(nm)) next
    
    if (grepl(pattern, nm, ignore.case = TRUE)) {
      inst <- try(ai[["Installed"]], silent = TRUE)
      if (!inherits(inst, "try-error") && !isTRUE(inst)) {
        try(ai[["Installed"]] <- TRUE, silent = TRUE)
      }
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
#' @family excel bridge
#' @export
rb_excel_quote <- function(x) {
  if (is.na(x)) return('""')
  if (is.numeric(x)) return(as.character(x))
  x <- gsub('"', '""', trimws(x), fixed = TRUE)
  paste0('"', x, '"')
}


#' Build a readiness predicate
#'
#' Returns a function(value) that decides whether a cell is "ready".
#' By default, treats blanks, Excel errors, explicit NA values, and vendor
#' "wait" tokens as NOT ready.
#'
#' @param wait_tokens Character vector of tokens treated as "not ready"
#'   (e.g., `c("Wait","Loading...","N/A")`). Case-insensitive.
#' @param treat_errors_as_ready Logical; if TRUE, Excel errors stop the wait.
#' @return A function `f(x) -> TRUE` if ready, `FALSE` otherwise.
#' @family excel bridge
#' @export
rb_ready_predicate <- function(wait_tokens = c("Wait", "Loading...", "N/A"),
                               treat_errors_as_ready = FALSE) {
  pat <- if (length(wait_tokens)) paste0("^\\s*(", paste(wait_tokens, collapse = "|"), ")\\s*$") else NULL
  function(x) {
    if (is.null(x)) return(FALSE)
    if (length(x) != 1L) return(FALSE)
    if (is.na(x)) return(FALSE)
    if (is.character(x) && trimws(x) == "") return(FALSE)
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
#' @param timeout_sec Max seconds to wait. Default 1.
#' @param visible Show the Excel window. Default FALSE.
#' @param ready_fn A readiness predicate built by rb_ready_predicate(). Default uses
#'   c("Wait","Loading...","N/A") as transient tokens.
#' @return Data frame with columns: row, formula, result, ready.
#' @family excel bridge
#' @export
rb_eval_single <- function(
    formulas,
    timeout_sec = 1,
    visible = FALSE,
    ready_fn = rb_ready_predicate()) {
  
  on.exit(try(rb_kill_excel(), silent = TRUE), add = TRUE)
  
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
  
  repeat{
    try(xl$Calculate(), silent = TRUE)
    for (i in seq_along(formulas)){
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
    # ensure each element is length 1 or NULL
    stopifnot(all(vapply(xs, function(z) is.null(z) || length(z) == 1L, logical(1))))
    
    non_null <- Filter(Negate(is.null), xs)
    
    # if all non-null are coercible to numeric, return numeric; else character
    if (length(non_null) > 0L) {
      can_num <- {
        tmp <- suppressWarnings(as.numeric(unlist(non_null, use.names = FALSE)))
        # treat TRUE/FALSE and numeric strings as OK; NA only allowed if the source was NA
        all(!is.na(tmp) | is.na(unlist(non_null, use.names = FALSE)))
      }
    } else {
      can_num <- FALSE
    }
    
    if (can_num) {
      vapply(xs,
             function(z) if (is.null(z)) NA_real_ else as.numeric(z),
             FUN.VALUE = numeric(1),
             USE.NAMES = FALSE)
    } else {
      vapply(xs,
             function(z) if (is.null(z)) NA_character_ else as.character(z),
             FUN.VALUE = character(1),
             USE.NAMES = FALSE)
    }
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


#' Evaluate Excel/ProphetX RTD queries with retries
#'
#' @description
#' Efficiently evaluates a vector of Excel or ProphetX RTD-style queries with
#' retry logic, partial re-evaluation of failed queries, and optional parallel
#' fallback for single-query execution. This function is used to safely handle
#' transient Excel COM errors or timeouts when running multiple RTD calls
#' through the bridge.
#'
#' Key optimizations include:
#' \itemize{
#'   \item Deduplication of identical queries (evaluated once).
#'   \item Retrying only failed entries in subsequent attempts.
#'   \item Conditional restarting of Excel only after failed attempts.
#'   \item Optional batching for large query sets to avoid overload.
#'   \item Optional parallel evaluation for per-query fallback.
#' }
#'
#' @family excel bridge
#'
#' @param queries Character vector of Excel or ProphetX formulas
#'   (e.g., RTD calls like `"=RTD('prophetx.rtd', , 'ZCZ25')"`).
#' @param control Named list from [rexcelbridge_controls()]. Uses:
#'   \itemize{
#'     \item `control$excel_eval_max_attempts` (default = 3)
#'     \item `control$excel_eval_retry_sleep`  (default = 0.15)
#'     \item `control$continuous_integration_session` (skip evaluation in CI)
#'   }
#' @param require_non_na Logical(1). If `TRUE`, retries continue until all
#'   results are non-`NA` or attempts are exhausted. If `FALSE`, any
#'   length-correct result ends the loop.
#' @param eval_fn Function used to evaluate queries (default = [rb_eval_single]).
#'   The function may return:
#'   \itemize{
#'     \item a `data.frame` with a `result` column (preferred),
#'     \item a list with a `$result` vector, or
#'     \item an atomic vector of results.
#'   }
#' @param batch_size Integer(1). Maximum number of queries per batch in
#'   vectorized evaluation (default = `Inf` for all-at-once).
#' @param kill_excel_on_retry Logical(1). If `TRUE`, calls [rb_kill_excel()]
#'   before retrying after a failed attempt (default = `TRUE`).
#' @param parallel_fallback Logical(1). If `TRUE`, uses parallel evaluation for
#'   per-query fallback when vectorized evaluation fails (default = `FALSE`).
#' @param mc.cores Integer(1). Number of parallel workers used in fallback mode
#'   (default = `getOption("mc.cores", 1L)`).
#'
#' @return
#' Character vector of evaluated results with the same length as `queries`.
#' Each entry corresponds to the evaluated output (or `NA_character_` if
#' evaluation failed after all retries).
#'
#' @details
#' This function is primarily designed for backend Excel or ProphetX bridge
#' environments where calls to Excel COM objects may intermittently fail or
#' return incomplete results. It minimizes redundant evaluations and can safely
#' handle large-scale evaluation workloads by chunking and retrying only the
#' failed queries.
#'
#' @export
rb_eval_with_retries <- function(
    queries,
    control,
    require_non_na = TRUE,
    eval_fn = rb_eval_single,
    batch_size = Inf,              # e.g., 100-1000 to avoid huge vector calls
    kill_excel_on_retry = TRUE,    # only kill Excel *after* a failed pass
    parallel_fallback = FALSE,     # set TRUE if per-query fallback is slow
    mc.cores = max(1L, getOption("mc.cores", 1L))
) {
  stopifnot(is.character(queries))
  n <- length(queries)
  if (n == 0L) return(character(0))
  
  # CI short-circuit
  if (isTRUE(control$continuous_integration_session)) {
    return(rep(NA_character_, n))
  }
  
  # Retry settings
  max_attempts <- suppressWarnings(as.integer(control$excel_eval_max_attempts))
  if (!is.finite(max_attempts) || max_attempts < 1L) max_attempts <- 3L
  
  retry_sleep <- suppressWarnings(as.numeric(control$excel_eval_retry_sleep))
  if (!is.finite(retry_sleep) || retry_sleep < 0) retry_sleep <- 0.15
  
  # robust extractor that accepts df/list/vector
  extract_result <- function(x) {
    if (is.null(x)) return(NULL)
    if (is.data.frame(x)) {
      if ("result" %in% names(x)) return(as.character(x[["result"]]))
      # if unnamed single column, take it
      if (ncol(x) == 1L) return(as.character(x[[1L]]))
      return(NULL)
    }
    if (is.list(x) && !is.null(x$result)) return(as.character(x$result))
    if (is.atomic(x)) return(as.character(x))
    NULL
  }
  
  # 1) De-duplicate to avoid redundant work
  uq  <- unique(queries)
  inv <- match(queries, uq)   # indices to rebuild in original order
  
  # Pre-allocate
  uq_out   <- rep(NA_character_, length(uq))
  remaining <- seq_along(uq)  # indices into uq that are still NA
  
  # Helper: safe vector eval for a chunk (no extra args!)
  safe_eval_vec <- function(qs) {
    res <- try(eval_fn(qs), silent = TRUE)
    if (inherits(res, "try-error")) {
      return(rep(NA_character_, length(qs)))
    }
    res <- extract_result(res)
    if (is.null(res)) {
      return(rep(NA_character_, length(qs)))
    }
    # If backend returned wrong length, fallback per-query for this chunk only
    if (length(res) != length(qs)) {
      if (!parallel_fallback || length(qs) < 2L) {
        res <- vapply(qs, function(q) {
          r1 <- try(eval_fn(q), silent = TRUE)
          r1 <- extract_result(r1)
          if (is.null(r1)) NA_character_ else as.character(r1)[1]
        }, FUN.VALUE = character(1))
      } else {
        res <- parallel::mclapply(qs, function(q) {
          r1 <- try(eval_fn(q), silent = TRUE)
          r1 <- extract_result(r1)
          if (is.null(r1)) NA_character_ else as.character(r1)[1]
        }, mc.cores = mc.cores)
        res <- vapply(res, `[`, character(1), 1L)
      }
    } else {
      res <- as.character(res)
    }
    res
  }
  
  # Attempt loop - only re-run the *remaining* queries
  for (attempt in seq_len(max_attempts)) {
    if (attempt > 1L && kill_excel_on_retry) {
      # Only kill/restart Excel if the previous pass left missing values
      try(rb_kill_excel(), silent = TRUE)
    }
    
    if (length(remaining) == 0L) break
    
    # 2) Batch the remaining work (prevents huge single calls)
    if (is.finite(batch_size) && batch_size > 0L && batch_size < length(remaining)) {
      idx_batches <- split(remaining, ceiling(seq_along(remaining) / batch_size))
    } else {
      idx_batches <- list(remaining)
    }
    
    for (batch in idx_batches) {
      qs  <- uq[batch]
      res <- safe_eval_vec(qs)
      uq_out[batch] <- res
    }
    
    # Update remaining set
    if (require_non_na) {
      remaining <- which(is.na(uq_out))
    } else {
      remaining <- integer(0)
    }
    
    if (length(remaining) == 0L) break
    if (attempt < max_attempts && retry_sleep > 0) Sys.sleep(retry_sleep)
  }
  
  # Rebuild in the original order
  out <- uq_out[inv]
  out[is.na(out)] <- NA_character_
  out
}


#' Force close all Excel processes (Windows only)
#'
#' Utility to terminate every running Excel instance on the system.
#' Useful if RTD sessions or hidden COM objects are left running after errors
#' or crashes. **Use with caution:** unsaved work in Excel will be lost.
#'
#' @note Windows-only; errors on non-Windows platforms.
#' @param force Logical; if TRUE, forcibly terminates without prompts (default).
#' @param verbose Logical; print output from `taskkill`. Default FALSE.
#' @return Invisibly returns the exit status code from `taskkill` (0 = success).
#' @family excel bridge
#' @export
rb_kill_excel <- function(force = TRUE, verbose = FALSE) {
  if (.Platform$OS.type != "windows") {
    stop("rb_kill_excel() is only available on Windows.")
  }
  args <- c("/IM", "EXCEL.EXE")
  if (force) args <- c(args, "/F")
  res <- suppressWarnings(
    system2("taskkill", args = args,
            stdout = if (verbose) "" else FALSE,
            stderr = if (verbose) "" else FALSE)
  )
  invisible(res)
}
