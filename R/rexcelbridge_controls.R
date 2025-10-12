#' Controls for rexcelbridge session
#'
#' Centralized knobs for batching and Excel/ProphetX evaluation behavior.
#'
#' @param query_limit integer(1). Max rows per batch (used for chunking). Must be > 0. Default: `500L`.
#' @param continuous_integration_session logical(1). If `TRUE`, skip Excel/
#'   ProphetX evaluation and other side effects; used for CI runs. Default: `FALSE`.
#' @param excel_eval_max_attempts integer(1). Max retry attempts when evaluating
#'   Excel RTD queries (per chunk). Must be >= 1. Default: `2L`.
#' @param excel_eval_retry_sleep numeric(1). Sleep (seconds) between Excel RTD
#'   retry attempts. Must be >= 0. Default: `1`.
#'
#' @return Named list of control values (validated and sanitized).
#' @export
rexcelbridge_controls <- function(
    query_limit                    = 500L,
    continuous_integration_session = FALSE,
    excel_eval_max_attempts        = 2L,
    excel_eval_retry_sleep         = 1
){
  
  # ---- validate & sanitize 
  if (!is.numeric(query_limit) || length(query_limit) != 1L ||
      !is.finite(query_limit) || query_limit <= 0) {
    stop("`query_limit` must be a single positive number.")
  }
  query_limit <- as.integer(query_limit)
  
  if (!is.logical(continuous_integration_session) || length(continuous_integration_session) != 1L ||
      is.na(continuous_integration_session)) {
    stop("`continuous_integration_session` must be a single logical (TRUE/FALSE).")
  }
  
  if (!is.numeric(excel_eval_max_attempts) || length(excel_eval_max_attempts) != 1L ||
      !is.finite(excel_eval_max_attempts) || excel_eval_max_attempts < 1) {
    stop("`excel_eval_max_attempts` must be a single number >= 1.")
  }
  excel_eval_max_attempts <- as.integer(excel_eval_max_attempts)
  
  if (!is.numeric(excel_eval_retry_sleep) || length(excel_eval_retry_sleep) != 1L ||
      !is.finite(excel_eval_retry_sleep) || excel_eval_retry_sleep < 0) {
    stop("`excel_eval_retry_sleep` must be a single non-negative number (seconds).")
  }
  
  # ---- return
  list(
    query_limit   = query_limit,
    continuous_integration_session = continuous_integration_session,
    excel_eval_max_attempts        = excel_eval_max_attempts,
    excel_eval_retry_sleep         = excel_eval_retry_sleep
  )
}

