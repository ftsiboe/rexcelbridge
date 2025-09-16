#' Evaluate a formula that spills a rectangular block
#'
#' Writes a single formula at (row, col) and reads back an nrow x ncol block.
#' Useful for array/RTD topics (e.g., multi-field history).
#'
#' @param formula Single Excel formula string.
#' @param nrow,ncol Expected output shape to read back.
#' @param top_left Integer vector c(row, col) where to place the formula. Default c(1,1).
#' @param timeout_sec Max seconds to wait. Default 180.
#' @param visible Show Excel window. Default FALSE.
#' @param ready_fn Readiness predicate; bottom-right cell is used as sentinel.
#' @return A data.frame with nrow rows and ncol columns (best-effort typed).
#' @export
rb_eval_block <- function(formula,
                          nrow,
                          ncol,
                          top_left = c(1L, 1L),
                          timeout_sec = 180,
                          visible = FALSE,
                          ready_fn = rb_ready_predicate()) {
  stopifnot(is.character(formula), length(formula) == 1L, nrow >= 1, ncol >= 1)
  xl <- wb <- sh <- NULL
  on.exit({
    if (!is.null(wb)) try(wb$Close(FALSE), silent = TRUE)
    if (!is.null(xl)) try(xl$Quit(),        silent = TRUE)
    invisible(gc())
  }, add = TRUE)
  
  xl <- rb_start_excel(visible = visible)
  
  wb <- xl[["Workbooks"]]$Add()
  sh <- wb[["ActiveSheet"]]
  
  r0 <- top_left[[1]]
  c0 <- top_left[[2]]
  
  # *** IMPORTANT: avoid complex assignment on the LHS ***
  tl_cell <- sh[["Cells"]]$Item(r0, c0)
  tl_cell[["Formula"]] <- formula
  
  try(xl$CalculateFullRebuild(), silent = TRUE)
  
  t0 <- Sys.time()
  sleep <- 0.1; max_sleep <- 0.75
  repeat {
    try(xl$Calculate(), silent = TRUE)
    br <- sh[["Cells"]]$Item(r0 + nrow - 1L, c0 + ncol - 1L)[["Value2"]]
    if (ready_fn(br)) break
    if (as.numeric(difftime(Sys.time(), t0, units = "secs")) > timeout_sec) break
    Sys.sleep(sleep); sleep <- min(max_sleep, sleep * 1.5)
  }
  
  rng <- sh[["Range"]](
    sh[["Cells"]]$Item(r0, c0),
    sh[["Cells"]]$Item(r0 + nrow - 1L, c0 + ncol - 1L)
  )
  vals <- rng[["Value2"]]
  
  # Coerce list-of-lists to data.frame, preserving text
  m <- matrix(NA_character_, nrow = nrow, ncol = ncol)
  for (i in seq_len(nrow)) {
    row_i <- vals[[i]]
    for (j in seq_len(ncol)) {
      v <- row_i[[j]]
      if (is.null(v)) {
        m[i, j] <- NA_character_
      } else if (is.numeric(v)) {
        m[i, j] <- as.character(v)
      } else {
        m[i, j] <- as.character(v)
      }
    }
  }
  df <- as.data.frame(m, stringsAsFactors = FALSE)
  
  # Try numeric where a clear majority are numeric
  for (j in seq_len(ncol)) {
    suppressWarnings({
      num <- as.numeric(df[[j]])
      if (sum(!is.na(num)) >= max(1L, floor(0.9 * length(num)))) df[[j]] <- num
    })
  }
  df
}