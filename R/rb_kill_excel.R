#' Force close all Excel processes (Windows only)
#'
#' Utility to terminate every running Excel instance on the system.
#' This is useful if RTD sessions or hidden COM objects are left running
#' after errors or crashes. Use with caution: unsaved work in Excel will be lost.
#'
#' @param force Logical; if TRUE, forcibly terminates without prompts (default).
#' @param verbose Logical; print output from taskkill. Default FALSE.
#'
#' @return Invisibly returns the exit status code from `taskkill` (0 = success).
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

