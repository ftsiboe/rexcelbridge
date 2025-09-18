#' Clear the package cache of downloaded data files
#'
#' Deletes the entire cache directory used by the **rexcelbridge** package to store
#' downloaded data files. Useful if you need to force re-download of data,
#' or free up disk space.
#' @family helpers
#' @return Invisibly returns `NULL`. A message is printed indicating which
#'   directory was cleared.
#' @export
#'
#' @examples
#' \dontrun{
#' # Remove all cached data files so they will be re-downloaded on next use
#' clear_rexcelbridge_cache()
#' }
clear_rexcelbridge_cache <- function(){
  dest_dir <- tools::R_user_dir("rexcelbridge", which = "cache")
  if (dir.exists(dest_dir)) {
    unlink(dest_dir, recursive = TRUE, force = TRUE)
  }
  message("Cleared cached files in ", dest_dir)
  invisible(NULL)
}

#' Controls for rexcelbridge session
#' @keywords internal
#' @noRd
#' @param dtn_prophetX_query_limit integer(1). Max rows per DTN ProphetX batch.
#' @param continuous_integration_session logical(1). Use CI-friendly, lightweight paths.
#' @return Named list of control values.
rexcelbridge_controls <- function(
    dtn_prophetX_query_limit = 3500,
    continuous_integration_session = FALSE
){
  list(
    dtn_prophetX_query_limit = dtn_prophetX_query_limit,
    continuous_integration_session = continuous_integration_session
  )
}

