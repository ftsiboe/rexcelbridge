.onLoad <- function(libname, pkgname) {
  # Note: data.table's `:=` creates these columns at runtime. We register them here
  # so that R CMD check does not flag no visible binding for global variable.
  if (getRversion() >= "2.15.1") {
    utils::globalVariables(
      strsplit(
        "field setNames symbols_county_price value",
        "\\s+"
      )[[1]]
    )
  }
}
