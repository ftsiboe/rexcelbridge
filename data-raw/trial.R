# devtools::install_github("omegahat/RDCOMClient")
library(RDCOMClient)

start_excel <- function(visible = FALSE) {
  xl <- COMCreate("Excel.Application")
  xl[["Visible"]] <- visible
  xl[["DisplayAlerts"]] <- FALSE
  # Ensure automatic calc (xlCalculationAutomatic = -4105)
  xl[["Calculation"]] <- -4105
  xl
}

# Optional: try to ensure the ProphetX add-in is loaded
ensure_px_addin <- function(xl, pattern = "ProphetX") {
  addins <- xl[["AddIns"]]
  n <- addins[["Count"]]
  for (i in seq_len(n)) {
    ai <- addins[["Item"]](i)
    name <- ai[["Name"]]
    if (!is.null(name) && grepl(pattern, name, ignore.case = TRUE)) {
      if (!isTRUE(ai[["Installed"]])) ai[["Installed"]] <- TRUE
    }
  }
  invisible(TRUE)
}

# Core: evaluate a character vector of formulas (e.g., "=PXQ(...)", "=PXH(...)", etc.)
px_eval <- function(formulas, timeout_sec = 60, visible = FALSE) {
  stopifnot(length(formulas) >= 1)
  
  xl <- wb <- sh <- NULL
  on.exit({
    if (!is.null(wb)) try(wb$Close(FALSE), silent = TRUE)  # Close w/o saving
    if (!is.null(xl)) try(xl$Quit(),        silent = TRUE)
    invisible(gc())
  }, add = TRUE)
  
  # Start Excel (properties use [[...]])
  xl <- COMCreate("Excel.Application")
  xl[["Visible"]]       <- visible
  xl[["DisplayAlerts"]] <- FALSE
  xl[["Calculation"]]   <- -4105  # xlCalculationAutomatic
  
  # New workbook/sheet (Add is a method → use $)
  wb <- xl[["Workbooks"]]$Add()
  sh <- xl[["ActiveSheet"]]   # ActiveSheet is a property → NO parentheses
  
  # Write each formula: get the cell object first, then set its Formula
  for (i in seq_along(formulas)) {
    cell <- sh[["Cells"]]$Item(i, 1)   # don't write into a call; grab the COM obj
    cell[["Formula"]] <- formulas[i]
  }
  
  # Recalculate
  try(xl$CalculateFullRebuild(), silent = TRUE)
  
  # Poll for values (RTD can be async)
  t0 <- Sys.time()
  ready <- rep(FALSE, length(formulas))
  vals  <- vector("list", length(formulas))
  
  is_xl_err <- function(x) is.character(x) && grepl("^#.+!$", x)
  
  repeat {
    for (i in seq_along(formulas)) {
      if (!ready[i]) {
        v <- sh[["Cells"]]$Item(i, 1)[["Value2"]]  # Value2 avoids locale formatting
        if (!is.null(v) && !(is.character(v) && trimws(v) == "") && !is_xl_err(v)) {
          vals[[i]] <- v
          ready[i] <- TRUE
        }
      }
    }
    if (all(ready)) break
    if (as.numeric(difftime(Sys.time(), t0, units = "secs")) > timeout_sec) break
    Sys.sleep(0.25)
  }
  
  data.frame(
    row     = seq_along(formulas),
    formula = formulas,
    result  = I(unlist(vals)),
    ready   = ready,
    stringsAsFactors = FALSE
  )
}


# 1) Pure Excel test (no PX)
px_eval("=SUM(2,3,5)", visible = TRUE)

# 2) Your ProphetX RTD (make sure all topic args are quoted as strings)
f <- '=IFERROR(RTD("prophetx.rtdserver","","AIHIST","BEANS.20254.B","Weekly","1",45908,"","Open","XD"),0)'
px_eval(f, timeout_sec = 60, visible = TRUE)
















# Wrap an argument so Excel sees a proper string literal.
excel_quote <- function(x) {
  if (is.na(x)) return('""')
  if (is.numeric(x)) return(as.character(x))  # numeric OK unquoted; quote if you prefer
  x <- trimws(x)
  x <- gsub('"', '""', x, fixed = TRUE)      # escape " for Excel
  paste0('"', x, '"')
}

# Example inputs (swap in yours)
symbol      <- "BEANS.20254.B"
time_scale  <- "Weekly"     # e.g., "Daily","Weekly","Monthly" per your PX spec
field       <- "Open"
# If you already have an Excel serial like 45908, keep it numeric.
# If you have an R Date, convert to Excel serial:
r_date      <- as.Date("2025-09-15")
excel_date  <- as.integer(r_date - as.Date("1899-12-30"))  # 45908

# Build the RTD formula string
formula <- sprintf(
  '=IFERROR(RTD("prophetx.rtdserver","","AIHIST",%s,%s,"1",%s,"",%s,"XD"),0)',
  excel_quote(symbol),
  excel_quote(time_scale),
  excel_date,                 # or excel_quote(excel_date) if you prefer quoted
  excel_quote(field)
)




res <- px_eval(formulas = formula, timeout_sec = 60, visible = TRUE)
print(res)















