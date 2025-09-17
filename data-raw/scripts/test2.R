rm(list = ls(all = TRUE)); gc()
library(RDCOMClient)

source("R/helpers_excel.R")
source("R/dtn_prophetx_evaluator.R")

# 1) Pure Excel test (should return 10)
print(dtn_prophetx_evaluator("=SUM(2,3,5)", visible = TRUE))

# 2) ProphetX RTD single-cell example:
#    Replace with your entitled symbol/field/date.
symbol <- "BEANS.20254"
scale  <- "Weekly"
field  <- "Close"
r_date <- as.Date("2025-09-08")
excel_serial <- as.integer(r_date - as.Date("1899-12-30"))  # e.g., 45908

px_formula <- sprintf(
  '=IFERROR(RTD("prophetx.rtdserver","","AIHIST",%s,%s,"1",%d,"",%s,"XD"),0)',
  excel_quote(symbol),
  excel_quote(scale),
  excel_serial,
  excel_quote(field)
)

print(dtn_prophetx_evaluator(px_formula, timeout_sec = 180, visible = TRUE))
