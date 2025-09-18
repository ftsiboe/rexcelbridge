test_that("dtn_prophetX_formula builds the expected single-field formula (char date)", {
  got <- dtn_prophetX_formula(
    symbol = "BEANS.20254.B",
    time_scale  = "Weekly",
    date   = "2025-09-15",   # char input
    field  = "Open"
  )
  # 2025-09-15 Excel serial = 45908 (since 1899-12-30)
  exp <- '=IFERROR(RTD(\"prophetx.rtdserver\",\"\",\"AIHIST\",\"BEANS.20254.B\",\"Weekly\",\"1\",45915,\"\",\"Open\",\"XD\"),0)'
  expect_identical(got, exp)
})

test_that("dtn_prophetX_formula behavior with NA inputs (symbol/field) and NA date", {
  # NA symbol/field are quoted as empty string by rb_excel_quote(); date -> NA serial
  got <- dtn_prophetX_formula(
    symbol = NA_character_,
    time_scale  = "Weekly",
    date   = NA,                 # as.integer(as.Date(NA) - ...) => NA
    field  = NA_character_
  )
  # We don't enforce an error here; we assert current behavior: NA appears in date position.
  expect_true(grepl(',NA,', got, fixed = TRUE))
  expect_true(grepl(',"","XD"\\),0\\)$', got))
})
