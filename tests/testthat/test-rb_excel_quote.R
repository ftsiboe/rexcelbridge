test_that("rb_excel_quote quotes strings, leaves numerics, handles NA, escapes quotes", {
  expect_identical(rb_excel_quote("IBM US Equity"), '"IBM US Equity"')
  expect_identical(rb_excel_quote(123), "123")
  expect_identical(rb_excel_quote(NA_character_), '""')
  expect_identical(rb_excel_quote('He said "Hi"'), '"He said ""Hi"""')
  expect_identical(rb_excel_quote("  padded  "), '"padded"') # trims
})
