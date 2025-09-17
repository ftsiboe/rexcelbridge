test_that("rb_start_excel sets properties and returns a COM-like handle", {
  skip_if_not(.Platform$OS.type == "windows", "Windows-specific test")
  
  # Fake COMCreate to return a plain list (supports [[<- assignments)
  fake_COMCreate <- function(x) {
    structure(list(), class = "fakeCOM")
  }
  
  mockery::stub(rb_start_excel, "COMCreate", fake_COMCreate)
  
  xl <- rb_start_excel(visible = TRUE)
  # After rb_start_excel, these should exist with expected values
  expect_true(is.list(xl))
  expect_identical(xl[["Visible"]], TRUE)
  expect_identical(xl[["DisplayAlerts"]], FALSE)
  expect_identical(xl[["Calculation"]], -4105)
})
