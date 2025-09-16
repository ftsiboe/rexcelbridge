# tests/testthat/test-rb_kill_excel.R

test_that("rb_kill_excel errors on non-Windows", {
  skip_on_os("windows")  # Only meaningful on non-Windows platforms
  
  expect_error(
    rb_kill_excel(),
    "only available on Windows",
    fixed = TRUE
  )
})

test_that("rb_kill_excel calls taskkill with /F when force = TRUE (default) and silent output", {
  skip_if_not(.Platform$OS.type == "windows", "Windows-specific test")
  
  # capture the args passed to system2
  called <- new.env(parent = emptyenv())
  called$args  <- NULL
  called$stdout <- NULL
  called$stderr <- NULL
  called$cmd <- NULL
  
  fake_system2 <- function(command, args = character(),
                           stdout = "", stderr = "", ...) {
    called$cmd    <- command
    called$args   <- args
    called$stdout <- stdout
    called$stderr <- stderr
    0L  # pretend success
  }
  
  # mock system2 inside rb_kill_excel
  mockery::stub(rb_kill_excel, "system2", fake_system2)
  
  res <- rb_kill_excel()  # defaults: force = TRUE, verbose = FALSE
  
  expect_identical(res, 0L)
  expect_identical(called$cmd, "taskkill")
  expect_identical(called$args, c("/IM", "EXCEL.EXE", "/F"))
  # verbose = FALSE -> stdout/stderr suppressed (FALSE)
  expect_identical(called$stdout, FALSE)
  expect_identical(called$stderr, FALSE)
})

test_that("rb_kill_excel omits /F when force = FALSE and enables output when verbose = TRUE", {
  skip_if_not(.Platform$OS.type == "windows", "Windows-specific test")
  
  called <- new.env(parent = emptyenv())
  fake_system2 <- function(command, args = character(),
                           stdout = "", stderr = "", ...) {
    called$cmd    <- command
    called$args   <- args
    called$stdout <- stdout
    called$stderr <- stderr
    0L
  }
  
  mockery::stub(rb_kill_excel, "system2", fake_system2)
  
  res <- rb_kill_excel(force = FALSE, verbose = TRUE)
  
  expect_identical(res, 0L)
  expect_identical(called$cmd, "taskkill")
  expect_identical(called$args, c("/IM", "EXCEL.EXE"))  # no /F
  # verbose = TRUE -> stdout/stderr are empty strings (inherit console)
  expect_identical(called$stdout, "")
  expect_identical(called$stderr, "")
})
