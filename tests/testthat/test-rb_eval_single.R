# tests/testthat/test-rb_eval_single.R

context("rb_eval_single")

library(mockery)

## ---- Helpers: fake COM Excel tree -----------------------------------------

# minimal "cell" env with Formula and Value2
make_cell <- function() {
  env <- new.env(parent = emptyenv())
  env[["Formula"]] <- NA_character_
  env[["Value2"]]  <- NULL
  env
}

# fake sheet: Cells$Item(r,c) -> env for that cell
make_fake_sheet <- function() {
  cells <- new.env(parent = emptyenv())

  Cells <- new.env(parent = emptyenv())
  Cells$Item <- function(r, c) {
    key <- paste(r, c, sep = ",")
    if (!exists(key, envir = cells, inherits = FALSE)) {
      assign(key, make_cell(), envir = cells)
    }
    get(key, envir = cells, inherits = FALSE)
  }

  sh <- new.env(parent = emptyenv())
  sh[["Cells"]] <- Cells
  sh
}

# workbook: exposes ActiveSheet and Close()
make_fake_workbook <- function(sheet) {
  wb <- new.env(parent = emptyenv())
  wb[["ActiveSheet"]] <- sheet
  wb$Close <- function(save = FALSE) invisible(TRUE)
  wb
}

# Excel.Application: Workbooks$Add(), CalculateFullRebuild(), Calculate(), Quit()
# updater(sheet) runs each time Calculate() is called to mutate cell$Value2
make_fake_excel_app <- function(sheet, updater) {
  xl <- new.env(parent = emptyenv())
  xl[["Visible"]]       <- FALSE
  xl[["DisplayAlerts"]] <- FALSE
  xl[["Calculation"]]   <- -4105

  Workbooks <- new.env(parent = emptyenv())
  Workbooks$Add <- function() make_fake_workbook(sheet)

  xl[["Workbooks"]] <- Workbooks
  xl$CalculateFullRebuild <- function() invisible(TRUE)
  xl$Calculate <- function() { updater(sheet); invisible(TRUE) }
  xl$Quit <- function() invisible(TRUE)
  xl
}

test_that("rb_eval_single times out cleanly and returns NA when never ready", {
  skip_if_not(.Platform$OS.type == "windows", "Windows-specific test")

  sh <- make_fake_sheet()

  # updater keeps the cell in 'Wait' forever
  updater <- function(sheet) {
    sheet[["Cells"]]$Item(1,1)[["Value2"]] <- "Wait"
  }

  fake_xl <- make_fake_excel_app(sh, updater)
  stub(rb_eval_single, "rb_start_excel", function(visible = FALSE) fake_xl)

  out <- rb_eval_single(
    formulas    = "=SOMEFORMULA()",
    timeout_sec = 0.5,
    ready_fn    = rb_ready_predicate(wait_tokens = "Wait")
  )

  expect_equal(nrow(out), 1)
  expect_false(out$ready[[1]])
  expect_true(is.na(out$result[[1]]))   # NULL-safe result becomes NA
})
