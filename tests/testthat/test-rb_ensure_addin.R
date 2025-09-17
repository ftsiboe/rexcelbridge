make_fake_addin <- function(name, installed = FALSE) {
  env <- new.env(parent = emptyenv())
  env[["Name"]] <- name
  env[["Installed"]] <- installed
  env
}

make_fake_xl_with_addins <- function(names_vec) {
  addins <- new.env(parent = emptyenv())
  addins[["Count"]] <- length(names_vec)
  addin_envs <- lapply(names_vec, make_fake_addin)
  
  # Item(i) returns the i-th addin env
  addins$Item <- function(i) addin_envs[[i]]
  
  xl <- new.env(parent = emptyenv())
  xl[["AddIns"]] <- addins
  xl
}

test_that("rb_ensure_addin installs a matching add-in (case-insensitive)", {
  skip_if_not(.Platform$OS.type == "windows", "Windows-specific test")
  
  xl <- make_fake_xl_with_addins(c("Solver Add-in", "ProphetX", "Bloomberg"))
  # Initially Installed = FALSE
  expect_false(xl[["AddIns"]]$Item(2)[["Installed"]])
  expect_false(xl[["AddIns"]]$Item(3)[["Installed"]])
  
  rb_ensure_addin(xl, "prophetx")
  expect_true(xl[["AddIns"]]$Item(2)[["Installed"]])   # toggled TRUE
  expect_false(xl[["AddIns"]]$Item(3)[["Installed"]])  # untouched
  
  rb_ensure_addin(xl, "BLOOMBERG")
  expect_true(xl[["AddIns"]]$Item(3)[["Installed"]])
})

test_that("rb_ensure_addin is a no-op if no add-in matches", {
  skip_if_not(.Platform$OS.type == "windows", "Windows-specific test")
  
  xl <- make_fake_xl_with_addins(c("Analysis ToolPak", "Solver Add-in"))
  rb_ensure_addin(xl, "ProphetX")
  expect_false(xl[["AddIns"]]$Item(1)[["Installed"]])
  expect_false(xl[["AddIns"]]$Item(2)[["Installed"]])
})
