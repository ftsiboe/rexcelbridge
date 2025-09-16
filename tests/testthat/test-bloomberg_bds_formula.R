test_that("bloomberg_bds_formula builds a basic BDS formula", {
  got <- bloomberg_bds_formula("SPX Index", "INDX_MEMBERS")
  exp <- '=BDS("SPX Index","INDX_MEMBERS")'
  expect_identical(got, exp)
})

test_that("bloomberg_bds_formula adds character overrides correctly (named vector)", {
  got <- bloomberg_bds_formula(
    "SPX Index",
    "INDX_MEMBERS",
    c(Currency = "USD", PricingSource = "BGN")
  )
  exp <- '=BDS("SPX Index","INDX_MEMBERS","Currency","USD","PricingSource","BGN")'
  expect_identical(got, exp)
})

test_that("bloomberg_bds_formula accepts overrides as a named list", {
  got <- bloomberg_bds_formula(
    "SPX Index",
    "INDX_MEMBERS",
    list(Currency = "USD", PricingSource = "BGN")
  )
  exp <- '=BDS("SPX Index","INDX_MEMBERS","Currency","USD","PricingSource","BGN")'
  expect_identical(got, exp)
})

# test_that("bloomberg_bds_formula handles numeric override values (unquoted per rb_excel_quote)", {
#   got <- bloomberg_bds_formula(
#     "SPX Index",
#     "INDX_MEMBERS",
#     c(Limit = 100)
#   )
#   # Note: 100 is unquoted since rb_excel_quote() returns numerics without quotes
#   exp <- '=BDS("SPX Index","INDX_MEMBERS","Limit",100)'
#   expect_identical(got, exp)
# })

test_that("bloomberg_bds_formula errors on unnamed overrides", {
  expect_error(
    bloomberg_bds_formula("SPX Index", "INDX_MEMBERS", c("USD")),
    "`overrides` must be a \\*named\\* vector/list"
  )
})

test_that("bloomberg_bds_formula errors when any override name is empty", {
  ov <- c("USD")
  names(ov) <- ""  # empty name
  expect_error(
    bloomberg_bds_formula("SPX Index", "INDX_MEMBERS", ov),
    "`overrides` must be a \\*named\\* vector/list"
  )
})

test_that("bloomberg_bds_formula escapes embedded quotes for security/field and overrides", {
  got <- bloomberg_bds_formula(
    'SPX "Index"',
    'INDX"MEMBERS',
    c(Notes = 'He said "Hi"')
  )
  # Excel escaping: internal " becomes ""
  exp <- '=BDS("SPX ""Index""","INDX""MEMBERS","Notes","He said ""Hi""")'
  expect_identical(got, exp)
})

test_that("bloomberg_bds_formula handles NA security/field by emitting empty strings", {
  got1 <- bloomberg_bds_formula(NA_character_, "INDX_MEMBERS")
  exp1 <- '=BDS("","INDX_MEMBERS")'
  expect_identical(got1, exp1)
  
  got2 <- bloomberg_bds_formula("SPX Index", NA_character_)
  exp2 <- '=BDS("SPX Index","")'
  expect_identical(got2, exp2)
})

test_that("bloomberg_bds_formula treats NA override values as empty strings", {
  got <- bloomberg_bds_formula(
    "SPX Index",
    "INDX_MEMBERS",
    c(Filter = NA_character_)
  )
  # rb_excel_quote(NA) -> '""'
  exp <- '=BDS("SPX Index","INDX_MEMBERS","Filter","")'
  expect_identical(got, exp)
})

test_that("bloomberg_bds_formula no-ops when overrides is length-0", {
  got <- bloomberg_bds_formula("SPX Index", "INDX_MEMBERS", overrides = character(0))
  exp <- '=BDS("SPX Index","INDX_MEMBERS")'
  expect_identical(got, exp)
})
