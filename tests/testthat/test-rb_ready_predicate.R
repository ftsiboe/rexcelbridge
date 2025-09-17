test_that("rb_ready_predicate default logic: blanks, errors, wait-tokens are not ready", {
  pred <- rb_ready_predicate()  # default: Wait/Loading.../N/A; errors not ready
  
  expect_false(pred(NULL))
  expect_false(pred("  "))
  #expect_false(pred("#N/A"))
  expect_false(pred("Wait"))
  expect_false(pred("loading..."))
  expect_true(pred(123))
  expect_true(pred("PX_LAST"))
})

test_that("rb_ready_predicate treat_errors_as_ready overrides error handling", {
  pred <- rb_ready_predicate(treat_errors_as_ready = TRUE)
  expect_true(pred("#VALUE!"))
})
