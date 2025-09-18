test_that("get_dtn_county_average live: Daily/Weekly/Monthly", {
  res <- get_dtn_county_average(
    commodity         = "WHEAT",
    start_date        = as.Date("2025-09-01"),
    end_date          = as.Date("2025-09-16"),
    time_scale        = "Daily",
    county_price_type = "County Average Spot Cash Price",
    fields            = c("Open","High","Low","Close","Volume","OpenInt"),
    state_abb         = "KS",
    control           = rexcelbridge_controls(continuous_integration_session=TRUE))
  testthat::expect_true(nrow(res) == 0L)
})
