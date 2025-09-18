# Live integration tests for get_dtn_county_average()
# These actually hit ProphetX via Excel when DTN_PX_LIVE=1.
test_that("get_dtn_county_average live: Daily/Weekly/Monthly", {
  # Gate live tests to safe environments only
  testthat::skip_if_not(Sys.getenv("DTN_PX_LIVE") == "1",
                        "Set DTN_PX_LIVE=1 to enable live ProphetX Excel tests")
  testthat::skip_if_not(.Platform$OS.type == "windows",
                        "Excel COM/ProphetX only available on Windows")
  testthat::skip_if_not(requireNamespace("rexcelbridge", quietly = TRUE),
                        "rexcelbridge not installed")
  # Basic objects/functions required by pipeline
  testthat::skip_if_not(exists("symbols_county_price"),
                        "symbols_county_price not found in environment")
  testthat::skip_if_not(exists("dtn_prophetX_formula"),
                        "dtn_prophetX_formula() not found in environment")
  testthat::skip_if_not(exists("rb_eval_single") && exists("rb_kill_excel"),
                        "rexcelbridge entry points not found")
  
  # Choose a small, valid slice from available symbols to ensure a hit
  allowed_types <- c(
    "County Average Spot Cash Price",
    "County Average Cash Price",
    "County Average Spot Basis Price",
    "County Average Basis"
  )
  
  cand <- subset(symbols_county_price,
                 county_price_type %in% allowed_types &
                   !is.na(state_abbreviation) & !is.na(county_name) &
                   nzchar(commodity))
  
  testthat::skip_if(nrow(cand) == 0L, "No matching rows in symbols_county_price")
  
  # Narrow to a single county_price_type and 1–2 states to keep the call small
  ptype  <- cand$county_price_type[1]
  states <- unique(cand$state_abbreviation)
  states <- states[!is.na(states)]
  states <- head(states, 1L)
  
  commodity <- cand$commodity[1]
  testthat::expect_true(nzchar(commodity))
  
  start <- Sys.Date() - 10
  end   <- Sys.Date()
  
  # ---- Daily -----------------------------------------------------------------
  res_d <- get_dtn_county_average(
    commodity          = commodity,
    start_date         = start,
    end_date           = end,
    time_scale         = "Daily",
    county_price_type  = ptype,
    fields             = c("Open","High","Low","Close"),
    state_abbreviation = states
  )
  
  testthat::expect_s3_class(res_d, "data.table")
  testthat::expect_true(nrow(res_d) >= 1)
  testthat::expect_true(all(c("open","high","low","close") %in% tolower(names(res_d))))
  testthat::expect_true(all(is.finite(res_d$close)))
  testthat::expect_identical(unique(res_d$time_scale), "daily")
  
  # ---- Weekly ----------------------------------------------------------------
  res_w <- get_dtn_county_average(
    commodity          = commodity,
    start_date         = start,
    end_date           = end,
    time_scale         = "Weekly",
    county_price_type  = ptype,
    fields             = c("Close"),
    state_abbreviation = states
  )
  
  testthat::expect_s3_class(res_w, "data.table")
  testthat::expect_true(nrow(res_w) >= 1)
  testthat::expect_true(all(c("close") %in% tolower(names(res_w))))
  testthat::expect_identical(unique(res_w$time_scale), "weekly")
  # Weekly dates should be Mondays (unless you later add holiday logic)
  wday <- as.POSIXlt(res_w$date)$wday
  testthat::expect_true(all(wday == 1L))
  
  # ---- Monthly ---------------------------------------------------------------
  res_m <- get_dtn_county_average(
    commodity          = commodity,
    start_date         = start - 35,  # include at least two months
    end_date           = end,
    time_scale         = "Monthly",
    county_price_type  = ptype,
    fields             = c("Close"),
    state_abbreviation = states
  )
  
  testthat::expect_s3_class(res_m, "data.table")
  testthat::expect_true(nrow(res_m) >= 1)
  testthat::expect_true(all(c("close") %in% tolower(names(res_m))))
  testthat::expect_identical(unique(res_m$time_scale), "monthly")
  
  # Monthly should use first weekday of each month in range (no weekends)
  testthat::expect_false(any(is_weekend(res_m$date)))
})
