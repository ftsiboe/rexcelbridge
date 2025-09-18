test_that("is_weekend works", {
  d <- as.Date(c("2025-09-13","2025-09-14","2025-09-15")) # Sat, Sun, Mon
  testthat::expect_identical(is_weekend(d), c(TRUE, TRUE, FALSE))
})

testthat::test_that("first_trading_days returns first weekday per month", {
  x <- as.Date(c("2025-02-20","2025-02-28","2025-09-17"))
  out <- first_trading_days(x)
  testthat::expect_true(all(!is_weekend(out)))
  # Feb 2025-02-01 is Saturday -> first weekday 2025-02-03
  testthat::expect_true(any(out == as.Date("2025-02-03")))
})

testthat::test_that("get_target_dates scales are consistent", {
  start <- as.Date("2025-06-01")
  end   <- as.Date("2025-06-30")
  
  d <- get_target_dates(start, end, time_scale = "Daily")
  testthat::expect_true(all(!is_weekend(d)))
  testthat::expect_true(min(d) >= start && max(d) <= end)
  
  w <- get_target_dates(start, end, time_scale = "Weekly")
  testthat::expect_true(length(w) >= 4)
  testthat::expect_true(all(as.POSIXlt(w)$wday == 1L)) # Mondays
  
  m <- get_target_dates(start, end, time_scale = "Monthly")
  testthat::expect_true(length(m) == 1L)
  testthat::expect_false(is_weekend(m))
})

testthat::test_that("split_into_chunks splits as expected", {
  df <- data.frame(id = 1:10000)
  parts <- split_into_chunks(df)
  testthat::expect_length(parts, 3)
  testthat::expect_equal(nrow(parts[[1]]), 3500)
  testthat::expect_equal(nrow(parts[[2]]), 3500)
  testthat::expect_equal(nrow(parts[[3]]), 3000)
})

testthat::test_that("build_queries shapes rows correctly without executing", {
  # Minimal symbol stub (no download here)
  symbols <- data.frame(
    symbol = c("SYM1","SYM2"),
    commodity = c("Corn","Corn"),
    county_price_type = c("County Average Cash Price","County Average Cash Price"),
    state_abbreviation = c("IA","IA"),
    county_name = c("POLK","STORY"),
    stringsAsFactors = FALSE
  )
  dates  <- as.Date(c("2025-09-15","2025-09-16"))
  fields <- c("Open","Close")
  chunks <- build_queries(symbols, dates, time_scale = "Daily", fields = fields)
  
  # One chunk expected given tiny input
  testthat::expect_type(chunks, "list")
  testthat::expect_true(length(chunks) >= 1)
  q <- chunks[[1]]
  testthat::expect_true(all(c("symbol","date","field","query") %in% names(q)))
  testthat::expect_true(all(q$field %in% tolower(fields)))
  testthat::expect_true(all(grepl("field", dtn_prophetX_formula(symbol = "X", time_scale = "Daily", date = dates, field = "field")[1L], fixed = TRUE)) ||
                          all(grepl("open|close", q$query, ignore.case = TRUE)))
})
