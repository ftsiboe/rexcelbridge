#' Pull DTN county average cash prices and basis for a commodity
#'
#' Builds and executes DTN ProphetX queries for one or more counties and
#' returns a tidy table over a weekday-only date range.
#'
#' @param commodity Character (length 1 or vector). Substring matched (case-insensitive)
#'   against `symbols_county_price$commodity`.
#' @param start_date,end_date Date. Inclusive range. Defaults to the last 7 days.
#'   Weekends are dropped.
#' @param time_scale One of "Daily","Weekly","Monthly". Default "Daily".
#' @param county_price_type Character. One or more of:
#'   "County Average Spot Cash Price","County Average Cash Price",
#'   "County Average Spot Basis Price","County Average Basis".
#' @param fields Character vector of ProphetX fields to retrieve.
#' @param county_fip Character/numeric vector of 5-digit county FIPS.
#' @param state_abbreviation Character vector of state abbreviations (e.g., "IA","NE").
#' @param control a list of control parameters -> dtn_prophetX_query_limit Max rows per chunk.
#' @return A data.table with one row per (symbol, date). Columns include symbol metadata,
#'   `date`, `time_scale`, and wide OHLCV columns (`open`, `high`, `low`, `close`, `volume`, `openint`).
#' @import data.table
#' @importFrom stats setNames
#' @export
get_dtn_county_average <- function(
    commodity,
    start_date = Sys.Date()-7,
    end_date   = Sys.Date(),
    time_scale      = "Daily",
    county_price_type = "County Average Spot Cash Price",
    fields = c("Open","High","Low","Close","Volume","OpenInt"),
    county_fip = NULL,
    state_abbreviation  = NULL,
    control = rexcelbridge_controls()){
  
  if(!control$continuous_integration_session){
    on.exit(try(rb_kill_excel(), silent = TRUE), add = TRUE)  # ensure cleanup even on error
  }
  
  
  # ---- Allowed types ---------------------------------------------------------
  allowed_types <- c(
    "County Average Spot Cash Price",
    "County Average Cash Price",
    "County Average Spot Basis Price",
    "County Average Basis"
  )
  if (!all(county_price_type %in% allowed_types)) {
    stop("`county_price_type` must be one or more of: ",
         paste(shQuote(allowed_types), collapse = ", "), call. = FALSE)
  }
  
  # ---- Input checks -----------------------------------------------------------
  if (missing(commodity) || length(commodity) < 1)
    stop("`commodity` must be provided.", call. = FALSE)
  
  start_date <- as.Date(start_date); end_date <- as.Date(end_date)
  if (end_date < start_date)
    stop("`end_date` must be on/after `start_date`.", call. = FALSE)
  
  # normalize time_scale
  time_scale <- match.arg(tolower(time_scale), c("daily","weekly","monthly"))
  time_scale <- switch(time_scale, daily = "Daily", weekly = "Weekly", monthly = "Monthly")
  
  # normalize/validate fields (accept case-insensitive; canonicalize to TitleCase for PX)
  canonical_fields <- c("Open","High","Low","Close","Volume","OpenInt")
  norm_title <- function(x) {
    tx <- tolower(x)
    map <- setNames(canonical_fields, tolower(canonical_fields))
    if (!all(tx %in% names(map))) {
      bad <- setdiff(x, canonical_fields)
      stop("Unknown `fields`: ", paste(shQuote(bad), collapse = ", "),
           ". Allowed: ", paste(shQuote(canonical_fields), collapse = ", "),
           call. = FALSE)
    }
    unname(map[tx])
  }
  fields <- norm_title(fields)
  
  # ---- County universe --------------------------------------------------------
  county_fips <- get_county_universe(county_fip = county_fip, state_abbreviation = state_abbreviation)
  if (nrow(county_fips) == 0) return(data.table::as.data.table(NULL))
  
  # ---- Symbol filter ----------------------------------------------------------
  symbols <- symbols_county_price[
    grepl(toupper(commodity), toupper(symbols_county_price$commodity)) &
      symbols_county_price$county_price_type %in% county_price_type &
      toupper(symbols_county_price$state_abbreviation) %in% toupper(unique(county_fips$abbr)) &
      toupper(symbols_county_price$county_name) %in% toupper(unique(county_fips$county)),
  ]
  
  if (NROW(symbols) == 0) return(data.table::as.data.table(NULL))
  
  # ensure symbol is character (avoid factor surprises)
  if (is.factor(symbols$symbol)) symbols$symbol <- as.character(symbols$symbol)
  
  # map county_fips onto symbol rows (NA if no match)
  symbols$county_fips <- as.character(
    factor(toupper(symbols$county_name),
           levels  = toupper(county_fips$county),
           labels  = county_fips$fips)
  )
  
  # ---- Target dates -----------------------------------------------------------
  target_dates <- get_target_dates(start_date = start_date, end_date = end_date, time_scale = time_scale)
  if (length(target_dates) == 0) return(data.table::as.data.table(NULL))
  
  # ---- Build & split queries --------------------------------------------------
  res <- build_queries(symbols = symbols, target_dates = target_dates,
                       time_scale = time_scale, fields = fields,
                       control=control)
  
  # ---- Download ---------------------------------------------------------------
  if(!control$continuous_integration_session){
  res <- data.table::rbindlist(
    lapply(res, function(df){
      tryCatch({
        df$value <- as.numeric(as.character(
          rb_eval_single(df$query, visible = FALSE)$result
        ))
        df
      }, error = function(e) NULL)
    }),
    fill = TRUE
  )
  
  # filter unusable values
  res <- res[!is.na(value) & is.finite(value), ]
  
  }else{
   
    res <- data.frame()
    
  }
  
  if(nrow(res) == 0L){return(data.table::as.data.table(NULL))}
  
  # cast wide
  res <- res[, c(names(res)[! names(res) %in% "query"]), with = FALSE] |> 
    tidyr::spread(field, value)
  
  res$time_scale <- tolower(time_scale) 
  
  data.table::as.data.table(res)

}


#' County/state filtered county universe
#' @keywords internal
#' @noRd
#' @param county_fip Optional vector of 5-digit FIPS.
#' @param state_abbreviation Optional vector of state abbreviations.
get_county_universe <- function(county_fip = NULL, state_abbreviation = NULL) {
  county_fips <- as.data.frame(usmap::countypop[, c("fips", "abbr", "county")])
  county_fips$county <- gsub(" COUNTY", "", toupper(county_fips$county))
  if (!is.null(state_abbreviation)) county_fips <- county_fips[toupper(county_fips$abbr) %in% toupper(state_abbreviation), ]
  if (!is.null(county_fip)) county_fips <- county_fips[county_fips$fips %in% county_fip, ]
  unique(county_fips)
}

#' Weekend flag (locale-agnostic)
#' @keywords internal
#' @noRd
#' @param x Date vector.
#' @return Logical vector; TRUE for weekend.
is_weekend <- function(x) {
  x <- as.Date(x); w <- as.POSIXlt(x)$wday
  w %in% c(0, 6)
}

#' First trading (Mon-Fri) day for each month
#' @keywords internal
#' @noRd
#' @param x Date vector.
#' @return Sorted unique first trading days.
first_trading_days <- function(x) {
  x <- as.Date(x)
  months <- unique(format(x, "%Y-%m"))
  ftd <- vapply(months, function(m) {
    d1 <- as.Date(paste0(m, "-01"))
    days <- seq(d1, by = "day", length.out = 7)
    min(days[!is_weekend(days)])
  }, as.Date(NA))
  sort(as.Date(ftd))
}

#' Target dates by time scale (Daily/Weekly/Monthly)
#' @keywords internal
#' @noRd
#' @param start_date,end_date Date range.
#' @param time_scale One of "Daily","Weekly","Monthly".
#' @return Vector of dates.
get_target_dates <- function(
    start_date = Sys.Date()-7,
    end_date   = Sys.Date(),
    time_scale      = "Daily") {
  all_dates <- seq(as.Date(start_date), as.Date(end_date), by = "day")
  weekdays_only <- all_dates[!is_weekend(all_dates)]
  up <- toupper(time_scale)
  
  if (up == "DAILY")   return(weekdays_only)
  if (up == "WEEKLY")  return(weekdays_only[as.POSIXlt(weekdays_only)$wday == 1L]) # Mondays
  if (up == "MONTHLY") return(first_trading_days(weekdays_only))
  
  weekdays_only
}

#' Split a data.frame into fixed-size chunks
#' @keywords internal
#' @noRd
#' @param df data.frame/data.table.
#' @param control a list of control parameters -> dtn_prophetX_query_limit Max rows per chunk.
#' @return List of data.frames.
split_into_chunks <- function(
    df,
    control = rexcelbridge_controls()){
  n <- nrow(df); if (n == 0) return(list(df))
  idx <- ceiling(seq_len(n) / control$dtn_prophetX_query_limit)
  split(df, idx)
}

#' Build ProphetX query rows and split for batching
#' @keywords internal
#' @noRd
#' @param symbols Filtered symbol table.
#' @param target_dates Vector of query dates.
#' @param time_scale ProphetX time_scale ("Daily","Weekly","Monthly").
#' @param fields Character vector of fields.
#' @param control a list of control parameters -> dtn_prophetX_query_limit Max rows per chunk.
#' @return List of data.frames (chunks) with columns: symbol metadata, date, field, query.
build_queries <- function(symbols, target_dates, time_scale, fields,
                          control = rexcelbridge_controls()){
  base <- data.table::rbindlist(
    lapply(seq_len(nrow(symbols)), function(i) {
      data.frame(
        symbols[i, ],
        date  = target_dates,
        query = dtn_prophetX_formula(
          symbol = symbols$symbol[i],
          time_scale  = time_scale,
          date   = target_dates,
          field  = "field"
        ),
        stringsAsFactors = FALSE
      )
    }), fill = TRUE
  )
  
  long_queries <- data.table::rbindlist(
    lapply(fields, function(f) {
      df <- base
      df$query <- sub("field", f, df$query, fixed = TRUE)
      df$field <- tolower(f)
      df
    }),
    fill = TRUE
  )
  
  split_into_chunks(df=long_queries, control=control)
}
