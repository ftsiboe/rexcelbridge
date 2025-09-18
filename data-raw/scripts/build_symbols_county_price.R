

county_average_spot_cash_price <- as.data.frame(
  readxl::read_excel("data-raw/supplementary_files/county_average_spot_cash_price.xlsx", 
                     sheet = "wheat_soft_red_winter"))

symbols_county_price <- tidyr::separate(
  county_average_spot_cash_price,"description",
  into = c("commodity","county_price_type","county_name","state_abbreviation"),
  remove = F,sep=", ")
symbols_county_price$data_source <- "ARPC using data from prophetX"
saveRDS(symbols_county_price,"data-raw/internal_datasets/symbols_county_price.rds")

