
# hr1 module - redundancy notifications
# table: ons.labour_market__redundancies_region

# config

HR1_REGION <- "GB"

# fetch

fetch_hr1 <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  
  tryCatch({
    query <- 'SELECT time_period, region, value
              FROM "ons"."labour_market__redundancies_region"'
    result <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("Failed to fetch HR1 data: ", e$message)
    tibble::tibble(
      time_period = character(),
      region = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}

# lookup

val_hr1 <- function(pg_data, period_label) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  
  match_row <- pg_data %>%
    filter(
      region == HR1_REGION,
      startsWith(time_period, period_label)
    )
  
  if (nrow(match_row) == 0) return(NA_real_)
  
  suppressWarnings(as.numeric(match_row$value[1]))
}

# compute

compute_hr1 <- function(pg_data) {
  # find latest date for gb
  gb_data <- pg_data %>%
    filter(region == HR1_REGION) %>%
    mutate(parsed_date = as.Date(substr(time_period, 1, 10))) %>%
    filter(!is.na(parsed_date)) %>%
    arrange(desc(parsed_date))
  
  if (nrow(gb_data) == 0) {
    return(list(cur = NA_real_, dm = NA_real_, dy = NA_real_, 
                dc = NA_real_, de = NA_real_, anchor = NA))
  }
  
  # latest date
  anchor <- gb_data$parsed_date[1]
  cur <- suppressWarnings(as.numeric(gb_data$value[1]))
  
  # comparison periods
  prev_m <- anchor %m-% months(1)
  prev_y <- anchor %m-% months(12)
  covid_date <- as.Date("2020-02-01")
  election_date <- as.Date("2024-06-01")
  
  # format for lookup (just the date part for startswith)
  lab_m <- format(prev_m, "%Y-%m-%d")
  lab_y <- format(prev_y, "%Y-%m-%d")
  lab_c <- "2020-02-01"
  lab_e <- "2024-06-01"
  
  val_m <- val_hr1(pg_data, lab_m)
  val_y <- val_hr1(pg_data, lab_y)
  val_c <- val_hr1(pg_data, lab_c)
  val_e <- val_hr1(pg_data, lab_e)
  
  dm <- if (!is.na(cur) && !is.na(val_m)) cur - val_m else NA_real_
  dy <- if (!is.na(cur) && !is.na(val_y)) cur - val_y else NA_real_
  dc <- if (!is.na(cur) && !is.na(val_c)) cur - val_c else NA_real_
  de <- if (!is.na(cur) && !is.na(val_e)) cur - val_e else NA_real_
  
  list(
    cur = cur,
    dm = dm,
    dy = dy,
    dc = dc,
    de = de,
    anchor = anchor
  )
}

# calculate hr1

calculate_hr1 <- function() {
  pg_data <- fetch_hr1()
  compute_hr1(pg_data)
}
