
# sector payroll module - rtisa sheet 23

# config


SECTOR_CODES <- list(
  HOSPITALITY = "I",
  RETAIL = "G",
  HEALTH = "Q"
)


fetch_sector_payroll <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  
  tryCatch({
    query <- 'SELECT time_period, sic_section, value
FROM "ons"."labour_market__employees_industry"'
    result <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("Failed to fetch sector payroll: ", e$message)
    tibble::tibble(
      time_period = character(),
      sic_section = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}


# lookup


val_sector <- function(pg_data, sic_section, period_label) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  
  match_row <- pg_data %>%
    filter(
      sic_section == !!sic_section,
      trimws(time_period) == trimws(period_label)
    )
  
  if (nrow(match_row) == 0) return(NA_real_)
  
  suppressWarnings(as.numeric(match_row$value[1]))
}


# compute


compute_sector <- function(pg_data, manual_mm, sic_section) {
  cm <- parse_manual_month(manual_mm)
  
  anchor <- cm %m-% months(1)
  prev_m <- anchor %m-% months(1)
  prev_y <- anchor %m-% months(12)
  
  covid_date <- as.Date("2020-02-01")
  election_date <- as.Date("2024-06-01")
  
  lab_cur <- make_payroll_label(anchor)
  lab_m <- make_payroll_label(prev_m)
  lab_y <- make_payroll_label(prev_y)
  lab_c <- make_payroll_label(covid_date)
  lab_e <- make_payroll_label(election_date)
  
  cur <- val_sector(pg_data, sic_section, lab_cur)
  val_m <- val_sector(pg_data, sic_section, lab_m)
  val_y <- val_sector(pg_data, sic_section, lab_y)
  val_c <- val_sector(pg_data, sic_section, lab_c)
  val_e <- val_sector(pg_data, sic_section, lab_e)
  
  dm <- if (!is.na(cur) && !is.na(val_m)) (cur - val_m) / 1000 else NA_real_
  dy <- if (!is.na(cur) && !is.na(val_y)) (cur - val_y) / 1000 else NA_real_
  dc <- if (!is.na(cur) && !is.na(val_c)) (cur - val_c) / 1000 else NA_real_
  de <- if (!is.na(cur) && !is.na(val_e)) (cur - val_e) / 1000 else NA_real_
  
  list(
    cur = cur / 1000,
    dm = dm,
    dy = dy,
    dc = dc,
    de = de,
    anchor = anchor
  )
}


# calculate all sectors


calculate_sector_payroll <- function(manual_mm) {
  pg_data <- fetch_sector_payroll()
  
  list(
    hospitality = compute_sector(pg_data, manual_mm, SECTOR_CODES$HOSPITALITY),
    retail = compute_sector(pg_data, manual_mm, SECTOR_CODES$RETAIL),
    health = compute_sector(pg_data, manual_mm, SECTOR_CODES$HEALTH)
  )
}
