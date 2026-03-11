
# days lost module - a01 sheet 18 (working days lost)



# config


DAYS_LOST_CODE <- "BBFW"


# fetch


fetch_days_lost <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  
  tryCatch({
    query <- 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__disputes"'
    result <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("Failed to fetch days lost: ", e$message)
    tibble::tibble(
      time_period = character(),
      dataset_identifier_code = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}


# compute


compute_days_lost <- function(pg_data, manual_mm) {
  cm <- parse_manual_month(manual_mm)
  
  # days lost uses 2 month lag from manual month
  anchor <- cm %m-% months(2)
  
  # format as "august 2025"
  lab_cur <- make_payroll_label(anchor)
  
  # use startswith to handle [p] or [r] suffixes
  match_row <- pg_data %>%
    filter(
      dataset_identifier_code == DAYS_LOST_CODE,
      startsWith(time_period, lab_cur)
    )
  
  if (nrow(match_row) == 0) {
    cur <- NA_real_
  } else {
    cur <- suppressWarnings(as.numeric(match_row$value[1]))
  }
  
  list(
    cur = cur,
    label = lab_cur,
    anchor = anchor
  )
}


# calculate days lost


calculate_days_lost <- function(manual_mm) {
  pg_data <- fetch_days_lost()
  compute_days_lost(pg_data, manual_mm)
}
