# payroll employees by age module
# table: ons.labour_market_employees_age
# time period format: "january 2026" (monthly)

suppressPackageStartupMessages({
  library(dplyr)
  library(tibble)
  library(DBI)
  library(RPostgres)
  library(lubridate)
})

# fetch

fetch_payroll_by_age <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())

  tryCatch({
    query <- 'SELECT age_group, time_period, value
              FROM "ons"."labour_market_employees_age"'
    res <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(res)
  },
  error = function(e) {
    warning("Failed to fetch payroll by age data: ", e$message)
    tibble::tibble(
      age_group = character(),
      time_period = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}

# helpers

parse_month_label_to_date <- function(label) {
  # "january 2026" -> 2026-01-01
  if (is.na(label) || !nzchar(label)) return(NA)
  suppressWarnings(as.Date(paste0("01 ", trimws(label)), format = "%d %B %Y"))
}

# compute

compute_payroll_by_age <- function(df, manual_mm) {
  # manual_mm unused here; module exposes latest month slices
  if (is.null(df) || nrow(df) == 0) {
    return(list(period = NA_character_, data = tibble()))
  }

  df2 <- df %>%
    mutate(
      value = suppressWarnings(as.numeric(value)),
      month_date = as.Date(vapply(time_period, parse_month_label_to_date, as.Date(NA)))
    ) %>%
    filter(!is.na(month_date))

  if (nrow(df2) == 0) return(list(period = NA_character_, data = tibble()))

  latest_date <- max(df2$month_date, na.rm = TRUE)
  latest_label <- df2 %>% filter(month_date == latest_date) %>% slice(1) %>% pull(time_period) %>% as.character() %>% trimws()

  latest_tbl <- df2 %>%
    filter(month_date == latest_date) %>%
    select(age_group, value) %>%
    arrange(desc(value))

  list(
    period = latest_label,
    data = latest_tbl
  )
}

# calculate

calculate_payroll_by_age <- function(manual_mm) {
  df <- fetch_payroll_by_age()
  compute_payroll_by_age(df, manual_mm)
}
