# workforce jobs (by industry) 
# table: ons.labour_market_workforce_jobs
# time period format: e.g. "mar 98 (r)" (quarterly, end-month label)

suppressPackageStartupMessages({
  library(dplyr)
  library(tibble)
  library(DBI)
  library(RPostgres)
  library(lubridate)
})

# fetch

fetch_workforce_jobs <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())

  tryCatch({
    query <- 'SELECT industry, sic_section, time_period, value
              FROM "ons"."labour_market_workforce_jobs"'
    res <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(res)
  },
  error = function(e) {
    warning("Failed to fetch workforce jobs data: ", e$message)
    tibble::tibble(
      industry = character(),
      sic_section = character(),
      time_period = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}

# helpers

parse_wfj_period_to_date <- function(x) {
  # "mar 98 (r)" -> 1998-03-01
  if (is.na(x) || !nzchar(x)) return(NA)
  x <- trimws(gsub("\\(.*\\)", "", x)) # strip "(r)" etc

  # expect "mon yy"
  parts <- strsplit(x, "\\s+")[[1]]
  if (length(parts) < 2) return(NA)

  mon <- tolower(substr(parts[1], 1, 3))
  yy  <- suppressWarnings(as.integer(parts[2]))

  mon_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)
  mm <- mon_map[[mon]]
  if (is.null(mm) || is.na(yy)) return(NA)

  # two-digit year -> assume 19xx for >= 50 else 20xx
  yyyy <- if (yy >= 50) 1900 + yy else 2000 + yy

  as.Date(sprintf("%04d-%02d-01", yyyy, mm))
}

latest_wfj_period_label <- function(df) {
  if (is.null(df) || nrow(df) == 0) return(NA_character_)
  d <- df %>%
    mutate(period_date = as.Date(vapply(time_period, parse_wfj_period_to_date, as.Date(NA)))) %>%
    filter(!is.na(period_date)) %>%
    arrange(desc(period_date))
  if (nrow(d) == 0) return(NA_character_)
  trimws(d$time_period[1])
}

# compute

compute_workforce_jobs <- function(df, manual_mm) {
  # manual_mm unused: workforce jobs isn't part of the lfs anchor logic yet.
  if (is.null(df) || nrow(df) == 0) {
    return(list(period = NA_character_, data = tibble()))
  }

  df2 <- df %>%
    mutate(
      value = suppressWarnings(as.numeric(value)),
      period_date = as.Date(vapply(time_period, parse_wfj_period_to_date, as.Date(NA)))
    ) %>%
    filter(!is.na(period_date))

  if (nrow(df2) == 0) return(list(period = NA_character_, data = tibble()))

  latest_date <- max(df2$period_date, na.rm = TRUE)
  latest_label <- df2 %>%
    filter(period_date == latest_date) %>%
    slice(1) %>%
    pull(time_period) %>%
    as.character() %>%
    trimws()

  latest_data <- df2 %>%
    filter(period_date == latest_date) %>%
    select(industry, sic_section, value) %>%
    arrange(desc(value))

  list(
    period = latest_label,
    data = latest_data
  )
}

# calculate

calculate_workforce_jobs <- function(manual_mm) {
  df <- fetch_workforce_jobs()
  compute_workforce_jobs(df, manual_mm)
}
