# unemployment by age module
# table: ons.labour_market_unemployment

suppressPackageStartupMessages({
  library(dplyr)
  library(tibble)
  library(DBI)
  library(RPostgres)
  library(lubridate)
})

# fetch

fetch_unemployment_by_age <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())

  tryCatch({
    query <- 'SELECT age_group, duration, value_type, dataset_identifier_code, time_period, value
              FROM "ons"."labour_market_unemployment"'
    res <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(res)
  },
  error = function(e) {
    warning("Failed to fetch unemployment by age data: ", e$message)
    tibble::tibble(
      age_group = character(),
      duration = character(),
      value_type = character(),
      dataset_identifier_code = character(),
      time_period = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}

# helpers

parse_lfs_period_to_end_date <- function(label) {
  # "mar-may 1992" -> 1992-05-01
  if (is.na(label) || !nzchar(label)) return(NA)
  label <- trimws(label)
  mons <- regmatches(label, gregexpr("[A-Za-z]{3}", label))[[1]]
  yrs  <- regmatches(label, gregexpr("[0-9]{4}", label))[[1]]
  if (length(mons) < 2 || length(yrs) < 1) return(NA)
  end_mon <- tolower(mons[2])
  yyyy <- suppressWarnings(as.integer(yrs[1]))
  mon_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)
  mm <- mon_map[[end_mon]]
  if (is.null(mm) || is.na(yyyy)) return(NA)
  as.Date(sprintf("%04d-%02d-01", yyyy, mm))
}

# compute

compute_unemployment_by_age <- function(df, manual_mm) {
  # manual_mm unused here; this module just exposes latest period slices
  if (is.null(df) || nrow(df) == 0) {
    return(list(period = NA_character_, level = tibble(), rate = tibble()))
  }

  df2 <- df %>%
    mutate(
      value = suppressWarnings(as.numeric(value)),
      end_date = as.Date(vapply(time_period, parse_lfs_period_to_end_date, as.Date(NA)))
    ) %>%
    filter(!is.na(end_date))

  if (nrow(df2) == 0) return(list(period = NA_character_, level = tibble(), rate = tibble()))

  latest_date <- max(df2$end_date, na.rm = TRUE)
  latest_label <- df2 %>% filter(end_date == latest_date) %>% slice(1) %>% pull(time_period) %>% as.character() %>% trimws()

  latest <- df2 %>% filter(end_date == latest_date)

  # try to keep "all" duration if present, otherwise keep everything
  if ("All" %in% latest$duration) latest <- latest %>% filter(duration == "All")

  level_tbl <- latest %>%
    filter(tolower(value_type) == "level") %>%
    select(age_group, duration, dataset_identifier_code, value) %>%
    arrange(desc(value))

  rate_tbl <- latest %>%
    filter(grepl("rate", tolower(value_type))) %>%
    select(age_group, duration, dataset_identifier_code, value) %>%
    arrange(desc(value))

  list(
    period = latest_label,
    level = level_tbl,
    rate = rate_tbl
  )
}

# calculate

calculate_unemployment_by_age <- function(manual_mm) {
  df <- fetch_unemployment_by_age()
  compute_unemployment_by_age(df, manual_mm)
}
