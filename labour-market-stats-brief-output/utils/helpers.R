# helpers.R — shared formatters , utility functions

suppressPackageStartupMessages({
  library(dplyr)
  library(lubridate)
  library(DBI)
  library(RPostgres)
  library(tibble)
})

# auto-detect manual_month from database (latest LFS period + 2 months)
auto_detect_manual_month <- function() {
  tryCatch({
    conn <- DBI::dbConnect(RPostgres::Postgres())
    on.exit(try(DBI::dbDisconnect(conn), silent = TRUE))

    res <- DBI::dbGetQuery(conn, 'SELECT DISTINCT time_period FROM "ons"."labour_market__age_group"')
    if (nrow(res) == 0) return(NULL)

    month_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)

    parse_end <- function(label) {
      mons <- regmatches(label, gregexpr("[A-Za-z]{3}", label))[[1]]
      yrs  <- regmatches(label, gregexpr("[0-9]{4}", label))[[1]]
      if (length(mons) < 2 || length(yrs) < 1) return(as.Date(NA))
      end_m <- month_map[tolower(mons[2])]
      yr <- suppressWarnings(as.integer(yrs[1]))
      if (is.na(end_m) || is.na(yr)) return(as.Date(NA))
      as.Date(sprintf("%04d-%02d-01", yr, end_m))
    }

    ends <- as.Date(vapply(res$time_period, parse_end, as.Date(NA)), origin = "1970-01-01")
    if (all(is.na(ends))) return(NULL)

    latest_end <- max(ends, na.rm = TRUE)
    # manual_month = LFS end + 2 months (convention: anchor month of the release)
    anchor <- latest_end %m+% months(2)
    tolower(paste0(format(anchor, "%b"), format(anchor, "%Y")))
  }, error = function(e) {
    warning("auto_detect_manual_month failed: ", e$message, "; using Sys.Date fallback")
    NULL
  })
}

# date/period helpers

# e manual month string like "feb2026" -> Date 2026-02-01
parse_manual_month <- function(mm) {
  if (is.null(mm) || is.na(mm) || !nzchar(mm)) return(NULL)
  mm <- tolower(trimws(mm))
  mon_str <- substr(mm, 1, 3)
  yr_str  <- sub("^[a-z]+", "", mm)
  mon <- match(mon_str, tolower(month.abb))
  yr  <- suppressWarnings(as.integer(yr_str))
  if (is.na(mon) || is.na(yr)) return(NULL)
  as.Date(sprintf("%04d-%02d-01", yr, mon))
}

#  LFS 3-month label: end_date -> "Oct-Dec 2025"
make_lfs_label <- function(end_date) {
  if (is.null(end_date) || is.na(end_date)) return("")
  end_date <- as.Date(end_date)
  start_date <- end_date %m-% months(2)
  sprintf("%s-%s %s", format(start_date, "%b"), format(end_date, "%b"), format(end_date, "%Y"))
}

# Make LFS long-form narrative label: end_date -> "October 2025 to December 2025"
lfs_label_narrative <- function(end_date) {
  if (is.null(end_date) || is.na(end_date)) return("")
  end_date <- as.Date(end_date)
  start_date <- end_date %m-% months(2)
  paste0(format(start_date, "%B %Y"), " to ", format(end_date, "%B %Y"))
}

# payroll/sector format: "July 2025"
make_payroll_label <- function(date) {
  format(date, "%B %Y")
}

# sheet 13/15 format: "2000-01-01"
make_ymd_label <- function(date) {
  format(date, "%Y-%m-%d")
}

# cpi/hr1 format
make_datetime_label <- function(date) {
  paste0(format(date, "%Y-%m-%d"), " 00:00:00")
}

# lookup: get value by dataset_identifier_code and time_period
val_by_code <- function(pg_data, code, period_label) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  match_row <- pg_data %>%
    filter(dataset_identifier_code == code, trimws(time_period) == trimws(period_label))
  if (nrow(match_row) == 0) return(NA_real_)
  suppressWarnings(as.numeric(match_row$value[1]))
}

# num formatters

#  to 1 decimal place (or more if rounds to zero)
fmt_one_dec <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("\u2014")
  if (x == 0) return(format(0, nsmall = 1, trim = TRUE))
  for (d in 1:4) {
    vr <- round(x, d)
    if (vr != 0) return(format(vr, nsmall = d, trim = TRUE))
  }
  format(round(x, 4), nsmall = 4, trim = TRUE)
}

#  as percentage: 5.1 -> "5.1%"
fmt_pct <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("\u2014")
  paste0(fmt_one_dec(x), "%")
}

#  as percentge points: 0.3 -> "0.3 percentage points"
fmt_pp <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("\u2014")
  paste0(fmt_one_dec(abs(x)), " percentage points")
}

# Direction word: positive -> up_word, negative -> down_word, zero -> "unchanged"
fmt_dir <- function(x, up_word = "up", down_word = "down") {
  if (is.na(x)) return("")
  if (x > 0) up_word
  else if (x < 0) down_word
  else "unchanged at"
}

#  integer with comma separators
fmt_int <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("\u2014")
  format(round(x), big.mark = ",", scientific = FALSE)
}

fmt_mill <- function(x) {
  ifelse(is.na(x), "\u2014", paste0(format(round(x, 1), nsmall = 1)))
}

# signed formatting functions (used by word output and sheets)

format_int <- function(x) {
  if (is.na(x)) return(NA_character_)
  v <- round(as.numeric(x), 0)
  if (is.na(v)) return(NA_character_)
  if (v == 0) return("0")
  paste0(ifelse(v > 0, "+", "-"), format(abs(v), big.mark = ","))
}

format_pp <- function(x) {
  if (is.na(x)) return(NA_character_)
  v <- as.numeric(x)
  if (is.na(v)) return(NA_character_)
  if (v == 0) return("0pp")
  for (d in 1:4) {
    vr <- round(v, d)
    if (vr != 0) return(paste0(ifelse(vr > 0, "+", "-"), format(abs(vr), nsmall = d), "pp"))
  }
  paste0(ifelse(v > 0, "+", "-"), format(abs(round(v, 4)), nsmall = 4), "pp")
}

format_pct1 <- function(x) {
  if (is.na(x)) return(NA_character_)
  v <- as.numeric(x)
  if (is.na(v)) return(NA_character_)
  if (v == 0) return("0%")
  for (d in 1:4) {
    vr <- round(v, d)
    if (vr != 0) return(paste0(ifelse(vr > 0, "+", "-"), format(abs(vr), nsmall = d), "%"))
  }
  paste0(ifelse(v > 0, "+", "-"), format(abs(round(v, 4)), nsmall = 4), "%")
}

format_gbp_signed0 <- function(x) {
  if (is.na(x)) return(NA_character_)
  v <- round(as.numeric(x), 0)
  if (is.na(v)) return(NA_character_)
  if (v == 0) return("\u00A30")
  paste0(ifelse(v > 0, "+\u00A3", "-\u00A3"), format(abs(v), big.mark = ","))
}

# unsigned formatting functions

format_int_unsigned <- function(x) {
  if (is.na(x)) return(NA_character_)
  v <- round(abs(as.numeric(x)), 0)
  if (is.na(v)) return(NA_character_)
  format(v, big.mark = ",")
}

format_pct1_unsigned <- function(x) {
  if (is.na(x)) return(NA_character_)
  v <- abs(as.numeric(x))
  if (is.na(v)) return(NA_character_)
  if (v == 0) return("0.0%")
  for (d in 1:4) {
    vr <- round(v, d)
    if (vr != 0) return(paste0(format(vr, nsmall = d), "%"))
  }
  paste0(format(round(v, 4), nsmall = 4), "%")
}