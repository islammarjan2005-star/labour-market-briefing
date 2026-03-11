# helpers.R — shared formatters and utility functions
# Sourced by: top_ten_stats.R, summary.R, calculations_from_excel.R,
#             excel_audit_workbook.R, app.R

suppressPackageStartupMessages({
  library(lubridate)
})

# --- Date/period helpers ---

# Parse manual month string like "feb2026" -> Date 2026-02-01
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

# Make LFS 3-month label: end_date -> "Oct-Dec 2025"
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

# --- Number formatters ---

# Format to 1 decimal place (or more if rounds to zero)
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

# Format as percentage: 5.1 -> "5.1%"
fmt_pct <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("\u2014")
  paste0(fmt_one_dec(x), "%")
}

# Format as percentage points: 0.3 -> "0.3 percentage points"
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

# Format integer with comma separators: 27600 -> "27,600"
fmt_int <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("\u2014")
  format(round(x), big.mark = ",", scientific = FALSE)
}
