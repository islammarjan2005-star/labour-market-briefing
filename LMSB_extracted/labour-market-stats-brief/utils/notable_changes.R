# notable_changes.R
# Generic scanner: reads every sheet and every numeric column from the 4 core
# Excel files (A01, HR1, X09, RTISA) and surfaces the most statistically
# unusual period-on-period movements using z-scores against the column's own
# historical volatility.

library(readxl)
library(lubridate)

# ---- date helpers (self-contained) -------------------------------------------

.nc_detect_dates <- function(x) {
  if (inherits(x, "Date")) return(floor_date(as.Date(x), "month"))
  if (inherits(x, c("POSIXct", "POSIXt"))) return(floor_date(as.Date(x), "month"))
  s <- as.character(x)
  num <- suppressWarnings(as.numeric(s))
  is_num <- !is.na(num) & grepl("^[0-9]+\\.?[0-9]*$", s)
  out <- rep(as.Date(NA), length(s))
  if (any(is_num)) out[is_num] <- as.Date(num[is_num], origin = "1899-12-30")
  if (any(!is_num)) {
    out[!is_num] <- suppressWarnings(
      as.Date(lubridate::parse_date_time(
        s[!is_num],
        orders = c("ymd", "mdy", "dmy", "bY", "BY", "Y b", "b Y", "Ym", "my")
      ))
    )
  }
  floor_date(as.Date(out), "month")
}

.nc_parse_lfs_label <- function(label) {
  m <- regmatches(label, regexec("^([A-Za-z]{3})-([A-Za-z]{3}) (\\d{4})$", label))
  if (length(m[[1]]) != 4) return(as.Date(NA))
  end_mon <- match(tools::toTitleCase(m[[1]][3]), month.abb)
  yr <- as.integer(m[[1]][4])
  if (is.na(end_mon) || is.na(yr)) return(as.Date(NA))
  as.Date(sprintf("%04d-%02d-01", yr, end_mon))
}

# ---- column label extraction ------------------------------------------------

.nc_col_label <- function(tbl, col_idx, data_start_row) {
  search_rows <- seq_len(min(max(data_start_row - 1, 0), 10))
  labels <- character(0)
  for (r in rev(search_rows)) {
    cell <- as.character(tbl[[col_idx]][r])
    if (!is.na(cell) && nzchar(trimws(cell))) {
      if (!grepl("^[0-9.eE+-]+$", trimws(cell))) {
        labels <- c(trimws(cell), labels)
      }
    }
  }
  if (length(labels) > 0) {
    best <- labels[which.max(nchar(labels))]
    if (nchar(best) > 60) best <- paste0(substr(best, 1, 57), "...")
    return(best)
  }
  paste0("Column ", col_idx)
}

# ---- per-sheet scanner (z-score based) ---------------------------------------

.scan_sheet <- function(tbl, file_label, sheet_name) {
  if (nrow(tbl) < 5 || ncol(tbl) < 2) return(NULL)

  col1 <- as.character(tbl[[1]])

  # Strategy 1: LFS 3-month labels ("Oct-Dec 2025")
  lfs_pattern <- "^[A-Za-z]{3}-[A-Za-z]{3} \\d{4}$"
  lfs_hits <- grep(lfs_pattern, trimws(col1))

  if (length(lfs_hits) >= 5) {
    dates <- vapply(trimws(col1[lfs_hits]), .nc_parse_lfs_label, as.Date(NA))
    valid <- which(!is.na(dates))
    if (length(valid) < 6) return(NULL)  # need decent history
    ord <- order(dates[valid])
    sorted_dates <- dates[valid[ord]]
    data_rows <- lfs_hits[valid[ord]]
    data_start <- data_rows[1]
    period_type <- "quarterly"
  } else {
    # Strategy 2: monthly dates
    parsed <- .nc_detect_dates(col1)
    valid_idx <- which(!is.na(parsed))
    if (length(valid_idx) < 6) return(NULL)
    ord <- order(parsed[valid_idx])
    sorted_dates <- parsed[valid_idx[ord]]
    data_rows <- valid_idx[ord]
    data_start <- data_rows[1]
    period_type <- "monthly"
  }

  # Determine the 3-year lookback window to exclude COVID-era distortions
  latest_date <- sorted_dates[length(sorted_dates)]
  cutoff_date <- latest_date %m-% years(3)

  results <- list()

  for (ci in 2:ncol(tbl)) {
    raw_vals <- as.character(tbl[[ci]][data_rows])
    vals <- suppressWarnings(as.numeric(gsub("[^0-9.eE+-]", "", raw_vals)))

    non_na <- which(!is.na(vals))
    if (length(non_na) < 6) next

    # Compute period-on-period changes across the full valid series
    # Use consecutive non-NA values (handles gaps gracefully)
    changes <- diff(vals[non_na])
    change_dates <- sorted_dates[non_na[-1]]
    if (length(changes) < 4) next

    latest_change <- changes[length(changes)]
    if (latest_change == 0) next

    # --- Robust scoring: Modified Z-score using MAD on a 3-year window ---
    #
    # Why MAD over SD?
    # - Labour market data has fat tails (COVID, recessions, elections)
    # - SD is pulled by outliers; MAD is robust to them
    # - Factor of 1.4826 normalises MAD to be comparable to SD for Gaussian data
    #
    # Why 3-year window?
    # - Avoids COVID-era (2020-21) distortions inflating volatility
    # - Recent enough to capture current regime behaviour
    # - Falls back to all history if <4 points in the window

    in_window <- which(change_dates >= cutoff_date &
                       seq_along(changes) < length(changes))  # exclude latest
    if (length(in_window) < 4) {
      # Fallback: use all history except latest
      in_window <- seq_len(length(changes) - 1)
    }

    hist_changes <- changes[in_window]
    med_hist <- median(hist_changes, na.rm = TRUE)
    mad_hist <- mad(hist_changes, constant = 1.4826, na.rm = TRUE)

    if (is.na(mad_hist) || mad_hist == 0) {
      # Near-zero variability — any non-zero change is interesting
      z <- if (abs(latest_change) > 0) 3.0 else 0
    } else {
      z <- abs(latest_change - med_hist) / mad_hist
    }

    if (z < 1.5) next

    cur <- vals[non_na[length(non_na)]]
    prv <- vals[non_na[length(non_na) - 1]]

    label <- .nc_col_label(tbl, ci, data_start)

    # Determine if rate/percentage (small magnitude) vs level (large magnitude)
    is_rate <- all(abs(vals[non_na]) < 200)

    if (is_rate) {
      detail <- sprintf("%s%.1fpp", ifelse(latest_change > 0, "+", ""), latest_change)
    } else {
      pct <- if (prv != 0) latest_change / abs(prv) * 100 else NA
      pct_str <- if (!is.na(pct)) sprintf(", %s%.1f%%", ifelse(pct > 0, "+", ""), pct) else ""
      detail <- sprintf("%s%s%s",
                        ifelse(latest_change > 0, "+", ""),
                        formatC(latest_change, format = "f", digits = 1, big.mark = ","),
                        pct_str)
    }

    z_label <- sprintf("%.1f\u03c3", z)

    txt <- sprintf("%s \u2192 Sheet '%s', %s: %s \u2192 %s (%s) \u2014 %s %s change",
                   file_label, sheet_name, label,
                   formatC(prv, format = "f", digits = 1, big.mark = ","),
                   formatC(cur, format = "f", digits = 1, big.mark = ","),
                   detail, z_label, period_type)

    results[[length(results) + 1]] <- list(
      score = z,
      text  = txt,
      file  = file_label,
      sheet = sheet_name,
      col   = ci
    )
  }

  results
}

# ---- main entry point -------------------------------------------------------

generate_notable_changes <- function(file_a01 = NULL, file_hr1 = NULL,
                                     file_x09 = NULL, file_rtisa = NULL,
                                     n = 15) {

  files <- list()
  if (!is.null(file_a01))  files[["A01"]]   <- file_a01
  if (!is.null(file_hr1))  files[["HR1"]]   <- file_hr1
  if (!is.null(file_x09))  files[["X09"]]   <- file_x09
  if (!is.null(file_rtisa)) files[["RTISA"]] <- file_rtisa

  if (length(files) == 0) return(character(0))

  all_candidates <- list()

  for (file_label in names(files)) {
    path <- files[[file_label]]
    sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))

    for (sh in sheets) {
      tbl <- tryCatch(
        readxl::read_excel(path, sheet = sh, col_names = FALSE,
                           .name_repair = "minimal"),
        error = function(e) NULL
      )
      if (is.null(tbl)) next

      candidates <- .scan_sheet(tbl, file_label, sh)
      if (length(candidates) > 0) {
        all_candidates <- c(all_candidates, candidates)
      }
    }
  }

  if (length(all_candidates) == 0) return(character(0))

  # Sort by z-score descending
  scores <- vapply(all_candidates, function(x) x$score, numeric(1))
  ord <- order(scores, decreasing = TRUE)
  all_candidates <- all_candidates[ord]

  # Diversity cap: max 3 items per file+sheet to ensure breadth
  seen <- list()
  kept <- list()
  for (cand in all_candidates) {
    key <- paste0(cand$file, "::", cand$sheet)
    count <- if (is.null(seen[[key]])) 0L else seen[[key]]
    if (count >= 3L) next
    seen[[key]] <- count + 1L
    kept[[length(kept) + 1]] <- cand
    if (length(kept) >= n) break
  }

  vapply(kept, function(x) x$text, character(1))
}
