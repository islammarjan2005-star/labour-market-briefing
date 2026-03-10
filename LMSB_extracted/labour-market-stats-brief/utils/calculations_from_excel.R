# calculations_from_excel.R
# Fresh rewrite — reads uploaded ONS Excel files (A01, HR1, X09, RTISA)
# and produces output variables consumed by manual_word_output.R,
# summary.R, and top_ten_stats.R.
#
# Key design decisions:
#   - Sheet names as STRINGS (not integer positions) to avoid readxl indexing bugs
#   - Row lookup by period label text, not hardcoded row numbers
#   - Column positions verified against actual Feb 2026 ONS files

suppressPackageStartupMessages({
  library(readxl)
  library(lubridate)
})

# load helpers if not already loaded
if (!exists("parse_manual_month", inherits = TRUE)) {
  source("utils/helpers.R")
}

# ============================================================================
# INTERNAL HELPERS
# ============================================================================

# Safe read: returns empty data.frame on failure
.read_sheet <- function(path, sheet) {
  tryCatch(
    suppressMessages(readxl::read_excel(path, sheet = sheet, col_names = FALSE)),
    error = function(e) {
      warning("Failed to read sheet '", sheet, "' from ", basename(path), ": ", e$message)
      data.frame()
    }
  )
}

# Find row index where column 1 matches a label (trimmed, case-insensitive)
.find_row <- function(tbl, label) {
  if (nrow(tbl) == 0 || ncol(tbl) == 0) return(NA_integer_)
  col1 <- trimws(as.character(tbl[[1]]))
  label <- trimws(label)
  idx <- which(tolower(col1) == tolower(label))
  if (length(idx) == 0) return(NA_integer_)
  idx[1]
}

# Extract numeric value at [row, col] — strips non-numeric chars
.cell_num <- function(tbl, row, col) {
  if (is.na(row) || row < 1 || row > nrow(tbl) || col > ncol(tbl)) return(NA_real_)
  x <- as.character(tbl[[col]][row])
  suppressWarnings(as.numeric(gsub("[^0-9.eE+-]", "", x)))
}

# Build LFS 3-month label: "Oct-Dec 2025" from end date 2025-12-01
.lfs_label <- function(end_date) {
  start_date <- end_date %m-% months(2)
  sprintf("%s-%s %s", format(start_date, "%b"), format(end_date, "%b"), format(end_date, "%Y"))
}

# Detect dates from mixed column (Excel serial numbers, datetimes, text)
.detect_dates <- function(x) {
  if (inherits(x, "Date")) return(floor_date(as.Date(x), "month"))
  if (inherits(x, c("POSIXct", "POSIXt"))) return(floor_date(as.Date(x), "month"))
  s <- as.character(x)
  num <- suppressWarnings(as.numeric(s))
  is_num <- !is.na(num) & grepl("^[0-9]+\\.?[0-9]*$", s)
  out <- rep(as.Date(NA), length(s))
  if (any(is_num)) out[is_num] <- as.Date(num[is_num], origin = "1899-12-30")
  if (any(!is_num)) {
    out[!is_num] <- suppressWarnings(
      lubridate::parse_date_time(
        s[!is_num],
        orders = c("ymd", "mdy", "dmy", "bY", "BY", "Y b", "b Y", "Ym", "my")
      )
    )
  }
  floor_date(as.Date(out), "month")
}

# Lookup value in a date-indexed data.frame
.val_by_date <- function(df_m, df_v, target_date) {
  idx <- which(df_m == target_date)
  if (length(idx) == 0) return(NA_real_)
  df_v[idx[1]]
}

# Average of values at specified dates
.avg_by_dates <- function(df_m, df_v, target_dates) {
  vals <- vapply(target_dates, function(d) .val_by_date(df_m, df_v, d), numeric(1))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

# Safe last non-NA value from a numeric vector
.safe_last <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) == 0) return(NA_real_)
  x[length(x)]
}

# Find column index by searching header rows for an ONS dataset identifier code
# Returns the column index, or fallback_col if not found
.find_col_by_code <- function(tbl, code, fallback_col = NA_integer_, search_rows = 1:min(10, nrow(tbl))) {
  if (nrow(tbl) == 0 || ncol(tbl) == 0) return(fallback_col)
  for (r in search_rows) {
    for (c in seq_len(ncol(tbl))) {
      cell <- as.character(tbl[[c]][r])
      if (!is.na(cell) && grepl(code, cell, fixed = TRUE)) return(c)
    }
  }
  fallback_col
}

# ============================================================================
# MAIN FUNCTION
# ============================================================================

run_calculations_from_excel <- function(manual_month,
                                         file_a01 = NULL,
                                         file_hr1 = NULL,
                                         file_x09 = NULL,
                                         file_rtisa = NULL,
                                         target_env = globalenv()) {

  cm <- parse_manual_month(manual_month)      # e.g. 2026-02-01 for "feb2026"
  anchor_m <- cm %m-% months(2)               # e.g. 2025-12-01 (LFS end month)

  # Comparison period labels (LFS 3-month rolling)
  lfs_end_cur  <- anchor_m                     # Dec 2025
  lfs_end_q    <- anchor_m %m-% months(3)      # Sep 2025
  lfs_end_y    <- anchor_m %m-% months(12)     # Dec 2024
  lfs_end_covid <- as.Date("2020-02-01")       # Dec-Feb 2020
  lfs_end_elec  <- as.Date("2024-06-01")        # Apr-Jun 2024

  lab_cur   <- .lfs_label(lfs_end_cur)   # "Oct-Dec 2025"
  lab_q     <- .lfs_label(lfs_end_q)     # "Jul-Sep 2025"
  lab_y     <- .lfs_label(lfs_end_y)     # "Oct-Dec 2024"
  lab_covid <- .lfs_label(lfs_end_covid) # "Dec-Feb 2020"
  lab_elec  <- .lfs_label(lfs_end_elec)  # "Apr-Jun 2024"

  # ==========================================================================
  # A01 — Sheet "1": Main LFS summary
  # Columns: A=period text, D=emp16+ level, E=unemp16+ level,
  #   I=unemp rate 16+, O=inact level 16-64, Q=emp rate 16-64, S=inact rate 16-64
  # ==========================================================================

  tbl_1 <- if (!is.null(file_a01)) .read_sheet(file_a01, "1") else data.frame()

  .lfs_metric <- function(tbl, col, labels) {
    rows <- vapply(labels, function(l) .find_row(tbl, l), integer(1))
    vals <- vapply(seq_along(rows), function(i) .cell_num(tbl, rows[i], col), numeric(1))
    names(vals) <- c("cur", "q", "y", "covid", "elec")
    list(
      cur = vals["cur"],
      dq  = vals["cur"] - vals["q"],
      dy  = vals["cur"] - vals["y"],
      dc  = vals["cur"] - vals["covid"],
      de  = vals["cur"] - vals["elec"]
    )
  }

  all_labels <- c(lab_cur, lab_q, lab_y, lab_covid, lab_elec)

  m_emp16   <- .lfs_metric(tbl_1, 4,  all_labels)   # Col D: employment 16+ level
  m_emprt   <- .lfs_metric(tbl_1, 17, all_labels)   # Col Q: employment rate 16-64
  m_unemp16 <- .lfs_metric(tbl_1, 5,  all_labels)   # Col E: unemployment 16+ level
  m_unemprt <- .lfs_metric(tbl_1, 9,  all_labels)   # Col I: unemployment rate 16+
  m_inact   <- .lfs_metric(tbl_1, 15, all_labels)   # Col O: inactivity level 16-64
  m_inactrt <- .lfs_metric(tbl_1, 19, all_labels)   # Col S: inactivity rate 16-64

  # Assign LFS main metrics
  for (prefix in c("emp16", "emp_rt", "unemp16", "unemp_rt", "inact", "inact_rt")) {
    m <- switch(prefix,
      emp16 = m_emp16, emp_rt = m_emprt,
      unemp16 = m_unemp16, unemp_rt = m_unemprt,
      inact = m_inact, inact_rt = m_inactrt
    )
    assign(paste0(prefix, "_cur"), m$cur, envir = target_env)
    assign(paste0(prefix, "_dq"),  m$dq,  envir = target_env)
    assign(paste0(prefix, "_dy"),  m$dy,  envir = target_env)
    assign(paste0(prefix, "_dc"),  m$dc,  envir = target_env)
    assign(paste0(prefix, "_de"),  m$de,  envir = target_env)
  }

  # ==========================================================================
  # A01 — Sheet "2": Age breakdown (50-64 inactivity)
  # Columns: A=period, BD(col 56)=inact 50-64 level, BE(col 57)=inact 50-64 rate
  # ==========================================================================

  tbl_2 <- if (!is.null(file_a01)) .read_sheet(file_a01, "2") else data.frame()

  m_5064   <- .lfs_metric(tbl_2, 56, all_labels)  # Col BD: inactivity 50-64 level
  m_5064rt <- .lfs_metric(tbl_2, 57, all_labels)  # Col BE: inactivity 50-64 rate

  for (prefix in c("inact5064", "inact5064_rt")) {
    m <- if (prefix == "inact5064") m_5064 else m_5064rt
    assign(paste0(prefix, "_cur"), m$cur, envir = target_env)
    assign(paste0(prefix, "_dq"),  m$dq,  envir = target_env)
    assign(paste0(prefix, "_dy"),  m$dy,  envir = target_env)
    assign(paste0(prefix, "_dc"),  m$dc,  envir = target_env)
    assign(paste0(prefix, "_de"),  m$de,  envir = target_env)
  }

  # ==========================================================================
  # A01 — Sheet "10": Redundancy
  # Columns: A=period, B(col 2)=level, C(col 3)=rate per 1000
  # ==========================================================================

  tbl_10 <- if (!is.null(file_a01)) .read_sheet(file_a01, "10") else data.frame()

  m_redund <- .lfs_metric(tbl_10, 3, all_labels)  # Col C: rate per 1000
  m_redund_level <- .lfs_metric(tbl_10, 2, all_labels)  # Col B: level

  assign("redund_cur", m_redund$cur, envir = target_env)
  assign("redund_dq",  m_redund$dq,  envir = target_env)
  assign("redund_dy",  m_redund$dy,  envir = target_env)
  assign("redund_dc",  m_redund$dc,  envir = target_env)
  assign("redund_de",  m_redund$de,  envir = target_env)

  # ==========================================================================
  # A01 — Sheet "13" (STRING name): AWE Total Pay (nominal)
  # Columns: A=period text, B(col 2)=weekly £, D(col 4)=3-month avg % YoY
  # ==========================================================================

  tbl_13 <- if (!is.null(file_a01)) .read_sheet(file_a01, "13") else data.frame()

  if (nrow(tbl_13) > 0 && ncol(tbl_13) >= 4) {
    w13_dates <- .detect_dates(tbl_13[[1]])
    w13_weekly <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[2]]))))
    w13_pct    <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[4]]))))

    latest_wages <- .val_by_date(w13_dates, w13_pct, anchor_m)

    # Weekly £ change for dashboard (quarterly and other comparisons)
    win3 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
    prev3 <- c(anchor_m %m-% months(3), anchor_m %m-% months(4), anchor_m %m-% months(5))
    yago3 <- win3 %m-% months(12)
    covid3 <- as.Date(c("2019-12-01", "2020-01-01", "2020-02-01"))
    election3 <- as.Date(c("2024-04-01", "2024-05-01", "2024-06-01"))

    .wage_change <- function(a_months, b_months) {
      a <- .avg_by_dates(w13_dates, w13_weekly, a_months)
      b <- .avg_by_dates(w13_dates, w13_weekly, b_months)
      if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
    }

    wages_change_q <- .wage_change(win3, prev3)
    wages_change_y <- .wage_change(win3, yago3)
    wages_change_covid <- .wage_change(win3, covid3)
    wages_change_election <- .wage_change(win3, election3)

    # Quarterly pp change in YoY growth rate (for narrative)
    prev_q_pct <- .val_by_date(w13_dates, w13_pct, anchor_m %m-% months(3))
    wages_total_qchange <- if (!is.na(latest_wages) && !is.na(prev_q_pct)) latest_wages - prev_q_pct else NA_real_
  } else {
    latest_wages <- wages_change_q <- wages_change_y <- NA_real_
    wages_change_covid <- wages_change_election <- wages_total_qchange <- NA_real_
    win3 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
  }

  # ==========================================================================
  # A01 — Sheet "15" (STRING name): AWE Regular Pay (nominal)
  # Columns: A=period text, B(col 2)=weekly £, D(col 4)=3-month avg % YoY
  # ==========================================================================

  tbl_15 <- if (!is.null(file_a01)) .read_sheet(file_a01, "15") else data.frame()

  if (nrow(tbl_15) > 0 && ncol(tbl_15) >= 4) {
    w15_dates <- .detect_dates(tbl_15[[1]])
    w15_pct   <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_15[[4]]))))

    latest_regular_cash <- .val_by_date(w15_dates, w15_pct, anchor_m)

    prev_q_reg <- .val_by_date(w15_dates, w15_pct, anchor_m %m-% months(3))
    wages_reg_qchange <- if (!is.na(latest_regular_cash) && !is.na(prev_q_reg)) latest_regular_cash - prev_q_reg else NA_real_
  } else {
    latest_regular_cash <- NA_real_
    wages_reg_qchange <- NA_real_
  }

  # Public/private sector wages from A01 sheets 13 and 15
  # Dynamic column search by ONS code, with hardcoded fallbacks
  # Sheet 13 (total pay): KAC9 = public YoY%, KAC6 = private YoY%
  # Sheet 15 (regular pay): KAJ7 = public YoY%, KAJ4 = private YoY%

  if (nrow(tbl_13) > 0 && ncol(tbl_13) >= 10 && exists("w13_dates")) {
    col_pub_total  <- .find_col_by_code(tbl_13, "KAC9", fallback_col = 10L)
    col_priv_total <- .find_col_by_code(tbl_13, "KAC6", fallback_col = 7L)

    w13_pub_pct  <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[col_pub_total]]))))
    w13_priv_pct <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[col_priv_total]]))))

    wages_total_public  <- .val_by_date(w13_dates, w13_pub_pct, anchor_m)
    wages_total_private <- .val_by_date(w13_dates, w13_priv_pct, anchor_m)
  } else {
    wages_total_public  <- NA_real_
    wages_total_private <- NA_real_
  }

  if (nrow(tbl_15) > 0 && ncol(tbl_15) >= 10 && exists("w15_dates")) {
    col_pub_reg  <- .find_col_by_code(tbl_15, "KAJ7", fallback_col = 10L)
    col_priv_reg <- .find_col_by_code(tbl_15, "KAJ4", fallback_col = 7L)

    w15_pub_pct  <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_15[[col_pub_reg]]))))
    w15_priv_pct <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_15[[col_priv_reg]]))))

    wages_reg_public  <- .val_by_date(w15_dates, w15_pub_pct, anchor_m)
    wages_reg_private <- .val_by_date(w15_dates, w15_priv_pct, anchor_m)
  } else {
    wages_reg_public  <- NA_real_
    wages_reg_private <- NA_real_
  }

  assign("latest_wages",          latest_wages,          envir = target_env)
  assign("wages_change_q",        wages_change_q,        envir = target_env)
  assign("wages_change_y",        wages_change_y,        envir = target_env)
  assign("wages_change_covid",    wages_change_covid,    envir = target_env)
  assign("wages_change_election", wages_change_election, envir = target_env)
  assign("wages_total_public",    wages_total_public,    envir = target_env)
  assign("wages_total_private",   wages_total_private,   envir = target_env)
  assign("wages_total_qchange",   wages_total_qchange,   envir = target_env)
  assign("latest_regular_cash",   latest_regular_cash,   envir = target_env)
  assign("wages_reg_public",      wages_reg_public,      envir = target_env)
  assign("wages_reg_private",     wages_reg_private,     envir = target_env)
  assign("wages_reg_qchange",     wages_reg_qchange,     envir = target_env)

  # ==========================================================================
  # X09 — Sheet "AWE Real_CPI": Real wages (CPI-adjusted)
  # Columns: A(1)=datetime, B(2)=real AWE £, E(5)=total 3m avg % YoY,
  #   I(9)=regular 3m avg % YoY
  # ==========================================================================

  tbl_cpi <- if (!is.null(file_x09)) .read_sheet(file_x09, "AWE Real_CPI") else data.frame()

  if (nrow(tbl_cpi) > 0 && ncol(tbl_cpi) >= 9) {
    cpi_months <- .detect_dates(tbl_cpi[[1]])
    cpi_real   <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[2]]))))
    cpi_total  <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[5]]))))
    cpi_reg    <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[9]]))))

    latest_wages_cpi   <- .val_by_date(cpi_months, cpi_total, anchor_m)
    latest_regular_cpi <- .val_by_date(cpi_months, cpi_reg, anchor_m)

    # £ changes for dashboard
    .cpi_change <- function(a_months, b_months) {
      a <- .avg_by_dates(cpi_months, cpi_real, a_months)
      b <- .avg_by_dates(cpi_months, cpi_real, b_months)
      if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
    }

    prev3_cpi <- c(anchor_m %m-% months(3), anchor_m %m-% months(4), anchor_m %m-% months(5))
    yago3_cpi <- win3 %m-% months(12)
    covid3_cpi <- as.Date(c("2019-12-01", "2020-01-01", "2020-02-01"))
    election3_cpi <- as.Date(c("2024-04-01", "2024-05-01", "2024-06-01"))

    wages_cpi_change_q <- .cpi_change(win3, prev3_cpi)
    wages_cpi_change_y <- .cpi_change(win3, yago3_cpi)
    wages_cpi_change_covid <- .cpi_change(win3, covid3_cpi)
    wages_cpi_change_election <- .cpi_change(win3, election3_cpi)

    # vs Dec 2007 and vs pre-pandemic (Dec 2019-Feb 2020, matching DB path)
    dec2007_val <- .val_by_date(cpi_months, cpi_real, as.Date("2007-12-01"))
    cur_cpi_real <- .avg_by_dates(cpi_months, cpi_real, win3)
    wages_cpi_total_vs_dec2007 <- if (!is.na(cur_cpi_real) && !is.na(dec2007_val) && dec2007_val != 0) {
      ((cur_cpi_real - dec2007_val) / dec2007_val) * 100
    } else NA_real_

    pandemic3 <- as.Date(c("2019-12-01", "2020-01-01", "2020-02-01"))
    pandemic_avg <- .avg_by_dates(cpi_months, cpi_real, pandemic3)
    wages_cpi_total_vs_pandemic <- if (!is.na(cur_cpi_real) && !is.na(pandemic_avg) && pandemic_avg != 0) {
      ((cur_cpi_real - pandemic_avg) / pandemic_avg) * 100
    } else NA_real_
  } else {
    latest_wages_cpi <- latest_regular_cpi <- NA_real_
    wages_cpi_change_q <- wages_cpi_change_y <- wages_cpi_change_covid <- wages_cpi_change_election <- NA_real_
    wages_cpi_total_vs_dec2007 <- wages_cpi_total_vs_pandemic <- NA_real_
  }

  assign("latest_wages_cpi",           latest_wages_cpi,           envir = target_env)
  assign("latest_regular_cpi",         latest_regular_cpi,         envir = target_env)
  assign("wages_cpi_change_q",         wages_cpi_change_q,         envir = target_env)
  assign("wages_cpi_change_y",         wages_cpi_change_y,         envir = target_env)
  assign("wages_cpi_change_covid",     wages_cpi_change_covid,     envir = target_env)
  assign("wages_cpi_change_election",  wages_cpi_change_election,  envir = target_env)
  assign("wages_cpi_total_vs_dec2007", wages_cpi_total_vs_dec2007, envir = target_env)
  assign("wages_cpi_total_vs_pandemic", wages_cpi_total_vs_pandemic, envir = target_env)

  # ==========================================================================
  # A01 — Sheet "19": Vacancies
  # Columns: A(1)=period text (3-month rolling), C(3)=level in thousands
  # ==========================================================================

  tbl_19 <- if (!is.null(file_a01)) .read_sheet(file_a01, "19") else data.frame()

  # Vacancies: use LFS-aligned period for dashboard consistency
  vac_end <- lfs_end_cur  # e.g. Dec 2025 → "Oct-Dec 2025" for Feb 2026 release
  vac_lab_cur   <- .lfs_label(vac_end)
  vac_lab_q     <- .lfs_label(vac_end %m-% months(3))
  vac_lab_y     <- .lfs_label(vac_end %m-% months(12))
  vac_lab_covid <- "Jan-Mar 2020"
  vac_lab_elec  <- .lfs_label(as.Date("2024-06-01"))  # "Apr-Jun 2024"

  if (nrow(tbl_19) > 0 && ncol(tbl_19) >= 3) {
    r_cur   <- .find_row(tbl_19, vac_lab_cur)
    vac_cur <- .cell_num(tbl_19, r_cur, 3)
    vac_dq  <- vac_cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_q), 3)
    vac_dy  <- vac_cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_y), 3)
    vac_dc  <- vac_cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_covid), 3)
    vac_de  <- vac_cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_elec), 3)
  } else {
    vac_cur <- vac_dq <- vac_dy <- vac_dc <- vac_de <- NA_real_
  }

  assign("vac_cur", vac_cur, envir = target_env)
  assign("vac_dq",  vac_dq,  envir = target_env)
  assign("vac_dy",  vac_dy,  envir = target_env)
  assign("vac_dc",  vac_dc,  envir = target_env)
  assign("vac_de",  vac_de,  envir = target_env)
  assign("vac", list(cur = vac_cur, dq = vac_dq, dy = vac_dy,
                     dc = vac_dc, de = vac_de, end = vac_end), envir = target_env)

  # ==========================================================================
  # A01 — Sheet "18": Days lost to labour disputes
  # Columns: A(1)=period, B(2)=thousands (monthly)
  # ==========================================================================

  tbl_18 <- if (!is.null(file_a01)) .read_sheet(file_a01, "18") else data.frame()

  if (nrow(tbl_18) > 0 && ncol(tbl_18) >= 2) {
    # Strip revision markers [r], [p], [x] from text before date parsing
    dl_raw <- gsub("\\s*\\[.*?\\]\\s*", "", as.character(tbl_18[[1]]))
    dl_dates <- .detect_dates(dl_raw)
    dl_vals  <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_18[[2]]))))
    valid_idx <- which(!is.na(dl_dates) & !is.na(dl_vals))
    if (length(valid_idx) > 0) {
      last_idx <- valid_idx[length(valid_idx)]
      days_lost_cur   <- dl_vals[last_idx]
      days_lost_label <- format(dl_dates[last_idx], "%B %Y")
    } else {
      days_lost_cur <- NA_real_
      days_lost_label <- ""
    }
  } else {
    days_lost_cur <- NA_real_
    days_lost_label <- ""
  }

  assign("days_lost_cur",   days_lost_cur,   envir = target_env)
  assign("days_lost_label", days_lost_label, envir = target_env)

  # ==========================================================================
  # RTISA — Payrolled employees
  # Sheet "1. Payrolled employees (UK)": A(1)=date text, B(2)=raw count
  # ==========================================================================

  rtisa_pay <- if (!is.null(file_rtisa)) {
    .read_sheet(file_rtisa, "1. Payrolled employees (UK)")
  } else data.frame()

  if (nrow(rtisa_pay) > 0 && ncol(rtisa_pay) >= 2) {
    # Parse text dates like "January 2026"
    rtisa_text <- trimws(as.character(rtisa_pay[[1]]))
    rtisa_parsed <- suppressWarnings(lubridate::parse_date_time(rtisa_text, orders = c("B Y", "bY", "BY")))
    rtisa_months <- floor_date(as.Date(rtisa_parsed), "month")
    rtisa_vals <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_pay[[2]]))))

    # Build clean data frame
    pay_df <- data.frame(m = rtisa_months, v = rtisa_vals, stringsAsFactors = FALSE)
    pay_df <- pay_df[!is.na(pay_df$m) & !is.na(pay_df$v), ]
    pay_df <- pay_df[order(pay_df$m), ]

    # 3-month averages aligned to LFS quarter
    months_cur  <- c(cm %m-% months(4), cm %m-% months(3), cm %m-% months(2))
    months_prev <- c(cm %m-% months(7), cm %m-% months(6), cm %m-% months(5))
    months_yago <- months_cur %m-% months(12)

    pay_cur_raw <- .avg_by_dates(pay_df$m, pay_df$v, months_cur)
    pay_prev3   <- .avg_by_dates(pay_df$m, pay_df$v, months_prev)
    pay_yago3   <- .avg_by_dates(pay_df$m, pay_df$v, months_yago)

    payroll_cur <- if (!is.na(pay_cur_raw)) pay_cur_raw / 1000 else NA_real_
    payroll_dq  <- if (!is.na(pay_cur_raw) && !is.na(pay_prev3)) (pay_cur_raw - pay_prev3) / 1000 else NA_real_
    payroll_dy  <- if (!is.na(pay_cur_raw) && !is.na(pay_yago3)) (pay_cur_raw - pay_yago3) / 1000 else NA_real_

    covid_base <- .avg_by_dates(pay_df$m, pay_df$v, as.Date(c("2019-12-01", "2020-01-01", "2020-02-01")))
    payroll_dc <- if (!is.na(pay_cur_raw) && !is.na(covid_base)) (pay_cur_raw - covid_base) / 1000 else NA_real_

    elec_base <- .avg_by_dates(pay_df$m, pay_df$v, as.Date(c("2024-04-01", "2024-05-01", "2024-06-01")))
    payroll_de <- if (!is.na(payroll_cur) && !is.na(elec_base)) payroll_cur - (elec_base / 1000) else NA_real_

    # Flash (single latest month)
    flash_anchor <- anchor_m
    flash_val    <- .val_by_date(pay_df$m, pay_df$v, anchor_m)
    flash_prev_m <- .val_by_date(pay_df$m, pay_df$v, anchor_m %m-% months(1))
    flash_prev_y <- .val_by_date(pay_df$m, pay_df$v, anchor_m %m-% months(12))
    flash_elec   <- .val_by_date(pay_df$m, pay_df$v, as.Date("2024-06-01"))

    payroll_flash_cur <- if (!is.na(flash_val)) flash_val / 1e6 else NA_real_
    payroll_flash_dm  <- if (!is.na(flash_val) && !is.na(flash_prev_m)) (flash_val - flash_prev_m) / 1000 else NA_real_
    payroll_flash_dy  <- if (!is.na(flash_val) && !is.na(flash_prev_y)) (flash_val - flash_prev_y) / 1000 else NA_real_
    payroll_flash_de  <- if (!is.na(flash_val) && !is.na(flash_elec)) (flash_val - flash_elec) / 1000 else NA_real_
  } else {
    payroll_cur <- payroll_dq <- payroll_dy <- payroll_dc <- payroll_de <- NA_real_
    payroll_flash_cur <- payroll_flash_dm <- payroll_flash_dy <- payroll_flash_de <- NA_real_
    flash_anchor <- anchor_m
  }

  assign("payroll_cur", payroll_cur, envir = target_env)
  assign("payroll_dq",  payroll_dq,  envir = target_env)
  assign("payroll_dy",  payroll_dy,  envir = target_env)
  assign("payroll_dc",  payroll_dc,  envir = target_env)
  assign("payroll_de",  payroll_de,  envir = target_env)
  assign("payroll_flash_cur", payroll_flash_cur, envir = target_env)
  assign("payroll_flash_dm",  payroll_flash_dm,  envir = target_env)
  assign("payroll_flash_dy",  payroll_flash_dy,  envir = target_env)
  assign("payroll_flash_de",  payroll_flash_de,  envir = target_env)

  # ==========================================================================
  # RTISA — Sheet "23. Employees (Industry)": Sector payroll
  # Columns: A(1)=date, H(8)=retail, J(10)=hospitality, R(18)=health
  # ==========================================================================

  rtisa_sec <- if (!is.null(file_rtisa)) {
    .read_sheet(file_rtisa, "23. Employees (Industry)")
  } else data.frame()

  # Sector payroll uses cm-1 anchor (not cm-2 like LFS), matching DB path
  # RTISA sector data has less lag than LFS
  sec_anchor <- cm %m-% months(1)

  if (nrow(rtisa_sec) > 0 && ncol(rtisa_sec) >= 18) {
    sec_text <- trimws(as.character(rtisa_sec[[1]]))
    sec_parsed <- suppressWarnings(lubridate::parse_date_time(sec_text, orders = c("B Y", "bY", "BY")))
    sec_months <- floor_date(as.Date(sec_parsed), "month")

    sec_retail <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_sec[[8]]))))
    sec_hosp   <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_sec[[10]]))))
    sec_health <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_sec[[18]]))))

    .sector_full <- function(vals) {
      now   <- .val_by_date(sec_months, vals, sec_anchor)
      prev  <- .val_by_date(sec_months, vals, sec_anchor %m-% months(1))
      yago  <- .val_by_date(sec_months, vals, sec_anchor %m-% months(12))
      covid <- .val_by_date(sec_months, vals, as.Date("2020-02-01"))
      elec  <- .val_by_date(sec_months, vals, as.Date("2024-06-01"))
      list(
        cur = if (!is.na(now)) now / 1000 else NA_real_,
        dm  = if (!is.na(now) && !is.na(prev)) (now - prev) / 1000 else NA_real_,
        dy  = if (!is.na(now) && !is.na(yago)) (now - yago) / 1000 else NA_real_,
        dc  = if (!is.na(now) && !is.na(covid)) (now - covid) / 1000 else NA_real_,
        de  = if (!is.na(now) && !is.na(elec)) (now - elec) / 1000 else NA_real_
      )
    }

    s_retail <- .sector_full(sec_retail)
    s_hosp   <- .sector_full(sec_hosp)
    s_health <- .sector_full(sec_health)
  } else {
    na_sector <- list(cur = NA_real_, dm = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_)
    s_retail <- s_hosp <- s_health <- na_sector
  }

  for (prefix in c("hosp", "retail", "health")) {
    s <- switch(prefix, hosp = s_hosp, retail = s_retail, health = s_health)
    assign(paste0(prefix, "_cur"), s$cur, envir = target_env)
    assign(paste0(prefix, "_dm"),  s$dm,  envir = target_env)
    assign(paste0(prefix, "_dy"),  s$dy,  envir = target_env)
    assign(paste0(prefix, "_dc"),  s$dc,  envir = target_env)
    assign(paste0(prefix, "_de"),  s$de,  envir = target_env)
  }

  # ==========================================================================
  # HR1 — Sheet "1a": Redundancy notifications
  # Columns: A(1)=datetime, M(13)=GB total
  # ==========================================================================

  hr1_tbl <- if (!is.null(file_hr1)) .read_sheet(file_hr1, "1a") else data.frame()

  if (nrow(hr1_tbl) > 0 && ncol(hr1_tbl) >= 13) {
    hr1_dates <- .detect_dates(hr1_tbl[[1]])
    hr1_vals  <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(hr1_tbl[[13]]))))

    valid_hr1 <- which(!is.na(hr1_dates) & !is.na(hr1_vals))
    if (length(valid_hr1) > 0) {
      last_hr1 <- valid_hr1[length(valid_hr1)]
      hr1_cur <- hr1_vals[last_hr1]
      hr1_month_label <- format(hr1_dates[last_hr1], "%B %Y")

      # Month-on-month change
      prev_hr1 <- if (length(valid_hr1) >= 2) {
        hr1_vals[valid_hr1[length(valid_hr1) - 1]]
      } else NA_real_
      hr1_dm <- if (!is.na(hr1_cur) && !is.na(prev_hr1)) hr1_cur - prev_hr1 else NA_real_

      # Year-on-year, vs COVID baseline, vs election
      hr1_cur_date <- hr1_dates[last_hr1]
      hr1_yago  <- .val_by_date(hr1_dates, hr1_vals, hr1_cur_date %m-% months(12))
      hr1_covid <- .val_by_date(hr1_dates, hr1_vals, as.Date("2020-02-01"))
      hr1_elec  <- .val_by_date(hr1_dates, hr1_vals, as.Date("2024-06-01"))

      hr1_dy <- if (!is.na(hr1_cur) && !is.na(hr1_yago)) hr1_cur - hr1_yago else NA_real_
      hr1_dc <- if (!is.na(hr1_cur) && !is.na(hr1_covid)) hr1_cur - hr1_covid else NA_real_
      hr1_de <- if (!is.na(hr1_cur) && !is.na(hr1_elec)) hr1_cur - hr1_elec else NA_real_
    } else {
      hr1_cur <- NA_real_
      hr1_dm <- hr1_dy <- hr1_dc <- hr1_de <- NA_real_
      hr1_month_label <- ""
    }
  } else {
    hr1_cur <- NA_real_
    hr1_dm <- hr1_dy <- hr1_dc <- hr1_de <- NA_real_
    hr1_month_label <- ""
  }

  assign("hr1_cur", hr1_cur, envir = target_env)
  assign("hr1_dm",  hr1_dm,  envir = target_env)
  assign("hr1_dy",  hr1_dy,  envir = target_env)
  assign("hr1_dc",  hr1_dc,  envir = target_env)
  assign("hr1_de",  hr1_de,  envir = target_env)
  assign("hr1_month_label", hr1_month_label, envir = target_env)

  # ==========================================================================
  # LABELS
  # ==========================================================================

  lfs_period_label <- lfs_label_narrative(lfs_end_cur)  # "October 2025 to December 2025"
  lfs_period_short_label <- make_lfs_label(lfs_end_cur) # "Oct-Dec 2025"
  vacancies_period_label <- lfs_label_narrative(vac_end)
  vacancies_period_short_label <- make_lfs_label(vac_end)
  payroll_flash_label_val <- format(flash_anchor, "%B %Y")
  sector_month_label <- format(sec_anchor, "%B %Y")

  assign("lfs_period_label",             lfs_period_label,             envir = target_env)
  assign("lfs_period_short_label",       lfs_period_short_label,       envir = target_env)
  assign("vacancies_period_label",       vacancies_period_label,       envir = target_env)
  assign("vacancies_period_short_label", vacancies_period_short_label, envir = target_env)
  assign("payroll_flash_label",          payroll_flash_label_val,      envir = target_env)
  assign("payroll_month_label",          format(anchor_m, "%B %Y"),    envir = target_env)
  assign("hr1_month_label",             hr1_month_label,              envir = target_env)
  assign("sector_month_label",          sector_month_label,           envir = target_env)
  assign("manual_month",                manual_month,                 envir = target_env)

  # Extra variables needed by downstream consumers
  assign("inact_driver_text", "", envir = target_env)

  # Payroll list object (needed by summary.R fallback)
  assign("payroll", list(
    cur = payroll_cur, dq = payroll_dq, dy = payroll_dy,
    dc = payroll_dc, de = payroll_de,
    flash_cur = payroll_flash_cur, flash_dm = payroll_flash_dm,
    flash_dy = payroll_flash_dy, flash_de = payroll_flash_de,
    flash_anchor = flash_anchor, anchor = anchor_m
  ), envir = target_env)

  # Wages nominal list object (needed by summary.R for period label)
  assign("wages_nom", list(
    total = list(cur = latest_wages, dq = wages_change_q,
                 dy = wages_change_y, dc = wages_change_covid,
                 de = wages_change_election,
                 public = wages_total_public, private = wages_total_private,
                 qchange = wages_total_qchange),
    regular = list(cur = latest_regular_cash,
                   public = wages_reg_public, private = wages_reg_private,
                   qchange = wages_reg_qchange),
    anchor = anchor_m
  ), envir = target_env)

  invisible(TRUE)
}
