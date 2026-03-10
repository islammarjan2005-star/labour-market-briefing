# calculations_from_excel.R
# reads uploaded ONS Excel files (A01, HR1, X09, RTISA) and produces
# the same output variables as calculations.R (database mode).
# mirrors the approach used in the original Rmd briefing.

suppressPackageStartupMessages({
  library(readxl)
  library(lubridate)
})

# load helpers if not already loaded
if (!exists("parse_manual_month", inherits = TRUE)) {
  source("utils/helpers.R")
}

# ============================================================================
# HELPER FUNCTIONS (from Rmd)
# ============================================================================

.get_series <- function(tbl, idx) {
  if (is.null(tbl) || !is.data.frame(tbl) || ncol(tbl) < idx) return(numeric(0))
  x <- gsub("[^0-9.-]", "", as.character(tbl[[idx]]))
  suppressWarnings(na.omit(as.numeric(x)))
}

.safe_last <- function(x, k = 1) {
  if (length(x) >= k) tail(x, k)[k] else NA_real_
}

.parse_cpi_dates <- function(x) {
  s <- as.character(x)
  as_num <- suppressWarnings(as.numeric(s))
  is_serial <- !is.na(as_num) & grepl("^[0-9]+$", s)
  out <- rep(as.Date(NA), length(s))
  if (any(is_serial)) out[is_serial] <- as.Date(as_num[is_serial], origin = "1899-12-30")
  if (any(!is_serial)) out[!is_serial] <- suppressWarnings(lubridate::mdy(s[!is_serial]))
  floor_date(out, "month")
}

.detect_dates <- function(x) {
  if (inherits(x, "Date")) return(floor_date(as.Date(x), "month"))
  if (inherits(x, c("POSIXct", "POSIXt"))) return(floor_date(as.Date(x), "month"))
  s <- as.character(x)
  num <- suppressWarnings(as.numeric(s))
  is_num <- !is.na(num) & grepl("^[0-9]+$", s)
  out <- rep(as.Date(NA), length(s))
  if (any(is_num)) out[is_num] <- as.Date(num[is_num], origin = "1899-12-30")
  if (any(!is_num)) {
    out[!is_num] <- suppressWarnings(
      lubridate::parse_date_time(
        s[!is_num],
        orders = c("ymd", "mdy", "dmy", "bY", "BY", "Y b", "b Y",
                    "Ym", "my", "%b-%Y", "%B-%Y", "%Y-%b", "%Y-%B")
      )
    )
  }
  floor_date(as.Date(out), "month")
}

.make_lfs_table_label <- function(end_date) {
  start_date <- end_date %m-% months(2)
  sprintf("%s-%s %s", format(start_date, "%b"), format(end_date, "%b"), format(end_date, "%Y"))
}

.val_by_label <- function(tbl, label, col_idx) {
  r <- which(trimws(as.character(tbl[[1]])) == label)[1]
  if (length(r) == 0 || is.na(r)) return(NA_real_)
  x <- gsub("[^0-9.-]", "", as.character(tbl[[col_idx]][r]))
  suppressWarnings(as.numeric(x))
}

.pick_vals <- function(month_vec, value_vec, keys) {
  idx <- match(keys, month_vec)
  if (any(is.na(idx))) return(rep(NA_real_, length(keys)))
  value_vec[idx]
}

.pick_avg <- function(month_vec, value_vec, keys) {
  v <- .pick_vals(month_vec, value_vec, keys)
  if (any(is.na(v))) return(NA_real_)
  mean(v)
}

.get_avg <- function(df, months_vec) {
  idx <- match(months_vec, df$m)
  if (any(is.na(idx))) return(NA_real_)
  mean(df$v[idx])
}

# safe read helper
.read_sheet <- function(path, sheet) {
  tryCatch(
    suppressMessages(readxl::read_excel(path, sheet = sheet, col_names = FALSE)),
    error = function(e) {
      warning("Failed to read sheet '", sheet, "': ", e$message)
      data.frame()
    }
  )
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

  cm <- parse_manual_month(manual_month)
  anchor_m <- cm %m-% months(2)

  # comparison periods
  COVID_LFS_LABEL <- "Dec-Feb 2020"
  ELECTION_LABEL <- "Apr-Jun 2024"

  # ==========================================================================
  # READ EXCEL SHEETS
  # ==========================================================================

  tbl_s2 <- if (!is.null(file_a01)) .read_sheet(file_a01, 2) else data.frame()
  tbl_10 <- if (!is.null(file_a01)) .read_sheet(file_a01, 10) else data.frame()
  tbl_13 <- if (!is.null(file_a01)) .read_sheet(file_a01, 13) else data.frame()
  tbl_15 <- if (!is.null(file_a01)) .read_sheet(file_a01, 15) else data.frame()
  tbl_18 <- if (!is.null(file_a01)) .read_sheet(file_a01, 18) else data.frame()
  tbl_19 <- if (!is.null(file_a01)) .read_sheet(file_a01, 19) else data.frame()

  tbl_cpi <- if (!is.null(file_x09)) .read_sheet(file_x09, "AWE Real_CPI") else data.frame()

  rtisa_s2  <- if (!is.null(file_rtisa)) .read_sheet(file_rtisa, 2) else data.frame()
  rtisa_s24 <- if (!is.null(file_rtisa)) .read_sheet(file_rtisa, 24) else data.frame()

  hr1_tbl <- if (!is.null(file_hr1)) .read_sheet(file_hr1, "1a") else data.frame()

  # ==========================================================================
  # LFS (A01 Sheet 2)
  # ==========================================================================

  .compute_lfs <- function(tbl, col_idx) {
    if (nrow(tbl) == 0) return(list(cur = NA_real_, dq = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_, end = anchor_m))
    end_cur <- cm %m-% months(2)
    end_q <- end_cur %m-% months(3)
    end_y <- end_cur %m-% months(12)

    cur <- .val_by_label(tbl, .make_lfs_table_label(end_cur), col_idx)
    dq <- cur - .val_by_label(tbl, .make_lfs_table_label(end_q), col_idx)
    dy <- cur - .val_by_label(tbl, .make_lfs_table_label(end_y), col_idx)
    dc <- cur - .val_by_label(tbl, COVID_LFS_LABEL, col_idx)
    de <- cur - .val_by_label(tbl, ELECTION_LABEL, col_idx)

    list(cur = cur, dq = dq, dy = dy, dc = dc, de = de, end = end_cur)
  }

  # Column indices for A01 Sheet 2
  COL_EMP16 <- 2; COL_UNEMP16 <- 4; COL_UNEMP_RT <- 5; COL_EMP_RT <- 11
  COL_INACT <- 16; COL_INACT_RT <- 17; COL_5064 <- 56; COL_5064_RT <- 57

  m_emp16   <- .compute_lfs(tbl_s2, COL_EMP16)
  m_emprt   <- .compute_lfs(tbl_s2, COL_EMP_RT)
  m_unemp16 <- .compute_lfs(tbl_s2, COL_UNEMP16)
  m_unemprt <- .compute_lfs(tbl_s2, COL_UNEMP_RT)
  m_inact   <- .compute_lfs(tbl_s2, COL_INACT)
  m_inactrt <- .compute_lfs(tbl_s2, COL_INACT_RT)
  m_5064    <- .compute_lfs(tbl_s2, COL_5064)
  m_5064rt  <- .compute_lfs(tbl_s2, COL_5064_RT)

  # assign to target env
  assign("emp16_cur",  m_emp16$cur,  envir = target_env)
  assign("emp16_dq",   m_emp16$dq,   envir = target_env)
  assign("emp16_dy",   m_emp16$dy,   envir = target_env)
  assign("emp16_dc",   m_emp16$dc,   envir = target_env)
  assign("emp16_de",   m_emp16$de,   envir = target_env)

  assign("emp_rt_cur", m_emprt$cur,  envir = target_env)
  assign("emp_rt_dq",  m_emprt$dq,   envir = target_env)
  assign("emp_rt_dy",  m_emprt$dy,   envir = target_env)
  assign("emp_rt_dc",  m_emprt$dc,   envir = target_env)
  assign("emp_rt_de",  m_emprt$de,   envir = target_env)

  assign("unemp16_cur",  m_unemp16$cur,  envir = target_env)
  assign("unemp16_dq",   m_unemp16$dq,   envir = target_env)
  assign("unemp16_dy",   m_unemp16$dy,   envir = target_env)
  assign("unemp16_dc",   m_unemp16$dc,   envir = target_env)
  assign("unemp16_de",   m_unemp16$de,   envir = target_env)

  assign("unemp_rt_cur", m_unemprt$cur,  envir = target_env)
  assign("unemp_rt_dq",  m_unemprt$dq,   envir = target_env)
  assign("unemp_rt_dy",  m_unemprt$dy,   envir = target_env)
  assign("unemp_rt_dc",  m_unemprt$dc,   envir = target_env)
  assign("unemp_rt_de",  m_unemprt$de,   envir = target_env)

  assign("inact_cur",    m_inact$cur,   envir = target_env)
  assign("inact_dq",     m_inact$dq,    envir = target_env)
  assign("inact_dy",     m_inact$dy,    envir = target_env)
  assign("inact_dc",     m_inact$dc,    envir = target_env)
  assign("inact_de",     m_inact$de,    envir = target_env)

  assign("inact_rt_cur", m_inactrt$cur, envir = target_env)
  assign("inact_rt_dq",  m_inactrt$dq,  envir = target_env)
  assign("inact_rt_dy",  m_inactrt$dy,  envir = target_env)
  assign("inact_rt_dc",  m_inactrt$dc,  envir = target_env)
  assign("inact_rt_de",  m_inactrt$de,  envir = target_env)

  assign("inact5064_cur",    m_5064$cur,   envir = target_env)
  assign("inact5064_dq",     m_5064$dq,    envir = target_env)
  assign("inact5064_dy",     m_5064$dy,    envir = target_env)
  assign("inact5064_dc",     m_5064$dc,    envir = target_env)
  assign("inact5064_de",     m_5064$de,    envir = target_env)

  assign("inact5064_rt_cur", m_5064rt$cur, envir = target_env)
  assign("inact5064_rt_dq",  m_5064rt$dq,  envir = target_env)
  assign("inact5064_rt_dy",  m_5064rt$dy,  envir = target_env)
  assign("inact5064_rt_dc",  m_5064rt$dc,  envir = target_env)
  assign("inact5064_rt_de",  m_5064rt$de,  envir = target_env)

  # ==========================================================================
  # VACANCIES (A01 Sheet 19)
  # ==========================================================================

  if (nrow(tbl_19) > 0) {
    end_vac <- cm %m-% months(1)
    vac_cur <- .val_by_label(tbl_19, .make_lfs_table_label(end_vac), 3)
    vac_dq  <- vac_cur - .val_by_label(tbl_19, .make_lfs_table_label(end_vac %m-% months(3)), 3)
    vac_dy  <- vac_cur - .val_by_label(tbl_19, .make_lfs_table_label(end_vac %m-% months(12)), 3)
    vac_dc  <- vac_cur - .val_by_label(tbl_19, "Jan-Mar 2020", 3)
    vac_de  <- vac_cur - .val_by_label(tbl_19, ELECTION_LABEL, 3)
  } else {
    vac_cur <- vac_dq <- vac_dy <- vac_dc <- vac_de <- NA_real_
    end_vac <- anchor_m
  }

  assign("vac_cur", vac_cur, envir = target_env)
  assign("vac_dq",  vac_dq,  envir = target_env)
  assign("vac_dy",  vac_dy,  envir = target_env)
  assign("vac_dc",  vac_dc,  envir = target_env)
  assign("vac_de",  vac_de,  envir = target_env)
  assign("vac", list(cur = vac_cur, dq = vac_dq, dy = vac_dy, dc = vac_dc, de = vac_de, end = end_vac), envir = target_env)

  # ==========================================================================
  # PAYROLL (RTISA Sheet 2)
  # ==========================================================================

  if (nrow(rtisa_s2) > 0 && ncol(rtisa_s2) >= 2) {
    rtisa_months <- .detect_dates(rtisa_s2[[1]])
    rtisa_vals <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_s2[[2]]))))
    pay_df <- data.frame(m = rtisa_months, v = rtisa_vals, stringsAsFactors = FALSE)
    pay_df <- pay_df[!is.na(pay_df$m) & !is.na(pay_df$v), ]
    pay_df <- pay_df[order(pay_df$m), ]

    months_cur  <- c(cm %m-% months(4), cm %m-% months(3), cm %m-% months(2))
    months_prev <- c(cm %m-% months(7), cm %m-% months(6), cm %m-% months(5))
    months_yago <- months_cur %m-% months(12)

    pay_cur_raw <- .get_avg(pay_df, months_cur)
    pay_prev3   <- .get_avg(pay_df, months_prev)
    pay_yago3   <- .get_avg(pay_df, months_yago)

    payroll_cur <- if (!is.na(pay_cur_raw)) pay_cur_raw / 1000 else NA_real_
    payroll_dq  <- if (!is.na(pay_cur_raw) && !is.na(pay_prev3)) (pay_cur_raw - pay_prev3) / 1000 else NA_real_
    payroll_dy  <- if (!is.na(pay_cur_raw) && !is.na(pay_yago3)) (pay_cur_raw - pay_yago3) / 1000 else NA_real_

    covid_base <- .get_avg(pay_df, as.Date(c("2019-12-01", "2020-01-01", "2020-02-01")))
    payroll_dc <- if (!is.na(pay_cur_raw) && !is.na(covid_base)) (pay_cur_raw - covid_base) / 1000 else NA_real_

    elec_base <- .get_avg(pay_df, as.Date(c("2024-04-01", "2024-05-01", "2024-06-01")))
    payroll_de <- if (!is.na(payroll_cur) && !is.na(elec_base)) payroll_cur - (elec_base / 1000) else NA_real_

    # flash (single month)
    flash_anchor <- anchor_m
    flash_val <- pay_df$v[match(anchor_m, pay_df$m)]
    flash_prev_m <- pay_df$v[match(anchor_m %m-% months(1), pay_df$m)]
    flash_prev_y <- pay_df$v[match(anchor_m %m-% months(12), pay_df$m)]
    flash_elec <- pay_df$v[match(as.Date("2024-06-01"), pay_df$m)]

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
  # WAGES NOMINAL (A01 Sheet 13 + 15)
  # ==========================================================================

  if (nrow(tbl_13) > 0 && ncol(tbl_13) >= 4) {
    nom_months <- .detect_dates(tbl_13[[1]])
    nom_v2 <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[2]]))))
    nom_v4 <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[4]]))))
    nom_df <- data.frame(m = nom_months, v2 = nom_v2, v4 = nom_v4, stringsAsFactors = FALSE)
    nom_df <- nom_df[!is.na(nom_df$m), ]
    nom_df <- nom_df[order(nom_df$m), ]

    latest_wages <- .pick_vals(nom_df$m, nom_df$v4, anchor_m)[1]

    win3 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
    prev3 <- c(anchor_m %m-% months(3), anchor_m %m-% months(4), anchor_m %m-% months(5))
    yago3 <- win3 %m-% months(12)
    covid3 <- as.Date(c("2019-12-01", "2020-01-01", "2020-02-01"))
    election3 <- as.Date(c("2024-04-01", "2024-05-01", "2024-06-01"))

    .wage_change <- function(a_months, b_months) {
      a <- .pick_avg(nom_df$m, nom_df$v2, a_months)
      b <- .pick_avg(nom_df$m, nom_df$v2, b_months)
      if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
    }

    wages_change_q <- .wage_change(win3, prev3)
    wages_change_y <- .wage_change(win3, yago3)
    wages_change_covid <- .wage_change(win3, covid3)
    wages_change_election <- .wage_change(win3, election3)

    # quarterly percentage point change for narrative
    prev_q_val <- .pick_vals(nom_df$m, nom_df$v4, anchor_m %m-% months(3))[1]
    wages_total_qchange <- if (!is.na(latest_wages) && !is.na(prev_q_val)) latest_wages - prev_q_val else NA_real_
  } else {
    latest_wages <- wages_change_q <- wages_change_y <- NA_real_
    wages_change_covid <- wages_change_election <- wages_total_qchange <- NA_real_
  }

  # A01 Sheet 15 - regular pay
  if (nrow(tbl_15) > 0 && ncol(tbl_15) >= 4) {
    a01s15_dates <- .detect_dates(tbl_15[[1]])
    a01s15_v4 <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_15[[4]]))))
    latest_regular_cash <- .pick_vals(a01s15_dates, a01s15_v4, anchor_m)[1]

    prev_q_reg <- .pick_vals(a01s15_dates, a01s15_v4, anchor_m %m-% months(3))[1]
    wages_reg_qchange <- if (!is.na(latest_regular_cash) && !is.na(prev_q_reg)) latest_regular_cash - prev_q_reg else NA_real_
  } else {
    latest_regular_cash <- NA_real_
    wages_reg_qchange <- NA_real_
  }

  # public/private sectors not available from A01 sheets directly - set NA
  wages_total_public  <- NA_real_
  wages_total_private <- NA_real_
  wages_reg_public    <- NA_real_
  wages_reg_private   <- NA_real_

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
  # WAGES CPI (X09 "AWE Real_CPI")
  # ==========================================================================

  if (nrow(tbl_cpi) > 0 && ncol(tbl_cpi) >= 9) {
    cpi_months <- .parse_cpi_dates(tbl_cpi[[1]])
    cpi_v2 <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[2]]))))
    cpi_v4 <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[4]]))))
    cpi_v9 <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[9]]))))
    cpi_df <- data.frame(m = cpi_months, v2 = cpi_v2, v4 = cpi_v4, v9 = cpi_v9, stringsAsFactors = FALSE)
    cpi_df <- cpi_df[!is.na(cpi_df$m), ]
    cpi_df <- cpi_df[order(cpi_df$m), ]

    latest_wages_cpi  <- .pick_vals(cpi_df$m, cpi_df$v4, anchor_m)[1]
    latest_regular_cpi <- .pick_vals(cpi_df$m, cpi_df$v9, anchor_m)[1]

    win3 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
    prev3 <- c(anchor_m %m-% months(3), anchor_m %m-% months(4), anchor_m %m-% months(5))
    yago3 <- win3 %m-% months(12)
    covid3 <- as.Date(c("2019-12-01", "2020-01-01", "2020-02-01"))
    election3 <- as.Date(c("2024-04-01", "2024-05-01", "2024-06-01"))

    .cpi_change <- function(a_months, b_months) {
      a <- .pick_avg(cpi_df$m, cpi_df$v2, a_months)
      b <- .pick_avg(cpi_df$m, cpi_df$v2, b_months)
      if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
    }

    wages_cpi_change_q <- .cpi_change(win3, prev3)
    wages_cpi_change_y <- .cpi_change(win3, yago3)
    wages_cpi_change_covid <- .cpi_change(win3, covid3)
    wages_cpi_change_election <- .cpi_change(win3, election3)

    # vs Dec 2007 and vs pandemic (2019 avg)
    dec2007_val <- .pick_vals(cpi_df$m, cpi_df$v2, as.Date("2007-12-01"))[1]
    cur_cpi_real <- .pick_avg(cpi_df$m, cpi_df$v2, win3)
    wages_cpi_total_vs_dec2007 <- if (!is.na(cur_cpi_real) && !is.na(dec2007_val) && dec2007_val != 0) {
      ((cur_cpi_real - dec2007_val) / dec2007_val) * 100
    } else NA_real_

    pandemic_months <- seq(as.Date("2019-01-01"), as.Date("2019-12-01"), by = "month")
    pandemic_avg <- .pick_avg(cpi_df$m, cpi_df$v2, pandemic_months)
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
  # DAYS LOST (A01 Sheet 18)
  # ==========================================================================

  if (nrow(tbl_18) > 0 && ncol(tbl_18) >= 2) {
    days_lost_cur <- .safe_last(.get_series(tbl_18, 2))
    # try to get the month label from col 1
    dl_dates <- .detect_dates(tbl_18[[1]])
    dl_vals <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_18[[2]]))))
    valid_idx <- which(!is.na(dl_dates) & !is.na(dl_vals))
    days_lost_label <- if (length(valid_idx) > 0) format(dl_dates[tail(valid_idx, 1)], "%B %Y") else ""
  } else {
    days_lost_cur <- NA_real_
    days_lost_label <- ""
  }

  assign("days_lost_cur",   days_lost_cur,   envir = target_env)
  assign("days_lost_label", days_lost_label, envir = target_env)

  # ==========================================================================
  # REDUNDANCY (A01 Sheet 10)
  # ==========================================================================

  if (nrow(tbl_10) > 0 && ncol(tbl_10) >= 3) {
    lab_cur <- .make_lfs_table_label(anchor_m)
    lab_q   <- .make_lfs_table_label(anchor_m %m-% months(3))
    lab_y   <- .make_lfs_table_label(anchor_m %m-% months(12))

    redund_cur <- .val_by_label(tbl_10, lab_cur, 3)
    redund_dq  <- redund_cur - .val_by_label(tbl_10, lab_q, 3)
    redund_dy  <- redund_cur - .val_by_label(tbl_10, lab_y, 3)
    redund_dc  <- redund_cur - .val_by_label(tbl_10, COVID_LFS_LABEL, 3)
    redund_de  <- redund_cur - .val_by_label(tbl_10, ELECTION_LABEL, 3)
  } else {
    redund_cur <- redund_dq <- redund_dy <- redund_dc <- redund_de <- NA_real_
  }

  assign("redund_cur", redund_cur, envir = target_env)
  assign("redund_dq",  redund_dq,  envir = target_env)
  assign("redund_dy",  redund_dy,  envir = target_env)
  assign("redund_dc",  redund_dc,  envir = target_env)
  assign("redund_de",  redund_de,  envir = target_env)

  # ==========================================================================
  # SECTOR PAYROLL (RTISA Sheet 24)
  # ==========================================================================

  if (nrow(rtisa_s24) > 0 && ncol(rtisa_s24) >= 18) {
    rtisa24_dates <- .detect_dates(rtisa_s24[[1]])

    colH <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_s24[[8]]))))
    colI <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_s24[[9]]))))
    colR <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_s24[[18]]))))

    valH_now  <- .pick_vals(rtisa24_dates, colH, anchor_m)[1]
    valI_now  <- .pick_vals(rtisa24_dates, colI, anchor_m)[1]
    valR_now  <- .pick_vals(rtisa24_dates, colR, anchor_m)[1]
    valH_yago <- .pick_vals(rtisa24_dates, colH, anchor_m %m-% months(12))[1]
    valI_yago <- .pick_vals(rtisa24_dates, colI, anchor_m %m-% months(12))[1]
    valR_yago <- .pick_vals(rtisa24_dates, colR, anchor_m %m-% months(12))[1]

    hosp_dy   <- if (!is.na(valH_now) && !is.na(valH_yago)) (valH_now - valH_yago) / 1000 else NA_real_
    retail_dy <- if (!is.na(valI_now) && !is.na(valI_yago)) (valI_now - valI_yago) / 1000 else NA_real_
    health_dy <- if (!is.na(valR_now) && !is.na(valR_yago)) (valR_now - valR_yago) / 1000 else NA_real_
  } else {
    hosp_dy <- retail_dy <- health_dy <- NA_real_
  }

  assign("hosp_dy",   hosp_dy,   envir = target_env)
  assign("hosp_dm",   NA_real_,  envir = target_env)
  assign("hosp_cur",  NA_real_,  envir = target_env)
  assign("hosp_dc",   NA_real_,  envir = target_env)
  assign("hosp_de",   NA_real_,  envir = target_env)
  assign("retail_dy", retail_dy, envir = target_env)
  assign("retail_dm", NA_real_,  envir = target_env)
  assign("retail_cur", NA_real_, envir = target_env)
  assign("retail_dc", NA_real_,  envir = target_env)
  assign("retail_de", NA_real_,  envir = target_env)
  assign("health_dy", health_dy, envir = target_env)
  assign("health_dm", NA_real_,  envir = target_env)
  assign("health_cur", NA_real_, envir = target_env)
  assign("health_dc", NA_real_,  envir = target_env)
  assign("health_de", NA_real_,  envir = target_env)

  # ==========================================================================
  # HR1 (Sheet "1a")
  # ==========================================================================

  if (nrow(hr1_tbl) > 0 && ncol(hr1_tbl) >= 13) {
    hr1_vals <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(hr1_tbl[[13]]))))
    hr1_cur <- .safe_last(hr1_vals[!is.na(hr1_vals)])
  } else {
    hr1_cur <- NA_real_
  }

  assign("hr1_cur", hr1_cur, envir = target_env)
  assign("hr1_dm",  NA_real_, envir = target_env)
  assign("hr1_dy",  NA_real_, envir = target_env)
  assign("hr1_dc",  NA_real_, envir = target_env)
  assign("hr1_de",  NA_real_, envir = target_env)

  # ==========================================================================
  # LABELS
  # ==========================================================================

  lfs_period_label <- lfs_label_narrative(m_emprt$end)
  lfs_period_short_label <- make_lfs_label(m_emprt$end)
  vacancies_period_short_label <- make_lfs_label(end_vac)
  payroll_flash_label <- format(flash_anchor, "%B %Y")
  hr1_month_label <- ""  # not reliably available from Excel without date parsing
  sector_month_label <- format(anchor_m, "%B %Y")

  assign("lfs_period_label",             lfs_period_label,             envir = target_env)
  assign("lfs_period_short_label",       lfs_period_short_label,       envir = target_env)
  assign("vacancies_period_short_label", vacancies_period_short_label, envir = target_env)
  assign("payroll_flash_label",          payroll_flash_label,          envir = target_env)
  assign("payroll_month_label",          format(anchor_m, "%B %Y"),    envir = target_env)
  assign("hr1_month_label",             hr1_month_label,              envir = target_env)
  assign("sector_month_label",          sector_month_label,           envir = target_env)
  assign("manual_month",                manual_month,                 envir = target_env)

  # inactivity driver text (not available from Excel)
  assign("inact_driver_text", "", envir = target_env)

  invisible(TRUE)
}
