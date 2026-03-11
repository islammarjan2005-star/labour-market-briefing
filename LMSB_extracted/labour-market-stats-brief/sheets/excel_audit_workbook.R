# excel_audit_workbook.R
# Creates the "LM Stats" audit workbook from uploaded ONS Excel files.
# Mirrors the structure of the manually-maintained "LM Stats.xlsx" workbook.
#
# Each data sheet is populated from the relevant ONS source file with
# computed comparison columns (change on quarter, year, COVID, election).

suppressPackageStartupMessages({
  library(openxlsx)
  library(readxl)
  library(lubridate)
})

# load shared helpers (parse_manual_month, lfs_label_narrative, etc.)
if (!exists("parse_manual_month", inherits = TRUE)) {
  source("utils/helpers.R")
}

# ============================================================================
# INTERNAL HELPERS
# ============================================================================

.safe_read <- function(path, sheet, ...) {
  if (is.null(path)) return(data.frame())
  tryCatch(
    suppressMessages(read_excel(path, sheet = sheet, col_names = FALSE, ...)),
    error = function(e) {
      message("[audit wb] Could not read '", sheet, "' from ", basename(path), ": ", e$message)
      data.frame()
    }
  )
}

.safe_read_named <- function(path, sheet, ...) {
  if (is.null(path)) return(data.frame())
  tryCatch(
    suppressMessages(read_excel(path, sheet = sheet, ...)),
    error = function(e) data.frame()
  )
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

# Build LFS 3-month label: "Oct-Dec 2025" from end date 2025-12-01
.lfs_label <- function(end_date) {
  start_date <- end_date %m-% months(2)
  sprintf("%s-%s %s", format(start_date, "%b"), format(end_date, "%b"), format(end_date, "%Y"))
}

# Find row by column-1 label (trimmed, case-insensitive)
.find_row <- function(tbl, label) {
  if (nrow(tbl) == 0 || ncol(tbl) == 0) return(NA_integer_)
  col1 <- trimws(as.character(tbl[[1]]))
  idx <- which(tolower(col1) == tolower(trimws(label)))
  if (length(idx) == 0) NA_integer_ else idx[1]
}

# Extract numeric value at [row, col]
.cell_num <- function(tbl, row, col) {
  if (is.na(row) || row < 1 || row > nrow(tbl) || col > ncol(tbl)) return(NA_real_)
  suppressWarnings(as.numeric(gsub("[^0-9.eE+-]", "", as.character(tbl[[col]][row]))))
}

# Lookup value by date
.val_by_date <- function(df_m, df_v, target_date) {
  idx <- which(df_m == target_date)
  if (length(idx) == 0) return(NA_real_)
  df_v[idx[1]]
}

# Average over multiple dates
.avg_by_dates <- function(df_m, df_v, target_dates) {
  vals <- vapply(target_dates, function(d) .val_by_date(df_m, df_v, d), numeric(1))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

# Compute LFS metric with comparison periods
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

# ---- Formatting helpers ----

.style_header <- function() {
  createStyle(fontSize = 11, fontColour = "#FFFFFF", fgFill = "#1F4E79",
              halign = "center", textDecoration = "bold",
              border = "TopBottomLeftRight", borderColour = "#1F4E79")
}

.style_title <- function() {
  createStyle(fontSize = 14, textDecoration = "bold", fontColour = "#1F4E79")
}

.style_subtitle <- function() {
  createStyle(fontSize = 11, textDecoration = "bold", fontColour = "#505050")
}

.style_number <- function() {
  createStyle(numFmt = "#,##0")
}

.style_pct <- function() {
  createStyle(numFmt = "0.0")
}

.style_positive <- function() {
  createStyle(fontColour = "#006100", fgFill = "#C6EFCE")
}

.style_negative <- function() {
  createStyle(fontColour = "#9C0006", fgFill = "#FFC7CE")
}

.style_separator <- function() {
  createStyle(fontSize = 16, textDecoration = "bold", fontColour = "#1F4E79",
              fgFill = "#D9E2F3", halign = "center", valign = "center")
}

# Write a sheet that is just a separator header (e.g. "Charts >>>")
.add_separator_sheet <- function(wb, name) {
  addWorksheet(wb, name, tabColour = "#B4C6E7")
  writeData(wb, name, data.frame(X = name), startRow = 1, colNames = FALSE)
  addStyle(wb, name, .style_separator(), rows = 1, cols = 1)
  setColWidths(wb, name, cols = 1, widths = 50)
  setRowHeights(wb, name, rows = 1, heights = 60)
}

# Write a source data sheet with title and auto-column-widths
.write_data_sheet <- function(wb, name, data, title = NULL, tab_colour = NULL) {
  addWorksheet(wb, name, tabColour = tab_colour)
  start_row <- 1
  if (!is.null(title)) {
    writeData(wb, name, data.frame(X = title), startRow = 1, colNames = FALSE)
    addStyle(wb, name, .style_title(), rows = 1, cols = 1)
    start_row <- 3
  }
  if (nrow(data) > 0) {
    writeData(wb, name, data, startRow = start_row, colNames = FALSE)
  }
}

# Copy a sheet from source file to workbook with optional rename
.copy_sheet <- function(wb, dest_name, src_path, src_sheet, title = NULL, tab_colour = NULL) {
  data <- .safe_read(src_path, src_sheet)
  if (nrow(data) == 0) {
    addWorksheet(wb, dest_name, tabColour = tab_colour)
    writeData(wb, dest_name, data.frame(Note = paste0("Source data not available (", src_sheet, ")")))
    return(invisible(NULL))
  }
  .write_data_sheet(wb, dest_name, data, title = title, tab_colour = tab_colour)
}


# ============================================================================
# MAIN FUNCTION
# ============================================================================

create_audit_workbook <- function(
    output_path,
    file_a01 = NULL,
    file_hr1 = NULL,
    file_x09 = NULL,
    file_rtisa = NULL,
    file_cla01 = NULL,
    file_x02 = NULL,
    file_oecd_unemp = NULL,
    file_oecd_emp = NULL,
    file_oecd_inact = NULL,
    calculations_path = NULL,
    config_path = NULL,
    vacancies_mode = "aligned",
    payroll_mode = "aligned",
    manual_month_override = NULL,
    verbose = FALSE
) {

  wb <- createWorkbook()

  # detect reference period from A01
  anchor_m <- NULL
  if (!is.null(file_a01)) {
    tbl1 <- .safe_read(file_a01, "1")
    if (nrow(tbl1) > 0) {
      col1 <- trimws(as.character(tbl1[[1]]))
      lfs_pat <- "^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\s+(\\d{4})$"
      hits <- grep(lfs_pat, col1, ignore.case = TRUE)
      if (length(hits) > 0) {
        last_label <- col1[hits[length(hits)]]
        parts <- regmatches(last_label, regexec(lfs_pat, last_label, ignore.case = TRUE))[[1]]
        end_mon <- match(tools::toTitleCase(tolower(parts[3])), month.abb)
        end_yr  <- as.integer(parts[4])
        if (!is.na(end_mon) && !is.na(end_yr)) {
          anchor_m <- as.Date(sprintf("%04d-%02d-01", end_yr, end_mon))
        }
      }
    }
  }
  if (is.null(anchor_m)) anchor_m <- Sys.Date() %m-% months(2)

  cm <- anchor_m %m+% months(2)  # reference month (e.g. Feb 2026)
  lfs_end_cur   <- anchor_m
  lfs_end_q     <- anchor_m %m-% months(3)
  lfs_end_y     <- anchor_m %m-% months(12)
  lfs_end_covid <- as.Date("2020-02-01")
  lfs_end_elec  <- as.Date("2024-06-01")

  lab_cur   <- .lfs_label(lfs_end_cur)
  lab_q     <- .lfs_label(lfs_end_q)
  lab_y     <- .lfs_label(lfs_end_y)
  lab_covid <- .lfs_label(lfs_end_covid)
  lab_elec  <- .lfs_label(lfs_end_elec)

  all_labels <- c(lab_cur, lab_q, lab_y, lab_covid, lab_elec)

  ref_label <- format(cm, "%B %Y")  # "February 2026"

  if (verbose) message("[audit wb] Reference: ", ref_label, " | LFS end: ", anchor_m)

  # ========================================================================
  # 1. HOW TO UPDATE
  # ========================================================================

  addWorksheet(wb, "How to update", tabColour = "#FFC000")
  info_text <- data.frame(V1 = c(
    paste0("Labour Market Statistics Briefing — ", ref_label),
    "",
    "HOW TO UPDATE THIS WORKBOOK",
    "----------------------------",
    "This workbook is auto-generated from ONS source datasets.",
    "To update, download the latest files from ONS and upload them via the app.",
    "",
    "Required files:",
    "  A01  — Summary of labour market statistics",
    "  HR1  — Advanced notification of potential redundancies",
    "  X09  — Real average weekly earnings using CPI (SA)",
    "  RTISA — Payrolled employees, seasonally adjusted (PAYE RTI)",
    "",
    "Supplementary files (for additional sheets):",
    "  CLA01 — Claimant Count (Jobseeker's Allowance / Universal Credit)",
    "  X02   — Labour Force Survey flows",
    "  OECD  — International comparison data (3 files: unemployment, employment, inactivity)",
    "",
    "All comparison periods:",
    paste0("  Current:    ", lab_cur),
    paste0("  vs Quarter: ", lab_q),
    paste0("  vs Year:    ", lab_y),
    paste0("  vs COVID:   ", lab_covid),
    paste0("  vs Election:", lab_elec)
  ))
  writeData(wb, "How to update", info_text, colNames = FALSE)
  addStyle(wb, "How to update", .style_title(), rows = 1, cols = 1)
  setColWidths(wb, "How to update", cols = 1, widths = 80)

  # ========================================================================
  # 2. DATA LINKS
  # ========================================================================

  addWorksheet(wb, "Data links", tabColour = "#FFC000")
  links_df <- data.frame(
    Sheet = c("1. Payrolled employees (UK)", "23. Employees Industry",
              "2", "3", "5", "10", "11", "13", "15", "18", "20", "21", "22",
              "1 UK", "AWE Real_CPI", "1a / 1b / 2a / 2b",
              "LFS Labour market flows SA", "RTI. Employee flows (UK)",
              "Final Table / Unemployment / Employment / Inactivity",
              "Regional breakdowns"),
    Source_File = c("RTISA", "RTISA",
                    "A01", "A01", "A01", "A01", "A01", "A01", "A01", "A01", "A01", "A01", "A01",
                    "CLA01", "X09", "HR1",
                    "X02", "RTISA",
                    "OECD Data Explorer",
                    "A01"),
    Source_Sheet = c("1. Payrolled employees (UK)", "23. Employees (Industry)",
                     "2", "3", "5", "10", "11", "13", "15", "18", "20", "21", "22",
                     "(claimant count)", "AWE Real_CPI", "1a, 1b, 2a, 2b",
                     "LFS Labour market flows SA", "6. Employee flows (UK)",
                     "Downloaded from OECD",
                     "22"),
    Description = c(
      "Monthly payrolled employee counts (HMRC PAYE), seasonally adjusted",
      "Payrolled employees by industry (SIC 2007 sections), SA",
      "Labour market activity by age group, seasonally adjusted",
      "Full-time, part-time and temporary workers",
      "Workforce jobs (seasonally adjusted)",
      "Redundancies levels and rates (LFS)",
      "Economic inactivity by reason and status",
      "Average Weekly Earnings — total pay (nominal)",
      "Average Weekly Earnings — regular pay (nominal)",
      "Labour disputes (working days lost, stoppages)",
      "Vacancies and unemployment comparison",
      "Vacancies by industry",
      "Regional Labour Force Survey",
      "Claimant Count (JSA / Universal Credit)",
      "Real Average Weekly Earnings using CPI (2015 prices)",
      "HR1 redundancy notifications by region",
      "Labour market flows between employment, unemployment, inactivity",
      "Flows of payrolled employees from PAYE RTI",
      "International comparison — OECD unemployment, employment, inactivity rates",
      "Employment, unemployment and inactivity rates by UK region"
    ),
    stringsAsFactors = FALSE
  )
  writeData(wb, "Data links", links_df, headerStyle = .style_header())
  setColWidths(wb, "Data links", cols = 1:4, widths = c(35, 15, 30, 60))

  # ========================================================================
  # 3. DASHBOARD
  # ========================================================================

  tbl_1 <- .safe_read(file_a01, "1")
  tbl_2 <- .safe_read(file_a01, "2")
  tbl_19 <- .safe_read(file_a01, "19")

  addWorksheet(wb, "Dashboard", tabColour = "#00703C")

  # Compute dashboard metrics from A01
  if (nrow(tbl_1) > 0) {
    m_emp16   <- .lfs_metric(tbl_1, 4,  all_labels)
    m_emprt   <- .lfs_metric(tbl_1, 17, all_labels)
    m_unemp16 <- .lfs_metric(tbl_1, 5,  all_labels)
    m_unemprt <- .lfs_metric(tbl_1, 9,  all_labels)
    m_inact   <- .lfs_metric(tbl_1, 15, all_labels)
    m_inactrt <- .lfs_metric(tbl_1, 19, all_labels)
  } else {
    na_m <- list(cur = NA, dq = NA, dy = NA, dc = NA, de = NA)
    m_emp16 <- m_emprt <- m_unemp16 <- m_unemprt <- m_inact <- m_inactrt <- na_m
  }

  if (nrow(tbl_2) > 0 && ncol(tbl_2) >= 57) {
    m_5064   <- .lfs_metric(tbl_2, 56, all_labels)
    m_5064rt <- .lfs_metric(tbl_2, 57, all_labels)
  } else {
    na_m <- list(cur = NA, dq = NA, dy = NA, dc = NA, de = NA)
    m_5064 <- m_5064rt <- na_m
  }

  # Vacancies
  vac_m <- list(cur = NA, dq = NA, dy = NA, dc = NA, de = NA)
  if (nrow(tbl_19) > 0 && ncol(tbl_19) >= 3) {
    vac_lab_cur   <- lab_cur
    vac_lab_q     <- .lfs_label(lfs_end_cur %m-% months(3))
    vac_lab_y     <- .lfs_label(lfs_end_cur %m-% months(12))
    vac_lab_covid <- "Jan-Mar 2020"
    vac_lab_elec  <- .lfs_label(as.Date("2024-06-01"))

    r_cur <- .find_row(tbl_19, vac_lab_cur)
    vac_m$cur <- .cell_num(tbl_19, r_cur, 3)
    vac_m$dq  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_q), 3)
    vac_m$dy  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_y), 3)
    vac_m$dc  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_covid), 3)
    vac_m$de  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, vac_lab_elec), 3)
  }

  # Payroll
  pay_m <- list(cur = NA, dq = NA, dy = NA, dc = NA, de = NA, flash_cur = NA, flash_dm = NA, flash_dy = NA)
  rtisa_pay <- .safe_read(file_rtisa, "1. Payrolled employees (UK)")
  if (nrow(rtisa_pay) > 0 && ncol(rtisa_pay) >= 2) {
    rtisa_text <- trimws(as.character(rtisa_pay[[1]]))
    rtisa_parsed <- suppressWarnings(lubridate::parse_date_time(rtisa_text, orders = c("B Y", "bY", "BY")))
    rtisa_months <- floor_date(as.Date(rtisa_parsed), "month")
    rtisa_vals <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_pay[[2]]))))

    pay_df <- data.frame(m = rtisa_months, v = rtisa_vals, stringsAsFactors = FALSE)
    pay_df <- pay_df[!is.na(pay_df$m) & !is.na(pay_df$v), ]
    pay_df <- pay_df[order(pay_df$m), ]

    if (nrow(pay_df) > 0) {
      rtisa_latest <- pay_df$m[nrow(pay_df)]
      months_cur  <- c(cm %m-% months(4), cm %m-% months(3), cm %m-% months(2))
      months_prev <- c(cm %m-% months(7), cm %m-% months(6), cm %m-% months(5))
      months_yago <- months_cur %m-% months(12)

      pc <- .avg_by_dates(pay_df$m, pay_df$v, months_cur)
      pp <- .avg_by_dates(pay_df$m, pay_df$v, months_prev)
      py <- .avg_by_dates(pay_df$m, pay_df$v, months_yago)
      pcov <- .avg_by_dates(pay_df$m, pay_df$v, as.Date(c("2019-12-01", "2020-01-01", "2020-02-01")))
      pelec <- .avg_by_dates(pay_df$m, pay_df$v, as.Date(c("2024-04-01", "2024-05-01", "2024-06-01")))

      pay_m$cur <- if (!is.na(pc)) pc / 1000 else NA
      pay_m$dq  <- if (!is.na(pc) && !is.na(pp)) (pc - pp) / 1000 else NA
      pay_m$dy  <- if (!is.na(pc) && !is.na(py)) (pc - py) / 1000 else NA
      pay_m$dc  <- if (!is.na(pc) && !is.na(pcov)) (pc - pcov) / 1000 else NA
      pay_m$de  <- if (!is.na(pc) && !is.na(pelec)) (pc - pelec) / 1000 else NA

      # Flash (latest single month)
      fv <- .val_by_date(pay_df$m, pay_df$v, rtisa_latest)
      fpm <- .val_by_date(pay_df$m, pay_df$v, rtisa_latest %m-% months(1))
      fpy <- .val_by_date(pay_df$m, pay_df$v, rtisa_latest %m-% months(12))
      pay_m$flash_cur <- if (!is.na(fv)) fv / 1e6 else NA
      pay_m$flash_dm  <- if (!is.na(fv) && !is.na(fpm)) (fv - fpm) / 1000 else NA
      pay_m$flash_dy  <- if (!is.na(fv) && !is.na(fpy)) (fv - fpy) / 1000 else NA
    }
  }

  # Wages (nominal total pay from A01 Sheet 13)
  wages_m <- list(cur = NA, dq = NA, dy = NA, dc = NA, de = NA)
  tbl_13 <- .safe_read(file_a01, "13")
  if (nrow(tbl_13) > 0 && ncol(tbl_13) >= 4) {
    w13_dates <- .detect_dates(tbl_13[[1]])
    w13_pct   <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[4]]))))
    w13_weekly <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[2]]))))
    wages_m$cur <- .val_by_date(w13_dates, w13_pct, anchor_m)

    win3 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
    prev3 <- c(anchor_m %m-% months(3), anchor_m %m-% months(4), anchor_m %m-% months(5))
    yago3 <- win3 %m-% months(12)
    covid3 <- as.Date(c("2019-12-01", "2020-01-01", "2020-02-01"))
    election3 <- as.Date(c("2024-04-01", "2024-05-01", "2024-06-01"))

    .wc <- function(a, b) {
      va <- .avg_by_dates(w13_dates, w13_weekly, a)
      vb <- .avg_by_dates(w13_dates, w13_weekly, b)
      if (is.na(va) || is.na(vb)) NA else (va - vb) * 52
    }
    wages_m$dq <- .wc(win3, prev3)
    wages_m$dy <- .wc(win3, yago3)
    wages_m$dc <- .wc(win3, covid3)
    wages_m$de <- .wc(win3, election3)
  }

  # Wages (CPI-adjusted from X09)
  wages_cpi_m <- list(cur = NA, dq = NA, dy = NA, dc = NA, de = NA)
  tbl_cpi <- .safe_read(file_x09, "AWE Real_CPI")
  if (nrow(tbl_cpi) > 0 && ncol(tbl_cpi) >= 9) {
    cpi_months <- .detect_dates(tbl_cpi[[1]])
    cpi_real   <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[2]]))))
    cpi_total  <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[5]]))))

    cpi_valid <- which(!is.na(cpi_months) & !is.na(cpi_total))
    cpi_anchor <- if (length(cpi_valid) > 0) cpi_months[cpi_valid[length(cpi_valid)]] else anchor_m
    wages_cpi_m$cur <- .val_by_date(cpi_months, cpi_total, cpi_anchor)

    cpi_win3 <- c(cpi_anchor, cpi_anchor %m-% months(1), cpi_anchor %m-% months(2))
    cpi_prev3 <- c(cpi_anchor %m-% months(3), cpi_anchor %m-% months(4), cpi_anchor %m-% months(5))
    cpi_yago3 <- cpi_win3 %m-% months(12)
    cpi_covid3 <- as.Date(c("2019-12-01", "2020-01-01", "2020-02-01"))
    cpi_elec3 <- as.Date(c("2024-04-01", "2024-05-01", "2024-06-01"))

    .cc <- function(a, b) {
      va <- .avg_by_dates(cpi_months, cpi_real, a)
      vb <- .avg_by_dates(cpi_months, cpi_real, b)
      if (is.na(va) || is.na(vb)) NA else (va - vb) * 52
    }
    wages_cpi_m$dq <- .cc(cpi_win3, cpi_prev3)
    wages_cpi_m$dy <- .cc(cpi_win3, cpi_yago3)
    wages_cpi_m$dc <- .cc(cpi_win3, cpi_covid3)
    wages_cpi_m$de <- .cc(cpi_win3, cpi_elec3)
  }

  # Build dashboard data frame
  dash_df <- data.frame(
    Metric = c(
      "Employment 16+ (000s)", "Employment rate 16-64 (%)",
      "Unemployment 16+ (000s)", "Unemployment rate 16+ (%)",
      "Economic inactivity 16-64 (000s)", "Economic inactivity rate 16-64 (%)",
      "Inactivity 50-64 (000s)", "Inactivity rate 50-64 (%)",
      "HMRC Payrolled employees (000s)",
      "Vacancies (000s)",
      "Wages — total pay growth (%)", "Wages — real (CPI) growth (%)"
    ),
    Current = c(
      m_emp16$cur / 1000, m_emprt$cur,
      m_unemp16$cur / 1000, m_unemprt$cur,
      m_inact$cur / 1000, m_inactrt$cur,
      m_5064$cur / 1000, m_5064rt$cur,
      pay_m$cur, vac_m$cur,
      wages_m$cur, wages_cpi_m$cur
    ),
    Change_on_quarter = c(
      m_emp16$dq / 1000, m_emprt$dq,
      m_unemp16$dq / 1000, m_unemprt$dq,
      m_inact$dq / 1000, m_inactrt$dq,
      m_5064$dq / 1000, m_5064rt$dq,
      pay_m$dq, vac_m$dq,
      wages_m$dq, wages_cpi_m$dq
    ),
    Change_on_year = c(
      m_emp16$dy / 1000, m_emprt$dy,
      m_unemp16$dy / 1000, m_unemprt$dy,
      m_inact$dy / 1000, m_inactrt$dy,
      m_5064$dy / 1000, m_5064rt$dy,
      pay_m$dy, vac_m$dy,
      wages_m$dy, wages_cpi_m$dy
    ),
    Change_since_COVID = c(
      m_emp16$dc / 1000, m_emprt$dc,
      m_unemp16$dc / 1000, m_unemprt$dc,
      m_inact$dc / 1000, m_inactrt$dc,
      m_5064$dc / 1000, m_5064rt$dc,
      pay_m$dc, vac_m$dc,
      wages_m$dc, wages_cpi_m$dc
    ),
    Change_since_election = c(
      m_emp16$de / 1000, m_emprt$de,
      m_unemp16$de / 1000, m_unemprt$de,
      m_inact$de / 1000, m_inactrt$de,
      m_5064$de / 1000, m_5064rt$de,
      pay_m$de, vac_m$de,
      wages_m$de, wages_cpi_m$de
    ),
    stringsAsFactors = FALSE
  )

  # Title rows
  writeData(wb, "Dashboard", data.frame(V1 = paste0("Labour Market Dashboard — ", ref_label)),
            startRow = 1, colNames = FALSE)
  addStyle(wb, "Dashboard", .style_title(), rows = 1, cols = 1)
  writeData(wb, "Dashboard", data.frame(V1 = paste0("LFS period: ", lab_cur)),
            startRow = 2, colNames = FALSE)
  addStyle(wb, "Dashboard", .style_subtitle(), rows = 2, cols = 1)

  # Column headers
  header_names <- c("Metric", "Current", "Change on quarter", "Change on year",
                     "Change since COVID-19", "Change since election")
  writeData(wb, "Dashboard", as.data.frame(t(header_names)), startRow = 4, colNames = FALSE)
  addStyle(wb, "Dashboard", .style_header(), rows = 4, cols = 1:6, gridExpand = TRUE)

  # Data
  writeData(wb, "Dashboard", dash_df, startRow = 5, colNames = FALSE)

  # Conditional formatting: green/red for changes
  for (col_idx in 3:6) {
    for (row_idx in 5:16) {
      conditionalFormatting(wb, "Dashboard", cols = col_idx, rows = row_idx,
                            type = "expression", rule = ">0",
                            style = .style_positive())
      conditionalFormatting(wb, "Dashboard", cols = col_idx, rows = row_idx,
                            type = "expression", rule = "<0",
                            style = .style_negative())
    }
  }

  setColWidths(wb, "Dashboard", cols = 1:6, widths = c(35, 15, 20, 18, 22, 22))

  # ========================================================================
  # 4. PAYROLLED EMPLOYEES (from RTISA)
  # ========================================================================

  .copy_sheet(wb, "1. Payrolled employees (UK)", file_rtisa,
              "1. Payrolled employees (UK)",
              title = "Payrolled employees, seasonally adjusted, UK",
              tab_colour = "#4472C4")

  # ========================================================================
  # 5. EMPLOYEES BY INDUSTRY (from RTISA)
  # ========================================================================

  .copy_sheet(wb, "23. Employees Industry", file_rtisa,
              "23. Employees (Industry)",
              title = "Payrolled employees by industry (SIC 2007), seasonally adjusted, UK",
              tab_colour = "#4472C4")

  # ========================================================================
  # 6-17. A01-SOURCED DATA SHEETS
  # ========================================================================

  # Sheet "2": Labour market by age group
  .copy_sheet(wb, "2", file_a01, "2",
              title = "Labour market activity by age group, seasonally adjusted, UK",
              tab_colour = "#548235")

  # Sheet "Sheet1": Unemployment time series (from A01 Sheet "1")
  .copy_sheet(wb, "Sheet1", file_a01, "1",
              title = "Summary of Labour Market Statistics, seasonally adjusted, UK",
              tab_colour = "#548235")

  # Sheet "3": Full-time, part-time and temporary workers
  .copy_sheet(wb, "3", file_a01, "3",
              title = "Full-time, part-time and temporary workers, seasonally adjusted, UK",
              tab_colour = "#548235")

  # Sheet "5": Workforce jobs
  .copy_sheet(wb, "5", file_a01, "5",
              title = "Workforce jobs, seasonally adjusted, UK",
              tab_colour = "#548235")

  # Sheet "10": Redundancies
  .copy_sheet(wb, "10", file_a01, "10",
              title = "Redundancies: levels and rates, seasonally adjusted, UK",
              tab_colour = "#548235")

  # Sheet "11": Economic inactivity by reason
  .copy_sheet(wb, "11", file_a01, "11",
              title = "Economic inactivity by reason, seasonally adjusted, UK, aged 16-64",
              tab_colour = "#548235")

  # Sheet "13": AWE Total Pay (nominal)
  .copy_sheet(wb, "13", file_a01, "13",
              title = "Average Weekly Earnings — total pay (nominal), seasonally adjusted",
              tab_colour = "#548235")

  # Sheet "15": AWE Regular Pay (nominal)
  .copy_sheet(wb, "15", file_a01, "15",
              title = "Average Weekly Earnings — regular pay (nominal), seasonally adjusted",
              tab_colour = "#548235")

  # Sheet "18": Labour disputes
  .copy_sheet(wb, "18", file_a01, "18",
              title = "Labour Disputes summary",
              tab_colour = "#548235")

  # Sheet "21": Vacancies by industry
  .copy_sheet(wb, "21", file_a01, "21",
              title = "Vacancies by industry, seasonally adjusted, UK",
              tab_colour = "#548235")

  # Sheet "20": Vacancies and Unemployment
  .copy_sheet(wb, "20", file_a01, "20",
              title = "Vacancies and Unemployment, seasonally adjusted, UK",
              tab_colour = "#548235")

  # Sheet "22": Regional LFS
  .copy_sheet(wb, "22", file_a01, "22",
              title = "Regional Labour Force Survey, seasonally adjusted",
              tab_colour = "#548235")

  # ========================================================================
  # 18. CLAIMANT COUNT (CLA01) — Sheet "1 UK"
  # ========================================================================

  if (!is.null(file_cla01)) {
    # CLA01 files typically have a sheet named "1" or "People SA"
    cla_sheets <- tryCatch(readxl::excel_sheets(file_cla01), error = function(e) character(0))
    cla_sheet <- NULL
    # Try common sheet names
    for (candidate in c("1", "People SA", "People", cla_sheets[1])) {
      if (candidate %in% cla_sheets) { cla_sheet <- candidate; break }
    }
    if (!is.null(cla_sheet)) {
      .copy_sheet(wb, "1 UK", file_cla01, cla_sheet,
                  title = "CLA01 — Claimant Count, seasonally adjusted, UK",
                  tab_colour = "#BF8F00")
    } else {
      addWorksheet(wb, "1 UK", tabColour = "#BF8F00")
      writeData(wb, "1 UK", data.frame(Note = "CLA01 file uploaded but no recognised sheet found."))
    }
  } else {
    addWorksheet(wb, "1 UK", tabColour = "#BF8F00")
    writeData(wb, "1 UK", data.frame(Note = "Upload CLA01 file to populate this sheet."))
  }

  # ========================================================================
  # 19. REAL WAGES (X09)
  # ========================================================================

  .copy_sheet(wb, "AWE Real_CPI", file_x09, "AWE Real_CPI",
              title = "Real Average Weekly Earnings using CPI, seasonally adjusted (2015=100 prices)",
              tab_colour = "#7030A0")

  # ========================================================================
  # 20-24. REDUNDANCY NOTIFICATIONS (HR1)
  # ========================================================================

  .add_separator_sheet(wb, "Redundancies >>>")

  for (hr1_sheet in c("1a", "1b", "2a", "2b")) {
    .copy_sheet(wb, hr1_sheet, file_hr1, hr1_sheet,
                title = paste0("HR1 — ", hr1_sheet, ": Redundancy notifications"),
                tab_colour = "#C00000")
  }

  # ========================================================================
  # 25-27. LABOUR MARKET FLOWS
  # ========================================================================

  .add_separator_sheet(wb, "Labour market flows >>>")

  if (!is.null(file_x02)) {
    x02_sheets <- tryCatch(readxl::excel_sheets(file_x02), error = function(e) character(0))
    # X02 typically has "People SA" or "LFS Labour market flows SA"
    x02_sheet <- NULL
    for (candidate in c("LFS Labour market flows SA", "People SA", "1", x02_sheets[1])) {
      if (candidate %in% x02_sheets) { x02_sheet <- candidate; break }
    }
    if (!is.null(x02_sheet)) {
      .copy_sheet(wb, "LFS Labour market flows SA", file_x02, x02_sheet,
                  title = "X02 — Labour Force Survey flows, seasonally adjusted, people aged 16-64",
                  tab_colour = "#ED7D31")
    } else {
      addWorksheet(wb, "LFS Labour market flows SA", tabColour = "#ED7D31")
      writeData(wb, "LFS Labour market flows SA", data.frame(Note = "X02 file uploaded but no recognised sheet."))
    }
  } else {
    addWorksheet(wb, "LFS Labour market flows SA", tabColour = "#ED7D31")
    writeData(wb, "LFS Labour market flows SA", data.frame(Note = "Upload X02 file to populate this sheet."))
  }

  # RTI Employee flows (from RTISA Sheet "6. Employee flows (UK)")
  .copy_sheet(wb, "RTI. Employee flows (UK)", file_rtisa,
              "6. Employee flows (UK)",
              title = "Flows of payrolled employees from PAYE RTI, seasonally adjusted, UK",
              tab_colour = "#4472C4")

  # ========================================================================
  # 28-32. INTERNATIONAL COMPARISONS (OECD)
  # ========================================================================

  .add_separator_sheet(wb, "International Comparisons >>>")

  # Final Table — combined international data
  # If OECD files are provided, build a combined comparison table
  has_oecd <- !is.null(file_oecd_unemp) || !is.null(file_oecd_emp) || !is.null(file_oecd_inact)

  if (has_oecd) {
    addWorksheet(wb, "Final Table", tabColour = "#2F5496")
    writeData(wb, "Final Table", data.frame(V1 = "OECD International Comparisons — Infra-annual labour statistics"),
              startRow = 1, colNames = FALSE)
    addStyle(wb, "Final Table", .style_title(), rows = 1, cols = 1)

    # Read each OECD file and write to individual sheets
    .read_oecd <- function(path) {
      if (is.null(path)) return(data.frame())
      # Try reading as CSV first (OECD Data Explorer exports CSVs)
      ext <- tolower(tools::file_ext(path))
      if (ext == "csv") {
        tryCatch(read.csv(path, stringsAsFactors = FALSE),
                 error = function(e) data.frame())
      } else {
        # Try as Excel
        tryCatch(
          suppressMessages(read_excel(path, col_names = TRUE)),
          error = function(e) data.frame()
        )
      }
    }

    oecd_unemp <- .read_oecd(file_oecd_unemp)
    oecd_emp   <- .read_oecd(file_oecd_emp)
    oecd_inact <- .read_oecd(file_oecd_inact)

    # Write combined overview
    ft_row <- 3
    if (nrow(oecd_unemp) > 0) {
      writeData(wb, "Final Table", data.frame(V1 = "Unemployment rates (15+, %)"),
                startRow = ft_row, colNames = FALSE)
      addStyle(wb, "Final Table", .style_subtitle(), rows = ft_row, cols = 1)
      writeData(wb, "Final Table", oecd_unemp, startRow = ft_row + 1, headerStyle = .style_header())
      ft_row <- ft_row + nrow(oecd_unemp) + 3
    }
    if (nrow(oecd_emp) > 0) {
      writeData(wb, "Final Table", data.frame(V1 = "Employment rates (15-64, %)"),
                startRow = ft_row, colNames = FALSE)
      addStyle(wb, "Final Table", .style_subtitle(), rows = ft_row, cols = 1)
      writeData(wb, "Final Table", oecd_emp, startRow = ft_row + 1, headerStyle = .style_header())
      ft_row <- ft_row + nrow(oecd_emp) + 3
    }
    if (nrow(oecd_inact) > 0) {
      writeData(wb, "Final Table", data.frame(V1 = "Inactivity rates (15-64, %)"),
                startRow = ft_row, colNames = FALSE)
      addStyle(wb, "Final Table", .style_subtitle(), rows = ft_row, cols = 1)
      writeData(wb, "Final Table", oecd_inact, startRow = ft_row + 1, headerStyle = .style_header())
    }

    # Individual OECD metric sheets
    .write_oecd_sheet <- function(wb, name, data, subtitle) {
      addWorksheet(wb, name, tabColour = "#2F5496")
      writeData(wb, name, data.frame(V1 = subtitle), startRow = 1, colNames = FALSE)
      addStyle(wb, name, .style_title(), rows = 1, cols = 1)
      if (nrow(data) > 0) {
        writeData(wb, name, data, startRow = 3, headerStyle = .style_header())
      } else {
        writeData(wb, name, data.frame(Note = "OECD data not provided."), startRow = 3)
      }
    }

    .write_oecd_sheet(wb, "Unemployment", oecd_unemp,
                      "Infra-annual labour statistics — Unemployment rate (15+, %)")
    .write_oecd_sheet(wb, "Employment", oecd_emp,
                      "Infra-annual labour statistics — Employment rate (15-64, %)")
    .write_oecd_sheet(wb, "Inactivity", oecd_inact,
                      "Infra-annual labour statistics — Inactivity rate (15-64, %)")
  } else {
    for (sn in c("Final Table", "Unemployment", "Employment", "Inactivity")) {
      addWorksheet(wb, sn, tabColour = "#2F5496")
      writeData(wb, sn, data.frame(Note = "Upload OECD data files to populate international comparisons."))
    }
  }

  # ========================================================================
  # 33-38. CHART DATA SHEETS
  # ========================================================================

  .add_separator_sheet(wb, "Charts >>>")

  # --- Wages charts ---
  addWorksheet(wb, "Wages charts", tabColour = "#A9D18E")
  if (nrow(tbl_13) > 0 && ncol(tbl_13) >= 4) {
    w_dates <- .detect_dates(tbl_13[[1]])
    w_total <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[4]]))))

    tbl_15_w <- .safe_read(file_a01, "15")
    w_reg <- if (nrow(tbl_15_w) > 0 && ncol(tbl_15_w) >= 4) {
      suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_15_w[[4]]))))
    } else rep(NA, length(w_dates))

    # CPI-adjusted
    cpi_total_pct <- if (nrow(tbl_cpi) > 0 && ncol(tbl_cpi) >= 5) {
      suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[5]]))))
    } else rep(NA, length(w_dates))
    cpi_reg_pct <- if (nrow(tbl_cpi) > 0 && ncol(tbl_cpi) >= 9) {
      suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[9]]))))
    } else rep(NA, length(w_dates))

    # Build wages chart data — limit to reasonable recent history (from 2008)
    valid_idx <- which(!is.na(w_dates) & w_dates >= as.Date("2008-01-01"))
    if (length(valid_idx) > 0) {
      wchart_df <- data.frame(
        Date = format(w_dates[valid_idx], "%b %Y"),
        Total_Pay_YoY_pct = w_total[valid_idx],
        Regular_Pay_YoY_pct = if (length(w_reg) >= max(valid_idx)) w_reg[valid_idx] else NA,
        stringsAsFactors = FALSE
      )
      writeData(wb, "Wages charts", data.frame(V1 = "Average Weekly Earnings — year-on-year growth (%)"),
                startRow = 1, colNames = FALSE)
      addStyle(wb, "Wages charts", .style_title(), rows = 1, cols = 1)
      writeData(wb, "Wages charts", wchart_df, startRow = 3, headerStyle = .style_header())
    }
  }

  # --- Emp, Unemp & Inac Chart ---
  addWorksheet(wb, "Emp, Unemp & Inac Chart", tabColour = "#A9D18E")
  if (nrow(tbl_1) > 0 && ncol(tbl_1) >= 19) {
    col1 <- trimws(as.character(tbl_1[[1]]))
    lfs_pat <- "^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\s+(\\d{4})$"
    lfs_rows <- grep(lfs_pat, col1, ignore.case = TRUE)
    if (length(lfs_rows) > 0) {
      # Filter from 2010 onwards
      periods <- col1[lfs_rows]
      emp_rt <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_1[[17]][lfs_rows]))))
      unemp_rt <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_1[[9]][lfs_rows]))))
      inact_rt <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_1[[19]][lfs_rows]))))

      chart_df <- data.frame(
        Period = periods,
        Employment_rate_16_64 = emp_rt,
        Unemployment_rate_16_plus = unemp_rt,
        Inactivity_rate_16_64 = inact_rt,
        stringsAsFactors = FALSE
      )
      # Limit to ~2010 onwards
      recent_idx <- which(grepl("201[0-9]|202[0-9]", chart_df$Period))
      if (length(recent_idx) > 0) chart_df <- chart_df[min(recent_idx):nrow(chart_df), ]

      writeData(wb, "Emp, Unemp & Inac Chart",
                data.frame(V1 = "Employment, Unemployment & Inactivity rates (%)"),
                startRow = 1, colNames = FALSE)
      addStyle(wb, "Emp, Unemp & Inac Chart", .style_title(), rows = 1, cols = 1)
      writeData(wb, "Emp, Unemp & Inac Chart", chart_df, startRow = 3, headerStyle = .style_header())
    }
  }

  # --- Payrolled Employees Chart ---
  addWorksheet(wb, "Payrolled Employees Chart", tabColour = "#A9D18E")
  if (exists("pay_df") && nrow(pay_df) > 0) {
    pe_chart <- data.frame(
      Date = format(pay_df$m, "%b %Y"),
      Payrolled_employees = pay_df$v,
      stringsAsFactors = FALSE
    )
    writeData(wb, "Payrolled Employees Chart",
              data.frame(V1 = "HMRC Payrolled Employees, seasonally adjusted, UK"),
              startRow = 1, colNames = FALSE)
    addStyle(wb, "Payrolled Employees Chart", .style_title(), rows = 1, cols = 1)
    writeData(wb, "Payrolled Employees Chart", pe_chart, startRow = 3, headerStyle = .style_header())
  }

  # --- Employee levels - LFS,RTI,WFJ ---
  addWorksheet(wb, "Employee levels - LFS,RTI,WFJ", tabColour = "#A9D18E")
  writeData(wb, "Employee levels - LFS,RTI,WFJ",
            data.frame(V1 = "Comparative employee counts: LFS, RTI (PAYE), Workforce Jobs"),
            startRow = 1, colNames = FALSE)
  addStyle(wb, "Employee levels - LFS,RTI,WFJ", .style_title(), rows = 1, cols = 1)
  # Build comparison from available sources
  comp_rows <- list()
  if (exists("pay_df") && nrow(pay_df) > 0) {
    for (i in seq_len(nrow(pay_df))) {
      comp_rows[[length(comp_rows) + 1]] <- data.frame(
        Date = format(pay_df$m[i], "%b %Y"),
        RTI_Payrolled = pay_df$v[i],
        stringsAsFactors = FALSE
      )
    }
  }
  if (length(comp_rows) > 0) {
    comp_df <- do.call(rbind, comp_rows)
    writeData(wb, "Employee levels - LFS,RTI,WFJ", comp_df, startRow = 3, headerStyle = .style_header())
  }

  # --- Vacancy and redundancy charts ---
  addWorksheet(wb, "Vacancy and redundancy charts", tabColour = "#A9D18E")
  writeData(wb, "Vacancy and redundancy charts",
            data.frame(V1 = "Vacancies, Unemployment and Redundancies"),
            startRow = 1, colNames = FALSE)
  addStyle(wb, "Vacancy and redundancy charts", .style_title(), rows = 1, cols = 1)
  if (nrow(tbl_19) > 0 && ncol(tbl_19) >= 3) {
    vac_periods <- trimws(as.character(tbl_19[[1]]))
    vac_vals    <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_19[[3]]))))
    valid_vac <- which(!is.na(vac_vals) & grepl("-", vac_periods))
    if (length(valid_vac) > 0) {
      vac_chart <- data.frame(
        Period = vac_periods[valid_vac],
        Vacancies_000s = vac_vals[valid_vac],
        stringsAsFactors = FALSE
      )
      writeData(wb, "Vacancy and redundancy charts", vac_chart, startRow = 3, headerStyle = .style_header())
    }
  }

  # ========================================================================
  # 39. INTERNATIONAL COMPARISONS (long time series from A01 or OECD)
  # ========================================================================

  if (has_oecd) {
    # Already added individual sheets above; add a combined long-series sheet
    addWorksheet(wb, "International Comparisons", tabColour = "#2F5496")
    writeData(wb, "International Comparisons",
              data.frame(V1 = "International comparison of unemployment and inactivity rates (long time series)"),
              startRow = 1, colNames = FALSE)
    addStyle(wb, "International Comparisons", .style_title(), rows = 1, cols = 1)
    writeData(wb, "International Comparisons",
              data.frame(Note = "See individual Unemployment, Employment, and Inactivity sheets for detailed OECD data."),
              startRow = 3)
  } else {
    addWorksheet(wb, "International Comparisons", tabColour = "#2F5496")
    writeData(wb, "International Comparisons",
              data.frame(Note = "Upload OECD data files to populate this sheet."))
  }

  # ========================================================================
  # 40. REGIONAL BREAKDOWNS
  # ========================================================================

  addWorksheet(wb, "Regional breakdowns", tabColour = "#843C0C")

  tbl_22 <- .safe_read(file_a01, "22")
  if (nrow(tbl_22) > 0) {
    writeData(wb, "Regional breakdowns",
              data.frame(V1 = "Regional Labour Market — Employment, Unemployment, Inactivity"),
              startRow = 1, colNames = FALSE)
    addStyle(wb, "Regional breakdowns", .style_title(), rows = 1, cols = 1)
    writeData(wb, "Regional breakdowns", tbl_22, startRow = 3, colNames = FALSE)
  } else {
    writeData(wb, "Regional breakdowns",
              data.frame(Note = "A01 Sheet 22 (regional data) not available."))
  }

  # ========================================================================
  # SAVE
  # ========================================================================

  if (verbose) message("[audit wb] Saving to ", output_path)
  saveWorkbook(wb, output_path, overwrite = TRUE)
  if (verbose) message("[audit wb] Done — ", length(worksheetOrder(wb)), " sheets created")

  invisible(output_path)
}
