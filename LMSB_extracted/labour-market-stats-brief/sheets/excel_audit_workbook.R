# excel_audit_workbook.R
# Creates the "LM Stats" audit workbook from uploaded ONS Excel files.
#
# Approach: Pure R — reads with readxl, writes with openxlsx.
# Each source sheet gets clean formatting and comparison summary rows
# matching the reference workbook layout.

suppressPackageStartupMessages({
  library(openxlsx)
  library(readxl)
  library(lubridate)
})

if (!exists("parse_manual_month", inherits = TRUE)) {
  source("utils/helpers.R")
}

# ============================================================================
# HELPERS (for Dashboard computation — still needs readxl for metric extraction)
# ============================================================================

.safe_read <- function(path, sheet, ...) {
  if (is.null(path)) return(data.frame())
  tryCatch(
    suppressMessages(read_excel(path, sheet = sheet, col_names = FALSE, ...)),
    error = function(e) data.frame()
  )
}

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

.lfs_label <- function(end_date) {
  start_date <- end_date %m-% months(2)
  sprintf("%s-%s %s", format(start_date, "%b"), format(end_date, "%b"), format(end_date, "%Y"))
}

.find_row <- function(tbl, label) {
  if (nrow(tbl) == 0 || ncol(tbl) == 0) return(NA_integer_)
  col1 <- trimws(as.character(tbl[[1]]))
  idx <- which(tolower(col1) == tolower(trimws(label)))
  if (length(idx) == 0) NA_integer_ else idx[1]
}

.cell_num <- function(tbl, row, col) {
  if (is.na(row) || row < 1 || row > nrow(tbl) || col > ncol(tbl)) return(NA_real_)
  suppressWarnings(as.numeric(gsub("[^0-9.eE+-]", "", as.character(tbl[[col]][row]))))
}

.val_by_date <- function(df_m, df_v, target_date) {
  idx <- which(df_m == target_date)
  if (length(idx) == 0) return(NA_real_)
  df_v[idx[1]]
}

.avg_by_dates <- function(df_m, df_v, target_dates) {
  vals <- vapply(target_dates, function(d) .val_by_date(df_m, df_v, d), numeric(1))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

.lfs_metric <- function(tbl, col, labels) {
  rows <- vapply(labels, function(l) .find_row(tbl, l), integer(1))
  vals <- vapply(seq_along(rows), function(i) .cell_num(tbl, rows[i], col), numeric(1))
  names(vals) <- c("cur", "q", "y", "covid", "elec")
  list(cur = vals["cur"], dq = vals["cur"] - vals["q"],
       dy = vals["cur"] - vals["y"], dc = vals["cur"] - vals["covid"],
       de = vals["cur"] - vals["elec"])
}

# openxlsx style helpers
.hs <- function() createStyle(fontName = "Arial", fontSize = 10, fontColour = "#FFFFFF",
                               fgFill = "#366092", halign = "center", textDecoration = "bold",
                               border = "TopBottomLeftRight", borderColour = "#244062")
.ts <- function() createStyle(fontName = "Arial", fontSize = 14, textDecoration = "bold", fontColour = "#1F4E79")
.ss <- function() createStyle(fontName = "Arial", fontSize = 11, textDecoration = "bold", fontColour = "#505050")
.pos <- function() createStyle(fontColour = "#006100", fgFill = "#C6EFCE")
.neg <- function() createStyle(fontColour = "#9C0006", fgFill = "#FFC7CE")
.sep <- function() createStyle(fontSize = 16, textDecoration = "bold", fontColour = "#1F4E79",
                                fgFill = "#D9E2F3", halign = "center", valign = "center")

# Comparison row styles
.cmp_label <- function() createStyle(fontName = "Arial", fontSize = 10, textDecoration = "bold")
.cmp_sep   <- function() createStyle(border = "Bottom", borderColour = "#366092", borderStyle = "medium")
.id_code   <- function() createStyle(fontName = "Arial", fontSize = 9, fontColour = "#808080",
                                      textDecoration = "italic")
.date_fmt  <- function() createStyle(numFmt = "MMM-YY")
.num_fmt   <- function() createStyle(numFmt = "#,##0")
.pct_fmt   <- function() createStyle(numFmt = "0.0%")
.pp_fmt    <- function() createStyle(numFmt = "0.0")
.pp2_fmt   <- function() createStyle(numFmt = "0.00")
.gbp_fmt   <- function() createStyle(numFmt = "\"\\u00a3\"#,##0")
.data_font <- function() createStyle(fontName = "Arial", fontSize = 10)

# ============================================================================
# SHEET WRITING HELPERS
# ============================================================================

# Write a source sheet: read from Excel, write to workbook with clean formatting
.write_source_sheet <- function(wb, sheet_name, source_path, source_sheet,
                                 tab_colour = "#4472C4", start_row = 1,
                                 date_col = NULL, date_fmt_str = "MMM-YY") {
  tbl <- .safe_read(source_path, source_sheet)
  if (nrow(tbl) == 0) return(invisible(NULL))

  addWorksheet(wb, sheet_name, tabColour = tab_colour)

  # Fix date columns before writing
  if (!is.null(date_col) && date_col <= ncol(tbl)) {
    x <- tbl[[date_col]]
    if (inherits(x, c("POSIXct", "POSIXt"))) {
      tbl[[date_col]] <- as.Date(x)
    } else if (is.numeric(x)) {
      tbl[[date_col]] <- as.Date(x, origin = "1899-12-30")
    }
  }

  # Coerce character columns to numeric where possible
  for (ci in seq_len(ncol(tbl))) {
    if (!is.null(date_col) && ci == date_col) next
    vals <- tbl[[ci]]
    if (is.character(vals)) {
      num_vals <- suppressWarnings(as.numeric(vals))
      n_non_na <- sum(!is.na(vals) & nchar(trimws(vals)) > 0)
      n_numeric <- sum(!is.na(num_vals) & !is.na(vals) & nchar(trimws(vals)) > 0)
      if (n_non_na > 0 && n_numeric / n_non_na > 0.5) {
        tbl[[ci]] <- num_vals
      }
    }
  }

  writeData(wb, sheet_name, tbl, colNames = FALSE, startRow = start_row)

  # Apply date format to date column
  if (!is.null(date_col) && date_col <= ncol(tbl)) {
    date_rows <- which(!is.na(tbl[[date_col]])) + start_row - 1
    if (length(date_rows) > 0) {
      addStyle(wb, sheet_name, createStyle(numFmt = date_fmt_str),
               rows = date_rows, cols = date_col, gridExpand = TRUE, stack = TRUE)
    }
  }

  # Apply base font
  if (nrow(tbl) > 0) {
    addStyle(wb, sheet_name, .data_font(),
             rows = start_row:(start_row + nrow(tbl) - 1),
             cols = 1:ncol(tbl), gridExpand = TRUE, stack = TRUE)
  }

  # Auto-size columns (capped at 25 chars)
  for (ci in seq_len(ncol(tbl))) {
    max_width <- max(nchar(as.character(tbl[[ci]])), na.rm = TRUE)
    max_width <- min(max(max_width, 8), 25)
    setColWidths(wb, sheet_name, cols = ci, widths = max_width + 2)
  }

  invisible(tbl)
}

# Write comparison header rows at top of a sheet
.write_cmp_rows <- function(wb, sheet_name, labels, values_matrix, start_row = 1,
                             col_offset = 1) {
  # labels: character vector of row labels (e.g. "Current", "Change on quarter")
  # values_matrix: matrix or list of vectors, one per label row
  for (i in seq_along(labels)) {
    r <- start_row + i - 1
    writeData(wb, sheet_name, labels[i], startRow = r, startCol = 1)
    addStyle(wb, sheet_name, .cmp_label(), rows = r, cols = 1, stack = TRUE)
    if (!is.null(values_matrix) && length(values_matrix) >= i) {
      vals <- values_matrix[[i]]
      if (!is.null(vals) && length(vals) > 0) {
        for (j in seq_along(vals)) {
          writeData(wb, sheet_name, vals[j], startRow = r, startCol = col_offset + j)
        }
      }
    }
  }
  # Separator line below comparison rows
  last_row <- start_row + length(labels) - 1
  n_cols <- if (!is.null(values_matrix) && length(values_matrix) > 0) {
    col_offset + max(sapply(values_matrix, length))
  } else { 5 }
  addStyle(wb, sheet_name, .cmp_sep(), rows = last_row, cols = 1:n_cols,
           gridExpand = TRUE, stack = TRUE)
}

# Auto-detect sheet name from a CLA01 or X02 file
.detect_sheet <- function(path, candidates) {
  if (is.null(path)) return(NULL)
  sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
  if (length(sheets) == 0) return(NULL)
  # Try exact match on candidates first
  for (cc in candidates) {
    if (cc %in% sheets) return(cc)
  }
  # Try case-insensitive partial match
  for (cc in candidates) {
    idx <- grep(tolower(cc), tolower(sheets), fixed = TRUE)
    if (length(idx) > 0) return(sheets[idx[1]])
  }
  # Fallback: try each sheet for actual data
  for (s in sheets) {
    tbl <- tryCatch(
      suppressMessages(readxl::read_excel(path, sheet = s, col_names = FALSE, n_max = 5)),
      error = function(e) data.frame()
    )
    if (nrow(tbl) > 0 && ncol(tbl) > 1) return(s)
  }
  sheets[1]
}


# ============================================================================
# MAIN FUNCTION
# ============================================================================

create_audit_workbook <- function(
    output_path,
    file_a01 = NULL, file_hr1 = NULL, file_x09 = NULL, file_rtisa = NULL,
    file_cla01 = NULL, file_x02 = NULL,
    file_oecd_unemp = NULL, file_oecd_emp = NULL, file_oecd_inact = NULL,
    calculations_path = NULL, config_path = NULL,
    vacancies_mode = "aligned", payroll_mode = "aligned",
    manual_month_override = NULL, verbose = FALSE
) {

  # ==========================================================================
  # STEP 1: Create workbook and write all source sheets (pure R)
  # ==========================================================================

  wb <- createWorkbook()

  # Helper: write a simple source sheet with no comparison rows
  .ws <- function(src, src_sheet, tgt_sheet, tab_col = "#4472C4", date_col = NULL) {
    if (is.null(src)) return()
    .write_source_sheet(wb, tgt_sheet, src, src_sheet, tab_colour = tab_col, date_col = date_col)
  }

  # --- A01 simple sheets (no comparison rows) ---
  .ws(file_a01, "1", "Sheet1", "#4472C4")
  .ws(file_a01, "3", "3", "#4472C4")
  .ws(file_a01, "19", "19", "#4472C4")
  .ws(file_a01, "22", "22", "#843C0C")

  # --- RTISA simple sheets ---
  .ws(file_rtisa, "6. Employee flows (UK)", "RTI. Employee flows (UK)", "#548235")

  # --- HR1 sheets (1b, 2a, 2b are simple; 1a gets comparison rows in Step 3) ---
  for (s in c("1b", "2a", "2b")) .ws(file_hr1, s, s, "#C00000")

  # --- CLA01 ---
  cla_sheet <- .detect_sheet(file_cla01, c("1", "People SA", "People", "People_SA", "Sheet1", "CLA01"))
  if (!is.null(cla_sheet)) .ws(file_cla01, cla_sheet, "1 UK", "#2F5496")

  # --- X02 ---
  x02_sheet <- .detect_sheet(file_x02, c("LFS Labour market flows SA", "People SA", "1"))
  if (!is.null(x02_sheet)) .ws(file_x02, x02_sheet, "LFS Labour market flows SA", "#2F5496")

  # --- OECD files ---
  for (oecd_info in list(
    list(file = file_oecd_unemp, name = "Unemployment"),
    list(file = file_oecd_emp,   name = "Employment"),
    list(file = file_oecd_inact, name = "Inactivity")
  )) {
    if (!is.null(oecd_info$file)) {
      ext <- tolower(tools::file_ext(oecd_info$file))
      if (ext == "csv") {
        csv_data <- tryCatch(read.csv(oecd_info$file, stringsAsFactors = FALSE),
                             error = function(e) data.frame())
        if (nrow(csv_data) > 0) {
          addWorksheet(wb, oecd_info$name, tabColour = "#2F5496")
          writeData(wb, oecd_info$name, csv_data, headerStyle = .hs())
        }
      } else {
        oecd_sh <- .detect_sheet(oecd_info$file, c(oecd_info$name, "Sheet1", "Data"))
        if (!is.null(oecd_sh)) .ws(oecd_info$file, oecd_sh, oecd_info$name, "#2F5496")
      }
    }
  }

  # ==========================================================================
  # STEP 2: Detect reference period (needed for comparison rows)
  # ==========================================================================

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
        if (!is.na(end_mon) && !is.na(end_yr))
          anchor_m <- as.Date(sprintf("%04d-%02d-01", end_yr, end_mon))
      }
    }
  }
  if (is.null(anchor_m)) anchor_m <- Sys.Date() %m-% months(2)

  cm <- anchor_m %m+% months(2)
  lab_cur   <- .lfs_label(anchor_m)
  lab_q     <- .lfs_label(anchor_m %m-% months(3))
  lab_y     <- .lfs_label(anchor_m %m-% months(12))
  lab_covid <- .lfs_label(as.Date("2020-02-01"))
  lab_elec  <- .lfs_label(as.Date("2024-06-01"))
  all_labels <- c(lab_cur, lab_q, lab_y, lab_covid, lab_elec)
  ref_label <- format(cm, "%B %Y")

  # Compute metrics for Dashboard
  tbl_1 <- .safe_read(file_a01, "1")
  tbl_2 <- .safe_read(file_a01, "2")
  tbl_19 <- .safe_read(file_a01, "19")

  na_m <- list(cur = NA, dq = NA, dy = NA, dc = NA, de = NA)

  if (nrow(tbl_1) > 0) {
    m_emp16   <- .lfs_metric(tbl_1, 4,  all_labels)
    m_emprt   <- .lfs_metric(tbl_1, 17, all_labels)
    m_unemp16 <- .lfs_metric(tbl_1, 5,  all_labels)
    m_unemprt <- .lfs_metric(tbl_1, 9,  all_labels)
    m_inact   <- .lfs_metric(tbl_1, 15, all_labels)
    m_inactrt <- .lfs_metric(tbl_1, 19, all_labels)
  } else {
    m_emp16 <- m_emprt <- m_unemp16 <- m_unemprt <- m_inact <- m_inactrt <- na_m
  }

  if (nrow(tbl_2) > 0 && ncol(tbl_2) >= 57) {
    m_5064 <- .lfs_metric(tbl_2, 56, all_labels)
    m_5064rt <- .lfs_metric(tbl_2, 57, all_labels)
  } else {
    m_5064 <- m_5064rt <- na_m
  }

  # Vacancies
  vac_m <- na_m
  if (nrow(tbl_19) > 0 && ncol(tbl_19) >= 3) {
    vac_m$cur <- .cell_num(tbl_19, .find_row(tbl_19, lab_cur), 3)
    vac_m$dq  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, .lfs_label(anchor_m %m-% months(3))), 3)
    vac_m$dy  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, .lfs_label(anchor_m %m-% months(12))), 3)
    vac_m$dc  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, "Jan-Mar 2020"), 3)
    vac_m$de  <- vac_m$cur - .cell_num(tbl_19, .find_row(tbl_19, .lfs_label(as.Date("2024-06-01"))), 3)
  }

  # Payroll
  pay_m <- na_m
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
      mc <- c(cm %m-% months(4), cm %m-% months(3), cm %m-% months(2))
      mp <- c(cm %m-% months(7), cm %m-% months(6), cm %m-% months(5))
      my <- mc %m-% months(12)
      pc <- .avg_by_dates(pay_df$m, pay_df$v, mc)
      pp <- .avg_by_dates(pay_df$m, pay_df$v, mp)
      py <- .avg_by_dates(pay_df$m, pay_df$v, my)
      pcov <- .avg_by_dates(pay_df$m, pay_df$v, as.Date(c("2019-12-01", "2020-01-01", "2020-02-01")))
      pelec <- .avg_by_dates(pay_df$m, pay_df$v, as.Date(c("2024-04-01", "2024-05-01", "2024-06-01")))
      pay_m$cur <- if (!is.na(pc)) pc / 1000 else NA
      pay_m$dq  <- if (!is.na(pc) && !is.na(pp)) (pc - pp) / 1000 else NA
      pay_m$dy  <- if (!is.na(pc) && !is.na(py)) (pc - py) / 1000 else NA
      pay_m$dc  <- if (!is.na(pc) && !is.na(pcov)) (pc - pcov) / 1000 else NA
      pay_m$de  <- if (!is.na(pc) && !is.na(pelec)) (pc - pelec) / 1000 else NA
    }
  }

  # Wages nominal
  wages_m <- na_m
  tbl_13 <- .safe_read(file_a01, "13")
  if (nrow(tbl_13) > 0 && ncol(tbl_13) >= 4) {
    w13_dates <- .detect_dates(tbl_13[[1]])
    w13_pct <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[4]]))))
    w13_weekly <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_13[[2]]))))
    wages_m$cur <- .val_by_date(w13_dates, w13_pct, anchor_m)
    win3 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
    .wc <- function(a, b) {
      va <- .avg_by_dates(w13_dates, w13_weekly, a); vb <- .avg_by_dates(w13_dates, w13_weekly, b)
      if (is.na(va) || is.na(vb)) NA else (va - vb) * 52
    }
    wages_m$dq <- .wc(win3, c(anchor_m %m-% months(3), anchor_m %m-% months(4), anchor_m %m-% months(5)))
    wages_m$dy <- .wc(win3, win3 %m-% months(12))
    wages_m$dc <- .wc(win3, as.Date(c("2019-12-01", "2020-01-01", "2020-02-01")))
    wages_m$de <- .wc(win3, as.Date(c("2024-04-01", "2024-05-01", "2024-06-01")))
  }

  # Wages CPI
  wages_cpi_m <- na_m
  tbl_cpi <- .safe_read(file_x09, "AWE Real_CPI")
  if (nrow(tbl_cpi) > 0 && ncol(tbl_cpi) >= 9) {
    cpi_months <- .detect_dates(tbl_cpi[[1]])
    cpi_real <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[2]]))))
    cpi_total <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[5]]))))
    cpi_valid <- which(!is.na(cpi_months) & !is.na(cpi_total))
    ca <- if (length(cpi_valid) > 0) cpi_months[cpi_valid[length(cpi_valid)]] else anchor_m
    wages_cpi_m$cur <- .val_by_date(cpi_months, cpi_total, ca)
    cw <- c(ca, ca %m-% months(1), ca %m-% months(2))
    .cc <- function(a, b) {
      va <- .avg_by_dates(cpi_months, cpi_real, a); vb <- .avg_by_dates(cpi_months, cpi_real, b)
      if (is.na(va) || is.na(vb)) NA else (va - vb) * 52
    }
    wages_cpi_m$dq <- .cc(cw, c(ca %m-% months(3), ca %m-% months(4), ca %m-% months(5)))
    wages_cpi_m$dy <- .cc(cw, cw %m-% months(12))
    wages_cpi_m$dc <- .cc(cw, as.Date(c("2019-12-01", "2020-01-01", "2020-02-01")))
    wages_cpi_m$de <- .cc(cw, as.Date(c("2024-04-01", "2024-05-01", "2024-06-01")))
  }

  # ==========================================================================
  # STEP 3: Write complex sheets with comparison rows
  # ==========================================================================

  # --- RTISA: 1. Payrolled employees (UK) ---
  if (!is.null(file_rtisa)) {
    tbl_rtisa <- .safe_read(file_rtisa, "1. Payrolled employees (UK)")
    if (nrow(tbl_rtisa) > 0) {
      sn <- "1. Payrolled employees (UK)"
      addWorksheet(wb, sn, tabColour = "#548235")

      # Comparison rows 1-3
      writeData(wb, sn, "", startRow = 1, startCol = 1)
      for (ci in 2:6) {
        lbl <- c("Current", "Change on month (singular)", "Change on quarter",
                 "Change year on year", "Change since Covid-19")[ci - 1]
        writeData(wb, sn, lbl, startRow = 1, startCol = ci)
        addStyle(wb, sn, .hs(), rows = 1, cols = ci, stack = TRUE)
      }

      # Row 2: Number, Row 3: %
      if (exists("pay_m") && !is.null(pay_m$cur)) {
        writeData(wb, sn, "Number", startRow = 2, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = 2, cols = 1, stack = TRUE)
        cur_val <- if (!is.na(pay_m$cur)) pay_m$cur * 1000 else NA
        writeData(wb, sn, cur_val, startRow = 2, startCol = 2)

        # Single month change (latest minus previous)
        if (exists("pay_df") && nrow(pay_df) > 1) {
          latest_single <- pay_df$v[nrow(pay_df)]
          prev_single   <- pay_df$v[nrow(pay_df) - 1]
          writeData(wb, sn, latest_single - prev_single, startRow = 2, startCol = 3)
        }
        if (!is.na(pay_m$dq)) writeData(wb, sn, pay_m$dq * 1000, startRow = 2, startCol = 4)
        if (!is.na(pay_m$dy)) writeData(wb, sn, pay_m$dy * 1000, startRow = 2, startCol = 5)
        if (!is.na(pay_m$dc)) writeData(wb, sn, pay_m$dc * 1000, startRow = 2, startCol = 6)

        # Row 3: percentages
        writeData(wb, sn, "%", startRow = 3, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = 3, cols = 1, stack = TRUE)
        if (!is.na(pay_m$dq) && !is.na(pay_m$cur) && pay_m$cur != 0) {
          pc_val <- pay_m$cur * 1000
          if (exists("pay_df") && nrow(pay_df) > 1) {
            writeData(wb, sn, (pay_df$v[nrow(pay_df)] - pay_df$v[nrow(pay_df) - 1]) / pc_val,
                      startRow = 3, startCol = 3)
          }
          writeData(wb, sn, pay_m$dq / pay_m$cur, startRow = 3, startCol = 4)
          writeData(wb, sn, pay_m$dy / pay_m$cur, startRow = 3, startCol = 5)
          writeData(wb, sn, pay_m$dc / pay_m$cur, startRow = 3, startCol = 6)
        }
        addStyle(wb, sn, .pct_fmt(), rows = 3, cols = 3:6, gridExpand = TRUE, stack = TRUE)
        addStyle(wb, sn, .num_fmt(), rows = 2, cols = 2:6, gridExpand = TRUE, stack = TRUE)
      }
      addStyle(wb, sn, .cmp_sep(), rows = 3, cols = 1:6, gridExpand = TRUE, stack = TRUE)

      # Write original data from row 5
      writeData(wb, sn, tbl_rtisa, colNames = FALSE, startRow = 5)
      addStyle(wb, sn, .data_font(), rows = 5:(5 + nrow(tbl_rtisa)),
               cols = 1:ncol(tbl_rtisa), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 18)
      setColWidths(wb, sn, cols = 2:min(ncol(tbl_rtisa), 6), widths = 18)
    }
  }

  # --- RTISA: 23. Employees Industry ---
  if (!is.null(file_rtisa)) {
    tbl_23 <- .safe_read(file_rtisa, "23. Employees (Industry)")
    if (nrow(tbl_23) > 0 && ncol(tbl_23) >= 2) {
      sn <- "23. Employees Industry"
      addWorksheet(wb, sn, tabColour = "#548235")

      # Detect dates and get yearly changes
      ind_dates <- .detect_dates(tbl_23[[1]])
      ind_valid <- which(!is.na(ind_dates))
      if (length(ind_valid) > 0) {
        latest_date <- ind_dates[ind_valid[length(ind_valid)]]
        yago_date   <- latest_date %m-% months(12)
        latest_idx  <- ind_valid[length(ind_valid)]
        yago_idx    <- which(ind_dates == yago_date)

        # Header row 1: "Yearly change" + industry names
        writeData(wb, sn, "Yearly change", startRow = 1, startCol = 1)
        addStyle(wb, sn, .hs(), rows = 1, cols = 1, stack = TRUE)
        if (ncol(tbl_23) > 1) {
          # Try to find header row in original data
          for (ci in 2:min(ncol(tbl_23), 20)) {
            header_val <- as.character(tbl_23[[ci]][1])
            if (!is.na(header_val) && nchar(header_val) > 0) {
              writeData(wb, sn, header_val, startRow = 1, startCol = ci)
              addStyle(wb, sn, .hs(), rows = 1, cols = ci, stack = TRUE)
            }
          }
        }

        # Row 2: Number (yearly change), Row 3: % change
        if (length(yago_idx) > 0) {
          writeData(wb, sn, "Number", startRow = 2, startCol = 1)
          writeData(wb, sn, "%", startRow = 3, startCol = 1)
          addStyle(wb, sn, .cmp_label(), rows = 2:3, cols = 1, gridExpand = TRUE, stack = TRUE)
          for (ci in 2:min(ncol(tbl_23), 20)) {
            cur_v <- suppressWarnings(as.numeric(as.character(tbl_23[[ci]][latest_idx])))
            yr_v  <- suppressWarnings(as.numeric(as.character(tbl_23[[ci]][yago_idx[1]])))
            if (!is.na(cur_v) && !is.na(yr_v)) {
              writeData(wb, sn, cur_v - yr_v, startRow = 2, startCol = ci)
              if (yr_v != 0) writeData(wb, sn, (cur_v - yr_v) / yr_v, startRow = 3, startCol = ci)
            }
          }
          addStyle(wb, sn, .num_fmt(), rows = 2, cols = 2:min(ncol(tbl_23), 20),
                   gridExpand = TRUE, stack = TRUE)
          addStyle(wb, sn, .pct_fmt(), rows = 3, cols = 2:min(ncol(tbl_23), 20),
                   gridExpand = TRUE, stack = TRUE)
        }
        addStyle(wb, sn, .cmp_sep(), rows = 3, cols = 1:min(ncol(tbl_23), 20),
                 gridExpand = TRUE, stack = TRUE)
      }

      # Write original data from row 5
      writeData(wb, sn, tbl_23, colNames = FALSE, startRow = 5)
      addStyle(wb, sn, .data_font(), rows = 5:(5 + nrow(tbl_23)),
               cols = 1:ncol(tbl_23), gridExpand = TRUE, stack = TRUE)
    }
  }

  # --- A01 Sheet "2": Age breakdown with comparisons ---
  if (!is.null(file_a01)) {
    tbl_2_full <- .safe_read(file_a01, "2")
    if (nrow(tbl_2_full) > 0 && ncol(tbl_2_full) >= 10) {
      sn <- "2"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Section headers (rows 1-4)
      writeData(wb, sn, "Aged 16 and over", startRow = 1, startCol = 2)
      writeData(wb, sn, "Aged 16-64", startRow = 1, startCol = 10)
      addStyle(wb, sn, .hs(), rows = 1, cols = c(2, 10), stack = TRUE)

      for (ci_pair in list(c(2, "Employment"), c(4, "Unemployment"), c(6, "Activity"),
                           c(8, "Inactivity"), c(10, "Employment"))) {
        writeData(wb, sn, ci_pair[2], startRow = 2, startCol = as.integer(ci_pair[1]))
        addStyle(wb, sn, .hs(), rows = 2, cols = as.integer(ci_pair[1]), stack = TRUE)
      }
      for (ci in c(2, 4, 6, 8, 10)) {
        writeData(wb, sn, "level", startRow = 3, startCol = ci)
        writeData(wb, sn, "rate (%)", startRow = 3, startCol = ci + 1)
      }
      addStyle(wb, sn, .hs(), rows = 3, cols = 2:11, gridExpand = TRUE, stack = TRUE)

      # Comparison data rows 4-9
      cmp_labels <- c("Current", "Quarterly change", "Change year on year",
                       "Change since Covid (Dec 19-feb 20)",
                       "change since 2010 election", "change since 2024 election")
      # Get values from tbl_2_full using .lfs_metric for ALL columns
      # Columns: 2=emp16 level, 3=emp16 rate, 4=unemp level, 5=unemp rate,
      #   6=activity level, 7=activity rate, 8=inactivity level, 9=inactivity rate,
      #   10=emp16-64 level (if exists), 11=emp16-64 rate (if exists)
      lab_2010 <- "Feb-Apr 2010"
      all_labels_2 <- c(lab_cur, lab_q, lab_y, lab_covid, lab_2010, lab_elec)
      # Extract values for each data column across all comparison periods
      s2_ncol <- min(ncol(tbl_2_full), 11)
      cmp_vals <- lapply(1:6, function(cmp_idx) {
        vapply(2:s2_ncol, function(ci) {
          cur_row <- .find_row(tbl_2_full, all_labels_2[1])
          cmp_row <- .find_row(tbl_2_full, all_labels_2[cmp_idx])
          cur_v <- .cell_num(tbl_2_full, cur_row, ci)
          if (cmp_idx == 1) return(if (is.na(cur_v)) NA_real_ else cur_v)
          cmp_v <- .cell_num(tbl_2_full, cmp_row, ci)
          if (is.na(cur_v) || is.na(cmp_v)) NA_real_ else cur_v - cmp_v
        }, numeric(1))
      })
      for (i in seq_along(cmp_labels)) {
        r <- 3 + i
        writeData(wb, sn, cmp_labels[i], startRow = r, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = r, cols = 1, stack = TRUE)
        for (j in seq_along(cmp_vals[[i]])) {
          if (!is.na(cmp_vals[[i]][j])) {
            writeData(wb, sn, cmp_vals[[i]][j], startRow = r, startCol = j + 1)
          }
        }
      }
      addStyle(wb, sn, .cmp_sep(), rows = 9, cols = 1:11, gridExpand = TRUE, stack = TRUE)

      # Write original data from row 11
      writeData(wb, sn, tbl_2_full, colNames = FALSE, startRow = 11)
      addStyle(wb, sn, .data_font(), rows = 11:(11 + nrow(tbl_2_full)),
               cols = 1:ncol(tbl_2_full), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 16)
    }
  }

  # --- A01 Sheet "5": Workforce jobs ---
  if (!is.null(file_a01)) {
    tbl_5 <- .safe_read(file_a01, "5")
    if (nrow(tbl_5) > 0) {
      sn <- "5"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Headers row 1
      col_hdrs <- c("Workforce jobs", "Employee jobs", "Self-employment jobs",
                     "HM Forces", "Government-supported trainees")
      for (ci in seq_along(col_hdrs)) {
        writeData(wb, sn, col_hdrs[ci], startRow = 1, startCol = ci + 1)
        addStyle(wb, sn, .hs(), rows = 1, cols = ci + 1, stack = TRUE)
      }

      # Comparison rows — compute from data
      cmp_labels_5 <- c("Current", "Quarterly change", "Max", "Jobs created since new gov")
      for (i in seq_along(cmp_labels_5)) {
        writeData(wb, sn, cmp_labels_5[i], startRow = i + 1, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = i + 1, cols = 1, stack = TRUE)
      }
      # Find last numeric data row (skip header/metadata rows)
      s5_col1 <- trimws(as.character(tbl_5[[1]]))
      s5_data_rows <- which(grepl("^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\s+\\d{2}", s5_col1, ignore.case = TRUE))
      # Find "Jun 24" row for "new gov" baseline
      s5_newgov_row <- grep("^Jun\\s+24", s5_col1, ignore.case = TRUE)

      if (length(s5_data_rows) > 0) {
        last_dr <- s5_data_rows[length(s5_data_rows)]
        prev_q_dr <- if (length(s5_data_rows) >= 2) s5_data_rows[length(s5_data_rows) - 1] else NA
        for (ci in 2:min(ncol(tbl_5), 6)) {
          col_vals <- suppressWarnings(as.numeric(as.character(tbl_5[[ci]])))
          cur_val <- col_vals[last_dr]
          if (!is.na(cur_val)) {
            writeData(wb, sn, cur_val, startRow = 2, startCol = ci)
            # Quarterly change
            if (!is.na(prev_q_dr)) {
              prev_val <- col_vals[prev_q_dr]
              if (!is.na(prev_val)) writeData(wb, sn, cur_val - prev_val, startRow = 3, startCol = ci)
            }
            # Max
            max_val <- max(col_vals[s5_data_rows], na.rm = TRUE)
            if (is.finite(max_val)) writeData(wb, sn, max_val, startRow = 4, startCol = ci)
            # Since new gov (Jun 24)
            if (length(s5_newgov_row) > 0) {
              gov_val <- col_vals[s5_newgov_row[1]]
              if (!is.na(gov_val)) writeData(wb, sn, cur_val - gov_val, startRow = 5, startCol = ci)
            }
          }
        }
        addStyle(wb, sn, .num_fmt(), rows = 2:5, cols = 2:6, gridExpand = TRUE, stack = TRUE)
      }
      addStyle(wb, sn, .cmp_sep(), rows = 5, cols = 1:6, gridExpand = TRUE, stack = TRUE)

      # Write original data from row 7
      writeData(wb, sn, tbl_5, colNames = FALSE, startRow = 7)
      addStyle(wb, sn, .data_font(), rows = 7:(7 + nrow(tbl_5)),
               cols = 1:ncol(tbl_5), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 16)
      setColWidths(wb, sn, cols = 2:6, widths = 18)
    }
  }

  # --- A01 Sheet "10": Redundancy ---
  if (!is.null(file_a01)) {
    tbl_10_full <- .safe_read(file_a01, "10")
    if (nrow(tbl_10_full) > 0 && ncol(tbl_10_full) >= 6) {
      sn <- "10"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Headers
      for (pair in list(c(2, "People"), c(4, "Men"), c(6, "Women"))) {
        writeData(wb, sn, pair[2], startRow = 1, startCol = as.integer(pair[1]))
        addStyle(wb, sn, .hs(), rows = 1, cols = as.integer(pair[1]), stack = TRUE)
      }
      for (ci in c(2, 4, 6)) {
        writeData(wb, sn, "Level", startRow = 2, startCol = ci)
        writeData(wb, sn, "Rate per thousand", startRow = 2, startCol = ci + 1)
      }
      addStyle(wb, sn, .hs(), rows = 2, cols = 2:7, gridExpand = TRUE, stack = TRUE)

      # Comparison rows — compute from source data for all 6 columns
      cmp_labels_10 <- c("Current", "Quarterly change", "year on year change",
                          "Since pandemic", "Since 2010 election")
      all_labels_10 <- c(lab_cur, lab_q, lab_y, "Dec-Feb 2020", "Feb-Apr 2010")
      for (i in seq_along(cmp_labels_10)) {
        r <- 2 + i
        writeData(wb, sn, cmp_labels_10[i], startRow = r, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = r, cols = 1, stack = TRUE)
        for (ci in 2:min(ncol(tbl_10_full), 7)) {
          cur_row <- .find_row(tbl_10_full, all_labels_10[1])
          cmp_row <- .find_row(tbl_10_full, all_labels_10[i])
          if (i == 1) {
            cur_v <- .cell_num(tbl_10_full, cur_row, ci)
            if (!is.na(cur_v)) writeData(wb, sn, cur_v, startRow = r, startCol = ci)
          } else {
            cur_v <- .cell_num(tbl_10_full, cur_row, ci)
            cmp_v <- .cell_num(tbl_10_full, cmp_row, ci)
            if (!is.na(cur_v) && !is.na(cmp_v)) writeData(wb, sn, cur_v - cmp_v, startRow = r, startCol = ci)
          }
        }
      }
      addStyle(wb, sn, .num_fmt(), rows = 3:7, cols = c(2, 4, 6), gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sn, .pp_fmt(), rows = 3:7, cols = c(3, 5, 7), gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sn, .cmp_sep(), rows = 7, cols = 1:7, gridExpand = TRUE, stack = TRUE)

      # Original data from row 9
      writeData(wb, sn, tbl_10_full, colNames = FALSE, startRow = 9)
      addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_10_full)),
               cols = 1:ncol(tbl_10_full), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 16)
    }
  }

  # --- A01 Sheet "11": Inactivity by reason ---
  if (!is.null(file_a01)) {
    tbl_11 <- .safe_read(file_a01, "11")
    if (nrow(tbl_11) > 0 && ncol(tbl_11) >= 6) {
      sn <- "11"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Headers
      writeData(wb, sn, "Economic inactivity by reason (thousands)", startRow = 1, startCol = 3)
      addStyle(wb, sn, .hs(), rows = 1, cols = 3, stack = TRUE)
      reason_hdrs <- c("Total economically inactive aged 16-64 (thousands)",
                        "Student", "Looking after family / home", "Temp sick", "Long-term sick",
                        "Discouraged workers", "Retired", "Other")
      n_cols_11 <- min(length(reason_hdrs), ncol(tbl_11) - 1)
      for (ci in seq_len(n_cols_11)) {
        writeData(wb, sn, reason_hdrs[ci], startRow = 2, startCol = ci + 1)
        addStyle(wb, sn, .hs(), rows = 2, cols = ci + 1, stack = TRUE)
      }

      # Comparison rows: extract from source data for ALL columns
      # Use LFS labels that include "Feb-Apr 2010" for 2010 election
      all_labels_11 <- c(lab_cur, lab_q, lab_y, "Dec-Feb 2020", "Feb-Apr 2010")
      cmp_labels_11 <- c("Current", "Quarterly change", "year on year change",
                          "Since pandemic", "Since 2010 election")
      for (i in seq_along(cmp_labels_11)) {
        r <- 2 + i
        writeData(wb, sn, cmp_labels_11[i], startRow = r, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = r, cols = 1, stack = TRUE)
      }

      # For each data column, compute comparisons
      for (ci in 2:min(ncol(tbl_11), n_cols_11 + 1)) {
        cur_row <- .find_row(tbl_11, lab_cur)
        q_row   <- .find_row(tbl_11, lab_q)
        y_row   <- .find_row(tbl_11, lab_y)
        cov_row <- .find_row(tbl_11, "Dec-Feb 2020")
        e10_row <- .find_row(tbl_11, "Feb-Apr 2010")
        cur_v <- .cell_num(tbl_11, cur_row, ci)
        if (!is.na(cur_v)) {
          writeData(wb, sn, cur_v, startRow = 3, startCol = ci)
          q_v <- .cell_num(tbl_11, q_row, ci)
          if (!is.na(q_v)) writeData(wb, sn, cur_v - q_v, startRow = 4, startCol = ci)
          y_v <- .cell_num(tbl_11, y_row, ci)
          if (!is.na(y_v)) writeData(wb, sn, cur_v - y_v, startRow = 5, startCol = ci)
          cov_v <- .cell_num(tbl_11, cov_row, ci)
          if (!is.na(cov_v)) writeData(wb, sn, cur_v - cov_v, startRow = 6, startCol = ci)
          e10_v <- .cell_num(tbl_11, e10_row, ci)
          if (!is.na(e10_v)) writeData(wb, sn, cur_v - e10_v, startRow = 7, startCol = ci)
        }
      }
      addStyle(wb, sn, .num_fmt(), rows = 3:7, cols = 2:min(ncol(tbl_11), n_cols_11 + 1),
               gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sn, .cmp_sep(), rows = 7, cols = 1:min(ncol(tbl_11), n_cols_11 + 1),
               gridExpand = TRUE, stack = TRUE)

      # Original data from row 9
      writeData(wb, sn, tbl_11, colNames = FALSE, startRow = 9)
      addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_11)),
               cols = 1:ncol(tbl_11), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 16)
      setColWidths(wb, sn, cols = 2:min(ncol(tbl_11), n_cols_11 + 1), widths = 20)
    }
  }

  # --- A01 Sheet "13": AWE Total Pay (nominal) with comparisons ---
  if (!is.null(file_a01) && exists("tbl_13") && nrow(tbl_13) > 0) {
    sn <- "13"
    addWorksheet(wb, sn, tabColour = "#4472C4")

    # Sector headers (rows 1-3 matching reference)
    for (pair in list(c(2, "Whole Economy"), c(5, "Private sector"), c(8, "Public sector"),
                      c(11, "Services, SIC 2007 sections G-S"))) {
      writeData(wb, sn, pair[2], startRow = 1, startCol = as.integer(pair[1]))
      addStyle(wb, sn, .hs(), rows = 1, cols = as.integer(pair[1]), stack = TRUE)
    }
    for (ci in c(2, 5, 8, 11)) {
      writeData(wb, sn, "Weekly Earnings (\u00a3)", startRow = 2, startCol = ci)
    }
    for (ci in c(3, 6, 9)) {
      writeData(wb, sn, "% changes year on year", startRow = 2, startCol = ci)
    }
    addStyle(wb, sn, .hs(), rows = 2, cols = 2:11, gridExpand = TRUE, stack = TRUE)
    for (ci in c(3, 6, 9)) {
      writeData(wb, sn, "Single month", startRow = 3, startCol = ci)
      writeData(wb, sn, "3 month average", startRow = 3, startCol = ci + 1)
    }
    addStyle(wb, sn, .hs(), rows = 3, cols = 2:11, gridExpand = TRUE, stack = TRUE)

    # Comparison rows 4-8
    cmp_labels_13 <- c("Current (3mo avg)", "Change on quarter", "Change year on year",
                        "Change since Covid-19", "Change since 2024 election")
    for (i in seq_along(cmp_labels_13)) {
      r <- 3 + i
      writeData(wb, sn, cmp_labels_13[i], startRow = r, startCol = 1)
      addStyle(wb, sn, .cmp_label(), rows = r, cols = 1, stack = TRUE)
    }
    # Fill in whole economy values (col B=weekly £, col D=3mo avg %)
    if (exists("w13_dates") && exists("w13_weekly") && exists("w13_pct")) {
      cur_weekly <- .val_by_date(w13_dates, w13_weekly, anchor_m)
      cur_pct    <- .val_by_date(w13_dates, w13_pct, anchor_m)
      if (!is.na(cur_weekly)) writeData(wb, sn, cur_weekly, startRow = 4, startCol = 2)
      if (!is.na(cur_pct))    writeData(wb, sn, cur_pct / 100, startRow = 4, startCol = 4)
      if (!is.na(wages_m$dq)) writeData(wb, sn, wages_m$dq, startRow = 5, startCol = 2)
      if (!is.na(wages_m$dy)) writeData(wb, sn, wages_m$dy, startRow = 6, startCol = 2)
      if (!is.na(wages_m$dc)) writeData(wb, sn, wages_m$dc, startRow = 7, startCol = 2)
      if (!is.na(wages_m$de)) writeData(wb, sn, wages_m$de, startRow = 8, startCol = 2)
    }
    addStyle(wb, sn, .gbp_fmt(), rows = 5:8, cols = 2, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sn, .pct_fmt(), rows = 4, cols = c(4, 7, 10), gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sn, .cmp_sep(), rows = 8, cols = 1:11, gridExpand = TRUE, stack = TRUE)

    # Write original data from row 9 with fixed dates
    if (is.character(tbl_13[[1]])) tbl_13[[1]] <- .detect_dates(tbl_13[[1]])
    if (inherits(tbl_13[[1]], c("POSIXct", "POSIXt"))) tbl_13[[1]] <- as.Date(tbl_13[[1]])
    if (is.numeric(tbl_13[[1]])) tbl_13[[1]] <- as.Date(tbl_13[[1]], origin = "1899-12-30")
    writeData(wb, sn, tbl_13, colNames = FALSE, startRow = 9)
    # Apply date format to col A
    date_rows_13 <- which(!is.na(tbl_13[[1]])) + 8
    if (length(date_rows_13) > 0) {
      addStyle(wb, sn, .date_fmt(), rows = date_rows_13, cols = 1, stack = TRUE)
    }
    addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_13)),
             cols = 1:ncol(tbl_13), gridExpand = TRUE, stack = TRUE)
    # Style identifier code row
    id_row_13 <- which(grepl("KAB9|KAC", as.character(tbl_13[[2]]))) + 8
    if (length(id_row_13) > 0) {
      addStyle(wb, sn, .id_code(), rows = id_row_13, cols = 1:ncol(tbl_13),
               gridExpand = TRUE, stack = TRUE)
    }
    setColWidths(wb, sn, cols = 1, widths = 14)
    setColWidths(wb, sn, cols = 2:min(ncol(tbl_13), 11), widths = 16)
  }

  # --- A01 Sheet "15": AWE Regular Pay (nominal) with comparisons ---
  if (!is.null(file_a01)) {
    tbl_15_full <- .safe_read(file_a01, "15")
    if (nrow(tbl_15_full) > 0) {
      sn <- "15"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Sector headers
      for (pair in list(c(2, "Whole Economy"), c(5, "Private sector"),
                        c(8, "Public sector"), c(11, "Services, SIC 2007 sections G-S"))) {
        writeData(wb, sn, pair[2], startRow = 1, startCol = as.integer(pair[1]))
        addStyle(wb, sn, .hs(), rows = 1, cols = as.integer(pair[1]), stack = TRUE)
      }
      for (ci in c(2, 5, 8, 11)) {
        writeData(wb, sn, "Weekly Earnings (\u00a3)", startRow = 2, startCol = ci)
      }
      for (ci in c(3, 6, 9)) {
        writeData(wb, sn, "% changes year on year", startRow = 2, startCol = ci)
      }
      addStyle(wb, sn, .hs(), rows = 2, cols = 2:11, gridExpand = TRUE, stack = TRUE)
      for (ci in c(3, 6, 9)) {
        writeData(wb, sn, "Single month", startRow = 3, startCol = ci)
        writeData(wb, sn, "3 month average", startRow = 3, startCol = ci + 1)
      }
      addStyle(wb, sn, .hs(), rows = 3, cols = 2:11, gridExpand = TRUE, stack = TRUE)

      # Comparison rows
      cmp_labels_15 <- c("Current", "Quarterly change", "year on year change",
                          "Since pandemic", "Since 2010 election")
      w15_dates_full <- .detect_dates(tbl_15_full[[1]])
      w15_weekly_full <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_15_full[[2]]))))
      w15_pct_full <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_15_full[[4]]))))

      cur_reg_weekly <- .val_by_date(w15_dates_full, w15_weekly_full, anchor_m)
      cur_reg_pct    <- .val_by_date(w15_dates_full, w15_pct_full, anchor_m)

      .reg_wc <- function(a, b) {
        va <- .avg_by_dates(w15_dates_full, w15_weekly_full, a)
        vb <- .avg_by_dates(w15_dates_full, w15_weekly_full, b)
        if (is.na(va) || is.na(vb)) NA else va - vb
      }
      win3_15 <- c(anchor_m, anchor_m %m-% months(1), anchor_m %m-% months(2))
      reg_dq <- .reg_wc(win3_15, c(anchor_m %m-% months(3), anchor_m %m-% months(4), anchor_m %m-% months(5)))
      reg_dy <- .reg_wc(win3_15, win3_15 %m-% months(12))
      reg_dc <- .reg_wc(win3_15, as.Date(c("2019-12-01", "2020-01-01", "2020-02-01")))

      for (i in seq_along(cmp_labels_15)) {
        r <- 3 + i
        writeData(wb, sn, cmp_labels_15[i], startRow = r, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = r, cols = 1, stack = TRUE)
      }
      if (!is.na(cur_reg_weekly)) writeData(wb, sn, cur_reg_weekly, startRow = 4, startCol = 2)
      if (!is.na(cur_reg_pct)) writeData(wb, sn, cur_reg_pct / 100, startRow = 4, startCol = 4)
      if (!is.na(reg_dq)) writeData(wb, sn, reg_dq, startRow = 5, startCol = 2)
      if (!is.na(reg_dy)) writeData(wb, sn, reg_dy, startRow = 6, startCol = 2)
      if (!is.na(reg_dc)) writeData(wb, sn, reg_dc, startRow = 7, startCol = 2)

      addStyle(wb, sn, .pct_fmt(), rows = 4, cols = c(4, 7, 10), gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sn, .cmp_sep(), rows = 8, cols = 1:11, gridExpand = TRUE, stack = TRUE)

      # Write original data from row 9 with fixed dates
      if (is.character(tbl_15_full[[1]])) tbl_15_full[[1]] <- .detect_dates(tbl_15_full[[1]])
      if (inherits(tbl_15_full[[1]], c("POSIXct", "POSIXt"))) tbl_15_full[[1]] <- as.Date(tbl_15_full[[1]])
      if (is.numeric(tbl_15_full[[1]])) tbl_15_full[[1]] <- as.Date(tbl_15_full[[1]], origin = "1899-12-30")
      writeData(wb, sn, tbl_15_full, colNames = FALSE, startRow = 9)
      date_rows_15 <- which(!is.na(tbl_15_full[[1]])) + 8
      if (length(date_rows_15) > 0) {
        addStyle(wb, sn, .date_fmt(), rows = date_rows_15, cols = 1, stack = TRUE)
      }
      addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_15_full)),
               cols = 1:ncol(tbl_15_full), gridExpand = TRUE, stack = TRUE)
      id_row_15 <- which(grepl("KAJ", as.character(tbl_15_full[[2]]))) + 8
      if (length(id_row_15) > 0) {
        addStyle(wb, sn, .id_code(), rows = id_row_15, cols = 1:ncol(tbl_15_full),
                 gridExpand = TRUE, stack = TRUE)
      }
      setColWidths(wb, sn, cols = 1, widths = 14)
      setColWidths(wb, sn, cols = 2:min(ncol(tbl_15_full), 11), widths = 16)
    }
  }

  # --- A01 Sheet "18": Working days lost ---
  if (!is.null(file_a01)) {
    tbl_18 <- .safe_read(file_a01, "18")
    if (nrow(tbl_18) > 0) {
      sn <- "18"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Headers
      col_hdrs_18 <- c("Working days lost (thousands)", "Number of stoppages",
                        "Workers involved (thousands)")
      for (ci in seq_along(col_hdrs_18)) {
        writeData(wb, sn, col_hdrs_18[ci], startRow = 1, startCol = ci + 1)
        addStyle(wb, sn, .hs(), rows = 1, cols = ci + 1, stack = TRUE)
      }

      # Comparison rows — 7 rows matching reference
      cmp_labels_18 <- c("Current (singular month)", "Change on quarter (3mo avg)",
                          "Change since Covid-19 (2019 average)",
                          "Change since 2024 election (3mo avg)",
                          "2019 average", "2023 average")
      for (i in seq_along(cmp_labels_18)) {
        writeData(wb, sn, cmp_labels_18[i], startRow = i + 1, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = i + 1, cols = 1, stack = TRUE)
      }

      # Compute values from Sheet 18 data (monthly date labels in col 1)
      s18_dates <- .detect_dates(tbl_18[[1]])
      s18_valid <- which(!is.na(s18_dates))
      if (length(s18_valid) >= 3) {
        for (ci in 2:min(ncol(tbl_18), 4)) {
          s18_vals <- suppressWarnings(as.numeric(as.character(tbl_18[[ci]])))
          # Current (latest single month)
          cur_18 <- s18_vals[s18_valid[length(s18_valid)]]
          if (!is.na(cur_18)) writeData(wb, sn, cur_18, startRow = 2, startCol = ci)
          # 3mo avg change (last 3 vs prior 3)
          if (length(s18_valid) >= 6) {
            last3  <- mean(s18_vals[s18_valid[(length(s18_valid)-2):length(s18_valid)]], na.rm = TRUE)
            prev3  <- mean(s18_vals[s18_valid[(length(s18_valid)-5):(length(s18_valid)-3)]], na.rm = TRUE)
            if (!is.na(last3) && !is.na(prev3)) writeData(wb, sn, last3 - prev3, startRow = 3, startCol = ci)
          }
          # 2019 average
          yr2019 <- which(format(s18_dates, "%Y") == "2019" & !is.na(s18_vals))
          avg2019 <- if (length(yr2019) > 0) mean(s18_vals[yr2019], na.rm = TRUE) else NA
          if (!is.na(avg2019)) writeData(wb, sn, avg2019, startRow = 6, startCol = ci)
          # 2023 average
          yr2023 <- which(format(s18_dates, "%Y") == "2023" & !is.na(s18_vals))
          avg2023 <- if (length(yr2023) > 0) mean(s18_vals[yr2023], na.rm = TRUE) else NA
          if (!is.na(avg2023)) writeData(wb, sn, avg2023, startRow = 7, startCol = ci)
          # Change since Covid-19 (current vs 2019 avg)
          if (!is.na(cur_18) && !is.na(avg2019)) writeData(wb, sn, cur_18 - avg2019, startRow = 4, startCol = ci)
          # Change since 2024 election (last 3mo avg vs Jul-Sep 2024 avg)
          elec_months <- as.Date(c("2024-07-01", "2024-08-01", "2024-09-01"))
          elec_vals <- vapply(elec_months, function(d) {
            idx <- which(s18_dates == d)
            if (length(idx) > 0) s18_vals[idx[1]] else NA_real_
          }, numeric(1))
          avg_elec <- if (all(!is.na(elec_vals))) mean(elec_vals) else NA
          if (!is.na(avg_elec) && length(s18_valid) >= 3) {
            last3_val <- mean(s18_vals[s18_valid[(length(s18_valid)-2):length(s18_valid)]], na.rm = TRUE)
            if (!is.na(last3_val)) writeData(wb, sn, last3_val - avg_elec, startRow = 5, startCol = ci)
          }
        }
        addStyle(wb, sn, .num_fmt(), rows = 2:7, cols = 2:4, gridExpand = TRUE, stack = TRUE)
      }
      addStyle(wb, sn, .cmp_sep(), rows = 7, cols = 1:4, gridExpand = TRUE, stack = TRUE)

      # Write original data from row 9
      writeData(wb, sn, tbl_18, colNames = FALSE, startRow = 9)
      addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_18)),
               cols = 1:ncol(tbl_18), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 40)
      setColWidths(wb, sn, cols = 2:4, widths = 20)
    }
  }

  # --- A01 Sheet "20": Vacancies & Unemployment ---
  if (!is.null(file_a01)) {
    tbl_20 <- .safe_read(file_a01, "20")
    if (nrow(tbl_20) > 0) {
      sn <- "20"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Headers
      col_hdrs_20 <- c("All Vacancies (thousands)", "Unemployment (thousands)",
                        "Number of unemployed people per vacancy")
      for (ci in seq_along(col_hdrs_20)) {
        writeData(wb, sn, col_hdrs_20[ci], startRow = 1, startCol = ci + 1)
        addStyle(wb, sn, .hs(), rows = 1, cols = ci + 1, stack = TRUE)
      }

      # Comparison rows
      cmp_labels_20 <- c("Current", "Quarterly change", "Year on year change",
                          "Pre-pandemic trend (Jan-Mar)", "Since 2024 election")
      for (i in seq_along(cmp_labels_20)) {
        r <- i + 1
        writeData(wb, sn, cmp_labels_20[i], startRow = r, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = r, cols = 1, stack = TRUE)
        if (i == 1 && !is.na(vac_m$cur)) writeData(wb, sn, vac_m$cur, startRow = r, startCol = 2)
        if (i == 2 && !is.na(vac_m$dq))  writeData(wb, sn, vac_m$dq, startRow = r, startCol = 2)
        if (i == 3 && !is.na(vac_m$dy))  writeData(wb, sn, vac_m$dy, startRow = r, startCol = 2)
        if (i == 4 && !is.na(vac_m$dc))  writeData(wb, sn, vac_m$dc, startRow = r, startCol = 2)
        if (i == 5 && !is.na(vac_m$de))  writeData(wb, sn, vac_m$de, startRow = r, startCol = 2)
      }
      addStyle(wb, sn, .num_fmt(), rows = 2:6, cols = 2:4, gridExpand = TRUE, stack = TRUE)

      # Notes
      writeData(wb, sn, "*note: covid period is different (Jan-Mar) due to reporting periods",
                startRow = 7, startCol = 1)
      addStyle(wb, sn, .cmp_sep(), rows = 7, cols = 1:4, gridExpand = TRUE, stack = TRUE)

      # Write original data from row 9
      writeData(wb, sn, tbl_20, colNames = FALSE, startRow = 9)
      addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_20)),
               cols = 1:ncol(tbl_20), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 16)
      setColWidths(wb, sn, cols = 2:4, widths = 22)
    }
  }

  # --- A01 Sheet "21": Vacancies by industry ---
  if (!is.null(file_a01)) {
    tbl_21 <- .safe_read(file_a01, "21")
    if (nrow(tbl_21) > 0 && ncol(tbl_21) >= 3) {
      sn <- "21"
      addWorksheet(wb, sn, tabColour = "#4472C4")

      # Row 1: Industry headers from source data
      for (ci in 2:min(ncol(tbl_21), 20)) {
        hdr <- as.character(tbl_21[[ci]][1])
        if (!is.na(hdr) && nchar(hdr) > 0) {
          writeData(wb, sn, hdr, startRow = 1, startCol = ci)
          addStyle(wb, sn, .hs(), rows = 1, cols = ci, stack = TRUE)
        }
      }

      # Comparison rows — lookup by LFS labels
      cmp_labels_21 <- c("Current", "Quarterly change", "year on year change",
                          "pre-pandemic trend (Dec-Feb)")
      for (i in seq_along(cmp_labels_21)) {
        writeData(wb, sn, cmp_labels_21[i], startRow = i + 1, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = i + 1, cols = 1, stack = TRUE)
      }
      # Vacancies current + quarterly from computed data
      if (!is.na(vac_m$cur)) writeData(wb, sn, vac_m$cur, startRow = 2, startCol = 3)
      if (!is.na(vac_m$dq))  writeData(wb, sn, vac_m$dq, startRow = 3, startCol = 3)
      if (!is.na(vac_m$dy))  writeData(wb, sn, vac_m$dy, startRow = 4, startCol = 3)
      if (!is.na(vac_m$dc))  writeData(wb, sn, vac_m$dc, startRow = 5, startCol = 3)
      addStyle(wb, sn, .cmp_sep(), rows = 5, cols = 1:min(ncol(tbl_21), 20),
               gridExpand = TRUE, stack = TRUE)

      # Write original data from row 7
      writeData(wb, sn, tbl_21, colNames = FALSE, startRow = 7)
      addStyle(wb, sn, .data_font(), rows = 7:(7 + nrow(tbl_21)),
               cols = 1:ncol(tbl_21), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 16)
    }
  }

  # --- X09 Sheet "AWE Real_CPI" with comparisons ---
  if (!is.null(file_x09) && exists("tbl_cpi") && nrow(tbl_cpi) > 0 && exists("cw")) {
    sn <- "AWE Real_CPI"
    addWorksheet(wb, sn, tabColour = "#BF8F00")

    # Headers (row 1)
    for (pair in list(c(2, "Total Pay Real AWE (2015 \u00a3)"),
                      c(3, "Total Pay Real AWE (%)"),
                      c(4, "Regular Pay Real AWE (2015 \u00a3)"),
                      c(5, "Regular Pay Real AWE (%)"))) {
      writeData(wb, sn, pair[2], startRow = 1, startCol = as.integer(pair[1]))
      addStyle(wb, sn, .hs(), rows = 1, cols = as.integer(pair[1]), stack = TRUE)
    }

    # Comparison rows 2-8
    cmp_labels_cpi <- c("Current (3mo avg)", "Change on quarter (3mo avg)",
                         "Change year on year (3mo avg)",
                         "Change since Covid-19 (2019 average)",
                         "Change since 2010", "Change since financial crisis",
                         "Change since 2024 election")
    for (i in seq_along(cmp_labels_cpi)) {
      r <- i + 1
      writeData(wb, sn, cmp_labels_cpi[i], startRow = r, startCol = 1)
      addStyle(wb, sn, .cmp_label(), rows = r, cols = 1, stack = TRUE)
    }

    # Fill CPI comparison values
    if (exists("cpi_months") && exists("cpi_real") && exists("cpi_total")) {
      cpi_reg <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(tbl_cpi[[9]]))))

      # Current (3mo avg)
      cur_total_real <- .avg_by_dates(cpi_months, cpi_real, cw)
      cur_reg_real   <- .avg_by_dates(cpi_months, cpi_reg, cw)
      if (!is.na(cur_total_real)) writeData(wb, sn, cur_total_real, startRow = 2, startCol = 2)
      if (!is.na(cur_reg_real))   writeData(wb, sn, cur_reg_real, startRow = 2, startCol = 4)

      # Change on quarter
      if (!is.na(wages_cpi_m$dq)) writeData(wb, sn, wages_cpi_m$dq, startRow = 3, startCol = 2)

      # Change year on year
      if (!is.na(wages_cpi_m$dy)) writeData(wb, sn, wages_cpi_m$dy, startRow = 4, startCol = 2)
      if (!is.na(wages_cpi_m$cur)) writeData(wb, sn, wages_cpi_m$cur / 100, startRow = 4, startCol = 3)

      # Change since Covid-19
      if (!is.na(wages_cpi_m$dc)) writeData(wb, sn, wages_cpi_m$dc, startRow = 5, startCol = 2)

      # Change since 2010
      yr2010_val <- .avg_by_dates(cpi_months, cpi_real,
                                   as.Date(c("2010-04-01", "2010-05-01", "2010-06-01")))
      if (!is.na(cur_total_real) && !is.na(yr2010_val)) {
        writeData(wb, sn, (cur_total_real - yr2010_val) * 52, startRow = 6, startCol = 2)
        if (yr2010_val != 0) writeData(wb, sn, (cur_total_real - yr2010_val) / yr2010_val,
                                       startRow = 6, startCol = 3)
      }

      # Change since financial crisis (Dec 2007)
      dec2007_val <- .val_by_date(cpi_months, cpi_real, as.Date("2007-12-01"))
      if (!is.na(cur_total_real) && !is.na(dec2007_val)) {
        writeData(wb, sn, (cur_total_real - dec2007_val) * 52, startRow = 7, startCol = 2)
        if (dec2007_val != 0) writeData(wb, sn, (cur_total_real - dec2007_val) / dec2007_val,
                                        startRow = 7, startCol = 3)
      }

      # Change since 2024 election
      if (!is.na(wages_cpi_m$de)) writeData(wb, sn, wages_cpi_m$de, startRow = 8, startCol = 2)
    }

    addStyle(wb, sn, .num_fmt(), rows = 2:8, cols = c(2, 4), gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sn, .pct_fmt(), rows = 2:8, cols = c(3, 5), gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sn, .cmp_sep(), rows = 8, cols = 1:9, gridExpand = TRUE, stack = TRUE)

    # Write original CPI data from row 9 with fixed dates
    if (is.character(tbl_cpi[[1]])) tbl_cpi[[1]] <- .detect_dates(tbl_cpi[[1]])
    if (inherits(tbl_cpi[[1]], c("POSIXct", "POSIXt"))) tbl_cpi[[1]] <- as.Date(tbl_cpi[[1]])
    if (is.numeric(tbl_cpi[[1]])) tbl_cpi[[1]] <- as.Date(tbl_cpi[[1]], origin = "1899-12-30")
    writeData(wb, sn, tbl_cpi, colNames = FALSE, startRow = 9)
    date_rows_cpi <- which(!is.na(tbl_cpi[[1]])) + 8
    if (length(date_rows_cpi) > 0) {
      addStyle(wb, sn, .date_fmt(), rows = date_rows_cpi, cols = 1, stack = TRUE)
    }
    addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_cpi)),
             cols = 1:ncol(tbl_cpi), gridExpand = TRUE, stack = TRUE)
    setColWidths(wb, sn, cols = 1, widths = 14)
    setColWidths(wb, sn, cols = 2:min(ncol(tbl_cpi), 9), widths = 16)
  }

  # --- HR1 Sheet "1a" with comparisons ---
  if (!is.null(file_hr1)) {
    tbl_1a <- .safe_read(file_hr1, "1a")
    if (nrow(tbl_1a) > 0 && ncol(tbl_1a) >= 2) {
      # If sheet was already created by simple .ws() call, skip
      if (!"1a" %in% names(wb)) addWorksheet(wb, "1a", tabColour = "#C00000")
      sn <- "1a"

      # Region headers (row 1) from source data
      for (ci in 2:min(ncol(tbl_1a), 12)) {
        hdr <- as.character(tbl_1a[[ci]][1])
        if (!is.na(hdr) && nchar(hdr) > 0) {
          writeData(wb, sn, hdr, startRow = 1, startCol = ci)
          addStyle(wb, sn, .hs(), rows = 1, cols = ci, stack = TRUE)
        }
      }

      # Comparison rows
      cmp_labels_1a <- c("Current", "Average since start of 2023",
                          "Average pre-covid (April 2019-February 2020)",
                          "Change on month", "Change on quarter", "Change on year")
      for (i in seq_along(cmp_labels_1a)) {
        writeData(wb, sn, cmp_labels_1a[i], startRow = i + 1, startCol = 1)
        addStyle(wb, sn, .cmp_label(), rows = i + 1, cols = 1, stack = TRUE)
      }

      # Compute values from HR1 data
      # Find where date data begins (skip metadata rows)
      hr1_dates <- .detect_dates(tbl_1a[[1]])
      hr1_valid <- which(!is.na(hr1_dates))
      if (length(hr1_valid) >= 3) {
        max_ci <- min(ncol(tbl_1a), 12)
        for (ci in 2:max_ci) {
          hr1_vals <- suppressWarnings(as.numeric(as.character(tbl_1a[[ci]])))
          n_valid <- sum(!is.na(hr1_vals[hr1_valid]))
          if (n_valid == 0) next

          # Current: average of last 3 months
          last3_idx <- hr1_valid[(length(hr1_valid) - min(2, length(hr1_valid) - 1)):length(hr1_valid)]
          cur_avg <- mean(hr1_vals[last3_idx], na.rm = TRUE)
          if (!is.na(cur_avg)) writeData(wb, sn, cur_avg, startRow = 2, startCol = ci)

          # Average since start of 2023
          since2023 <- which(hr1_dates >= as.Date("2023-01-01") & !is.na(hr1_vals))
          if (length(since2023) > 0) {
            avg2023 <- mean(hr1_vals[since2023], na.rm = TRUE)
            if (!is.na(avg2023)) writeData(wb, sn, avg2023, startRow = 3, startCol = ci)
          }

          # Average pre-covid (April 2019 - February 2020)
          precov <- which(hr1_dates >= as.Date("2019-04-01") & hr1_dates <= as.Date("2020-02-01") & !is.na(hr1_vals))
          if (length(precov) > 0) {
            avg_precov <- mean(hr1_vals[precov], na.rm = TRUE)
            if (!is.na(avg_precov)) writeData(wb, sn, avg_precov, startRow = 4, startCol = ci)
          }

          # Change on month
          if (length(hr1_valid) >= 2) {
            v_last <- hr1_vals[hr1_valid[length(hr1_valid)]]
            v_prev <- hr1_vals[hr1_valid[length(hr1_valid) - 1]]
            if (!is.na(v_last) && !is.na(v_prev) && v_prev != 0) {
              writeData(wb, sn, (v_last - v_prev) / v_prev, startRow = 5, startCol = ci)
            }
          }

          # Change on quarter (3mo avg change)
          if (length(hr1_valid) >= 6) {
            prev3_idx <- hr1_valid[(length(hr1_valid) - 5):(length(hr1_valid) - 3)]
            prev_avg <- mean(hr1_vals[prev3_idx], na.rm = TRUE)
            if (!is.na(cur_avg) && !is.na(prev_avg) && prev_avg != 0) {
              writeData(wb, sn, (cur_avg - prev_avg) / prev_avg, startRow = 6, startCol = ci)
            }
          }

          # Change on year (3mo avg vs same 3 months last year)
          yr_ago_dates <- hr1_dates[last3_idx] %m-% months(12)
          yr_vals <- vapply(yr_ago_dates, function(d) {
            idx <- which(hr1_dates == d)
            if (length(idx) > 0) hr1_vals[idx[1]] else NA_real_
          }, numeric(1))
          yr_avg <- mean(yr_vals, na.rm = TRUE)
          if (!is.na(cur_avg) && !is.na(yr_avg) && yr_avg != 0) {
            writeData(wb, sn, (cur_avg - yr_avg) / yr_avg, startRow = 7, startCol = ci)
          }
        }
        addStyle(wb, sn, .num_fmt(), rows = 2:4, cols = 2:max_ci, gridExpand = TRUE, stack = TRUE)
        addStyle(wb, sn, .pct_fmt(), rows = 5:7, cols = 2:max_ci, gridExpand = TRUE, stack = TRUE)
      }
      addStyle(wb, sn, .cmp_sep(), rows = 7, cols = 1:min(ncol(tbl_1a), 12),
               gridExpand = TRUE, stack = TRUE)

      # Write original data from row 9
      writeData(wb, sn, tbl_1a, colNames = FALSE, startRow = 9)
      addStyle(wb, sn, .data_font(), rows = 9:(9 + nrow(tbl_1a)),
               cols = 1:ncol(tbl_1a), gridExpand = TRUE, stack = TRUE)
      setColWidths(wb, sn, cols = 1, widths = 20)
    }
  }

  # ==========================================================================
  # STEP 4: Add generated sheets (Dashboard, info, separators)
  # ==========================================================================

  # --- How to update ---
  addWorksheet(wb, "How to update", tabColour = "#FFC000")
  writeData(wb, "How to update", data.frame(V1 = c(
    paste0("Labour Market Statistics Briefing \u2014 ", ref_label), "",
    "HOW TO UPDATE THIS WORKBOOK", "----------------------------",
    "This workbook is auto-generated from ONS source datasets.",
    "To update, download the latest files from ONS and upload them via the app.", "",
    "Required: A01, HR1, X09, RTISA",
    "Optional: CLA01, X02, OECD (3 files)", "",
    paste0("LFS period: ", lab_cur),
    paste0("Comparison periods: vs ", lab_q, " | vs ", lab_y, " | vs ", lab_covid, " | vs ", lab_elec)
  )), colNames = FALSE)
  addStyle(wb, "How to update", .ts(), rows = 1, cols = 1)
  setColWidths(wb, "How to update", cols = 1, widths = 80)

  # --- Data links ---
  addWorksheet(wb, "Data links", tabColour = "#FFC000")
  writeData(wb, "Data links", data.frame(
    Sheet = c("1. Payrolled employees (UK)", "23. Employees Industry", "2", "3", "5",
              "10", "11", "13", "15", "18", "20", "21", "22",
              "1 UK", "AWE Real_CPI", "1a/1b/2a/2b",
              "LFS Labour market flows SA", "RTI. Employee flows (UK)",
              "Unemployment/Employment/Inactivity", "Regional breakdowns"),
    Source = c("RTISA", "RTISA", rep("A01", 11), "CLA01", "X09", "HR1",
               "X02", "RTISA", "OECD", "A01"),
    stringsAsFactors = FALSE
  ), headerStyle = .hs())
  setColWidths(wb, "Data links", cols = 1:2, widths = c(40, 15))

  # --- Dashboard ---
  addWorksheet(wb, "Dashboard", tabColour = "#00703C")
  writeData(wb, "Dashboard", data.frame(V1 = paste0("Labour Market Dashboard \u2014 ", ref_label)),
            startRow = 1, colNames = FALSE)
  addStyle(wb, "Dashboard", .ts(), rows = 1, cols = 1)
  writeData(wb, "Dashboard", data.frame(V1 = paste0("LFS period: ", lab_cur)),
            startRow = 2, colNames = FALSE)
  addStyle(wb, "Dashboard", .ss(), rows = 2, cols = 1)

  hdrs <- c("Metric", "Current", "Change on quarter", "Change on year",
            "Change since COVID-19", "Change since election")
  writeData(wb, "Dashboard", as.data.frame(t(hdrs)), startRow = 4, colNames = FALSE)
  addStyle(wb, "Dashboard", .hs(), rows = 4, cols = 1:6, gridExpand = TRUE)

  dash_df <- data.frame(
    Metric = c("Employment 16+ (000s)", "Employment rate 16-64 (%)",
               "Unemployment 16+ (000s)", "Unemployment rate 16+ (%)",
               "Inactivity 16-64 (000s)", "Inactivity rate 16-64 (%)",
               "Inactivity 50-64 (000s)", "Inactivity rate 50-64 (%)",
               "Payrolled employees (000s)", "Vacancies (000s)",
               "Wages total pay (%)", "Wages CPI-adjusted (%)"),
    Current = c(m_emp16$cur/1000, m_emprt$cur, m_unemp16$cur/1000, m_unemprt$cur,
                m_inact$cur/1000, m_inactrt$cur, m_5064$cur/1000, m_5064rt$cur,
                pay_m$cur, vac_m$cur, wages_m$cur, wages_cpi_m$cur),
    Qtr = c(m_emp16$dq/1000, m_emprt$dq, m_unemp16$dq/1000, m_unemprt$dq,
            m_inact$dq/1000, m_inactrt$dq, m_5064$dq/1000, m_5064rt$dq,
            pay_m$dq, vac_m$dq, wages_m$dq, wages_cpi_m$dq),
    Yr = c(m_emp16$dy/1000, m_emprt$dy, m_unemp16$dy/1000, m_unemprt$dy,
           m_inact$dy/1000, m_inactrt$dy, m_5064$dy/1000, m_5064rt$dy,
           pay_m$dy, vac_m$dy, wages_m$dy, wages_cpi_m$dy),
    COVID = c(m_emp16$dc/1000, m_emprt$dc, m_unemp16$dc/1000, m_unemprt$dc,
              m_inact$dc/1000, m_inactrt$dc, m_5064$dc/1000, m_5064rt$dc,
              pay_m$dc, vac_m$dc, wages_m$dc, wages_cpi_m$dc),
    Elec = c(m_emp16$de/1000, m_emprt$de, m_unemp16$de/1000, m_unemprt$de,
             m_inact$de/1000, m_inactrt$de, m_5064$de/1000, m_5064rt$de,
             pay_m$de, vac_m$de, wages_m$de, wages_cpi_m$de),
    stringsAsFactors = FALSE
  )
  writeData(wb, "Dashboard", dash_df, startRow = 5, colNames = FALSE)

  for (ci in 3:6) for (ri in 5:16) {
    conditionalFormatting(wb, "Dashboard", cols = ci, rows = ri,
                          type = "expression", rule = ">0", style = .pos())
    conditionalFormatting(wb, "Dashboard", cols = ci, rows = ri,
                          type = "expression", rule = "<0", style = .neg())
  }
  setColWidths(wb, "Dashboard", cols = 1:6, widths = c(35, 15, 20, 18, 22, 22))

  # ==========================================================================
  # STEP 4b: Generate chart sheets (ggplot2 images)
  # ==========================================================================

  if (requireNamespace("ggplot2", quietly = TRUE)) {
    library(ggplot2)

    .govuk_theme <- function() {
      theme_minimal(base_family = "sans", base_size = 11) +
        theme(
          plot.title = element_text(face = "bold", size = 14, colour = "#1F4E79"),
          plot.subtitle = element_text(size = 10, colour = "#505050"),
          panel.grid.minor = element_blank(),
          legend.position = "bottom",
          axis.title = element_text(size = 10),
          plot.margin = margin(10, 15, 10, 10)
        )
    }

    .insert_chart <- function(wb, sheet_name, plot_obj, title, tab_col = "#B4C6E7",
                               width = 22, height = 13) {
      addWorksheet(wb, sheet_name, tabColour = tab_col)
      writeData(wb, sheet_name, title, startRow = 1, startCol = 1)
      addStyle(wb, sheet_name, .ts(), rows = 1, cols = 1)
      tmp_png <- tempfile(fileext = ".png")
      tryCatch({
        ggsave(tmp_png, plot = plot_obj, width = width, height = height, units = "cm", dpi = 200, bg = "white")
        insertImage(wb, sheet_name, tmp_png, startRow = 3, startCol = 1,
                    width = width, height = height, units = "cm")
      }, error = function(e) {
        writeData(wb, sheet_name, paste("Chart generation failed:", e$message),
                  startRow = 3, startCol = 1)
      })
      tryCatch(file.remove(tmp_png), error = function(e) NULL)
    }

    # --- Wages charts ---
    if (exists("tbl_cpi") && nrow(tbl_cpi) > 0) {
      cpi_d <- .detect_dates(tbl_cpi[[1]])
      cpi_tp <- suppressWarnings(as.numeric(as.character(tbl_cpi[[2]])))
      cpi_rp <- if (ncol(tbl_cpi) >= 6) suppressWarnings(as.numeric(as.character(tbl_cpi[[6]]))) else rep(NA, nrow(tbl_cpi))
      wage_df <- data.frame(
        Date = rep(cpi_d, 2),
        Value = c(cpi_tp, cpi_rp),
        Series = rep(c("Total Pay", "Regular Pay"), each = length(cpi_d)),
        stringsAsFactors = FALSE
      )
      wage_df <- wage_df[!is.na(wage_df$Date) & !is.na(wage_df$Value), ]
      if (nrow(wage_df) > 0) {
        p_wages <- ggplot(wage_df, aes(x = Date, y = Value, colour = Series)) +
          geom_line(linewidth = 0.8) +
          scale_colour_manual(values = c("Total Pay" = "#1F4E79", "Regular Pay" = "#C00000")) +
          labs(title = "Real Average Weekly Earnings",
               subtitle = "Constant 2015 prices, seasonally adjusted (\u00a3)",
               x = NULL, y = "Weekly earnings (\u00a3)", colour = NULL) +
          .govuk_theme()
        .insert_chart(wb, "Wages charts", p_wages, "Real Average Weekly Earnings", "#BF8F00")
      }
    }

    # --- Emp, Unemp & Inac Chart ---
    if (exists("tbl_2_full") && nrow(tbl_2_full) > 0 && ncol(tbl_2_full) >= 9) {
      # Extract rates from tbl_2_full: col1=labels, col3=emp rate, col5=unemp rate, col9=inact rate
      s2_labels <- trimws(as.character(tbl_2_full[[1]]))
      lfs_pat <- "^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\s+(\\d{4})$"
      s2_lfs <- grep(lfs_pat, s2_labels)
      if (length(s2_lfs) > 10) {
        # Parse end dates from LFS labels
        s2_dates <- vapply(s2_lfs, function(idx) {
          parts <- regmatches(s2_labels[idx], regexec(lfs_pat, s2_labels[idx]))[[1]]
          end_mon <- match(tools::toTitleCase(tolower(parts[3])), month.abb)
          end_yr <- as.integer(parts[4])
          if (!is.na(end_mon) && !is.na(end_yr)) as.Date(sprintf("%04d-%02d-01", end_yr, end_mon))
          else NA_real_
        }, numeric(1))
        s2_dates <- as.Date(s2_dates, origin = "1970-01-01")

        emp_rate <- suppressWarnings(as.numeric(as.character(tbl_2_full[[3]][s2_lfs])))
        unemp_rate <- suppressWarnings(as.numeric(as.character(tbl_2_full[[5]][s2_lfs])))
        inact_rate <- suppressWarnings(as.numeric(as.character(tbl_2_full[[9]][s2_lfs])))

        eui_df <- data.frame(
          Date = rep(s2_dates, 3),
          Rate = c(emp_rate, unemp_rate, inact_rate),
          Indicator = rep(c("Employment rate (%)", "Unemployment rate (%)", "Inactivity rate (%)"), each = length(s2_dates)),
          stringsAsFactors = FALSE
        )
        eui_df <- eui_df[!is.na(eui_df$Date) & !is.na(eui_df$Rate), ]
        # Only show last 20 years
        eui_df <- eui_df[eui_df$Date >= as.Date("2005-01-01"), ]
        if (nrow(eui_df) > 0) {
          p_eui <- ggplot(eui_df, aes(x = Date, y = Rate, colour = Indicator)) +
            geom_line(linewidth = 0.7) +
            scale_colour_manual(values = c("Employment rate (%)" = "#1F4E79",
                                           "Unemployment rate (%)" = "#C00000",
                                           "Inactivity rate (%)" = "#548235")) +
            labs(title = "Employment, Unemployment & Inactivity Rates",
                 subtitle = "16-64 age group, seasonally adjusted",
                 x = NULL, y = "Rate (%)", colour = NULL) +
            .govuk_theme()
          .insert_chart(wb, "Emp, Unemp & Inac Chart", p_eui, "Employment, Unemployment & Inactivity Rates", "#4472C4")
        }
      }
    }

    # --- Payrolled Employees Chart ---
    if (exists("pay_df") && nrow(pay_df) > 0) {
      pay_chart_df <- pay_df[pay_df$m >= as.Date("2014-01-01"), ]
      if (nrow(pay_chart_df) > 0) {
        p_pay <- ggplot(pay_chart_df, aes(x = m, y = v / 1000)) +
          geom_line(colour = "#1F4E79", linewidth = 0.8) +
          labs(title = "Payrolled Employees",
               subtitle = "UK, seasonally adjusted (thousands)",
               x = NULL, y = "Employees (thousands)") +
          .govuk_theme()
        .insert_chart(wb, "Payrolled Employees Chart", p_pay, "Payrolled Employees", "#548235")
      }
    }

    # --- Vacancy and redundancy charts ---
    vac_chart_made <- FALSE
    if (exists("tbl_20") && nrow(tbl_20) > 0 && ncol(tbl_20) >= 3) {
      # col1 = LFS labels, col2 = vacancies, col3 = unemployment
      v20_labels <- trimws(as.character(tbl_20[[1]]))
      v20_lfs <- grep(lfs_pat, v20_labels)
      if (length(v20_lfs) > 10) {
        v20_dates <- vapply(v20_lfs, function(idx) {
          parts <- regmatches(v20_labels[idx], regexec(lfs_pat, v20_labels[idx]))[[1]]
          end_mon <- match(tools::toTitleCase(tolower(parts[3])), month.abb)
          end_yr <- as.integer(parts[4])
          if (!is.na(end_mon) && !is.na(end_yr)) as.Date(sprintf("%04d-%02d-01", end_yr, end_mon))
          else NA_real_
        }, numeric(1))
        v20_dates <- as.Date(v20_dates, origin = "1970-01-01")
        vac_vals <- suppressWarnings(as.numeric(as.character(tbl_20[[2]][v20_lfs])))
        vac_chart_df <- data.frame(Date = v20_dates, Vacancies = vac_vals, stringsAsFactors = FALSE)
        vac_chart_df <- vac_chart_df[!is.na(vac_chart_df$Date) & !is.na(vac_chart_df$Vacancies), ]
        vac_chart_df <- vac_chart_df[vac_chart_df$Date >= as.Date("2005-01-01"), ]
        if (nrow(vac_chart_df) > 0) {
          p_vac <- ggplot(vac_chart_df, aes(x = Date, y = Vacancies)) +
            geom_line(colour = "#1F4E79", linewidth = 0.8) +
            labs(title = "Vacancies (thousands)",
                 subtitle = "UK, seasonally adjusted, rolling 3-month average",
                 x = NULL, y = "Vacancies (000s)") +
            .govuk_theme()
          .insert_chart(wb, "Vacancy and redundancy charts", p_vac,
                        "Vacancies & Redundancies", "#4472C4")
          vac_chart_made <- TRUE
        }
      }
    }

  } else {
    # ggplot2 not available — create placeholder chart sheets
    for (ch_name in c("Wages charts", "Emp, Unemp & Inac Chart",
                       "Payrolled Employees Chart", "Vacancy and redundancy charts")) {
      if (!ch_name %in% names(wb)) {
        addWorksheet(wb, ch_name, tabColour = "#B4C6E7")
        writeData(wb, ch_name, data.frame(Note = "Install ggplot2 package to generate charts automatically."))
      }
    }
  }

  # --- Separator sheets ---
  for (sep_name in c("Redundancies >>>", "Labour market flows >>>",
                      "International Comparisons >>>", "Charts >>>")) {
    if (!sep_name %in% names(wb)) {
      addWorksheet(wb, sep_name, tabColour = "#B4C6E7")
      writeData(wb, sep_name, data.frame(X = sep_name), startRow = 1, colNames = FALSE)
      addStyle(wb, sep_name, .sep(), rows = 1, cols = 1)
      setColWidths(wb, sep_name, cols = 1, widths = 50)
      setRowHeights(wb, sep_name, rows = 1, heights = 60)
    }
  }

  # --- Placeholder sheets for missing supplementary data ---
  for (missing_info in list(
    list(name = "1 UK", file = file_cla01, msg_null = "Upload CLA01 file to populate this sheet.",
         msg_fail = "CLA01 file uploaded but data could not be read. Check the file has a sheet named '1' or 'People SA'."),
    list(name = "LFS Labour market flows SA", file = file_x02, msg_null = "Upload X02 file to populate this sheet.",
         msg_fail = "X02 file uploaded but data could not be read. Check the file has a sheet named 'LFS Labour market flows SA'.")
  )) {
    if (!missing_info$name %in% names(wb)) {
      msg <- if (is.null(missing_info$file)) missing_info$msg_null else missing_info$msg_fail
      addWorksheet(wb, missing_info$name)
      writeData(wb, missing_info$name, data.frame(Note = msg))
    }
  }

  has_oecd <- !is.null(file_oecd_unemp) || !is.null(file_oecd_emp) || !is.null(file_oecd_inact)
  if (!has_oecd) {
    for (sn in c("Final Table", "Unemployment", "Employment", "Inactivity")) {
      if (!sn %in% names(wb)) {
        addWorksheet(wb, sn, tabColour = "#2F5496")
        writeData(wb, sn, data.frame(Note = "Upload OECD data files to populate international comparisons."))
      }
    }
  } else {
    # Build Final Table combining OECD data
    if (!"Final Table" %in% names(wb)) {
      addWorksheet(wb, "Final Table", tabColour = "#2F5496")
      writeData(wb, "Final Table", data.frame(V1 = "See Unemployment, Employment, Inactivity sheets for OECD data."),
                startRow = 1, colNames = FALSE)
    }
  }

  # Regional breakdowns (pointer)
  if (!"Regional breakdowns" %in% names(wb)) {
    addWorksheet(wb, "Regional breakdowns", tabColour = "#843C0C")
    writeData(wb, "Regional breakdowns",
              data.frame(Note = "See Sheet '22' for regional Labour Force Survey data."))
  }

  # International Comparisons (long series pointer)
  if (!"International Comparisons" %in% names(wb)) {
    addWorksheet(wb, "International Comparisons", tabColour = "#2F5496")
    writeData(wb, "International Comparisons",
              data.frame(Note = "See Unemployment, Employment, Inactivity sheets."))
  }

  # ==========================================================================
  # STEP 5: Reorder sheets to match reference workbook
  # ==========================================================================

  desired_order <- c(
    "How to update", "Data links", "Dashboard",
    "1. Payrolled employees (UK)", "23. Employees Industry",
    "2", "Sheet1", "3", "5", "10", "11", "13", "15", "18", "21", "20", "22",
    "1 UK", "AWE Real_CPI",
    "Redundancies >>>", "1a", "1b", "2a", "2b",
    "Labour market flows >>>", "LFS Labour market flows SA", "RTI. Employee flows (UK)",
    "International Comparisons >>>", "Final Table", "Unemployment", "Employment", "Inactivity",
    "Charts >>>",
    "Wages charts", "Emp, Unemp & Inac Chart", "Payrolled Employees Chart",
    "Vacancy and redundancy charts",
    "International Comparisons", "Regional breakdowns"
  )

  current_sheets <- names(wb)
  # Build reorder: desired sheets that exist, then any extras not in the desired list
  new_order <- c()
  for (s in desired_order) {
    idx <- which(current_sheets == s)
    if (length(idx) > 0) new_order <- c(new_order, idx[1])
  }
  extras <- setdiff(seq_along(current_sheets), new_order)
  new_order <- c(new_order, extras)

  worksheetOrder(wb) <- new_order

  # ==========================================================================
  # STEP 6: Save
  # ==========================================================================

  if (verbose) message("[audit wb] Saving ", length(current_sheets), " sheets to ", output_path)
  saveWorkbook(wb, output_path, overwrite = TRUE)
  if (verbose) message("[audit wb] Done")

  invisible(output_path)
}
