# excel_audit_workbook.R
# Creates the "LM Stats" audit workbook from uploaded ONS Excel files.
#
# Approach:
#   1. Python (openpyxl) copies source sheets with FULL formatting preserved
#   2. openxlsx adds Dashboard, info sheets, separators on top
#   3. Sheets reordered to match the reference workbook structure

suppressPackageStartupMessages({
  library(openxlsx)
  library(readxl)
  library(lubridate)
  library(jsonlite)
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
.hs <- function() createStyle(fontSize = 11, fontColour = "#FFFFFF", fgFill = "#1F4E79",
                               halign = "center", textDecoration = "bold",
                               border = "TopBottomLeftRight", borderColour = "#1F4E79")
.ts <- function() createStyle(fontSize = 14, textDecoration = "bold", fontColour = "#1F4E79")
.ss <- function() createStyle(fontSize = 11, textDecoration = "bold", fontColour = "#505050")
.pos <- function() createStyle(fontColour = "#006100", fgFill = "#C6EFCE")
.neg <- function() createStyle(fontColour = "#9C0006", fgFill = "#FFC7CE")
.sep <- function() createStyle(fontSize = 16, textDecoration = "bold", fontColour = "#1F4E79",
                                fgFill = "#D9E2F3", halign = "center", valign = "center")

# Auto-detect sheet name from a CLA01 or X02 file
.detect_sheet <- function(path, candidates) {
  if (is.null(path)) return(NULL)
  sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
  for (c in candidates) {
    if (c %in% sheets) return(c)
  }
  if (length(sheets) > 0) sheets[1] else NULL
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
  # STEP 1: Build sheet spec for Python copier
  # ==========================================================================

  sheet_spec <- list()

  .add <- function(src, src_sheet, tgt_sheet) {
    if (!is.null(src)) {
      sheet_spec[[length(sheet_spec) + 1]] <<- list(
        source = src, source_sheet = src_sheet, target_sheet = tgt_sheet
      )
    }
  }

  # A01 sheets
  for (s in c("1", "2", "3", "5", "10", "11", "13", "15", "18", "19", "20", "21", "22")) {
    dest <- if (s == "1") "Sheet1" else s
    .add(file_a01, s, dest)
  }

  # RTISA sheets
  .add(file_rtisa, "1. Payrolled employees (UK)", "1. Payrolled employees (UK)")
  .add(file_rtisa, "23. Employees (Industry)", "23. Employees Industry")
  .add(file_rtisa, "6. Employee flows (UK)", "RTI. Employee flows (UK)")

  # X09
  .add(file_x09, "AWE Real_CPI", "AWE Real_CPI")

  # HR1
  for (s in c("1a", "1b", "2a", "2b")) .add(file_hr1, s, s)

  # CLA01 — auto-detect sheet name
  cla_sheet <- .detect_sheet(file_cla01, c("1", "People SA", "People"))
  if (!is.null(cla_sheet)) .add(file_cla01, cla_sheet, "1 UK")

  # X02 — auto-detect sheet name
  x02_sheet <- .detect_sheet(file_x02, c("LFS Labour market flows SA", "People SA", "1"))
  if (!is.null(x02_sheet)) .add(file_x02, x02_sheet, "LFS Labour market flows SA")

  # OECD files (could be CSV or xlsx)
  for (oecd_info in list(
    list(file = file_oecd_unemp, name = "Unemployment"),
    list(file = file_oecd_emp,   name = "Employment"),
    list(file = file_oecd_inact, name = "Inactivity")
  )) {
    if (!is.null(oecd_info$file)) {
      ext <- tolower(tools::file_ext(oecd_info$file))
      if (ext == "csv") {
        # CSV: Python will handle it (copy_csv_as_sheet)
        sheet_spec[[length(sheet_spec) + 1]] <- list(
          source = oecd_info$file, source_sheet = "", target_sheet = oecd_info$name
        )
      } else {
        oecd_sh <- .detect_sheet(oecd_info$file, c(oecd_info$name, "Sheet1", "Data"))
        if (!is.null(oecd_sh)) .add(oecd_info$file, oecd_sh, oecd_info$name)
      }
    }
  }

  # ==========================================================================
  # STEP 2: Call Python to copy source sheets with formatting
  # ==========================================================================

  py_script <- "utils/copy_sheets.py"
  tmp_assembled <- tempfile(fileext = ".xlsx")

  if (length(sheet_spec) > 0) {
    spec_json <- tempfile(fileext = ".json")
    writeLines(jsonlite::toJSON(sheet_spec, auto_unbox = TRUE), spec_json)

    if (verbose) message("[audit wb] Copying ", length(sheet_spec), " source sheets via Python...")

    result <- system2("python3", args = c(py_script, spec_json, tmp_assembled),
                      stdout = TRUE, stderr = TRUE)
    if (verbose) message("[audit wb] Python result: ", paste(result, collapse = "\n"))

    if (!file.exists(tmp_assembled)) {
      warning("Python sheet copier failed. Falling back to empty workbook.\n",
              paste(result, collapse = "\n"))
      wb <- createWorkbook()
    } else {
      wb <- loadWorkbook(tmp_assembled)
    }
  } else {
    wb <- createWorkbook()
  }

  # ==========================================================================
  # STEP 3: Detect reference period and compute Dashboard metrics
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
    list(name = "1 UK", cond = is.null(file_cla01), msg = "Upload CLA01 file to populate this sheet."),
    list(name = "LFS Labour market flows SA", cond = is.null(file_x02), msg = "Upload X02 file to populate this sheet.")
  )) {
    if (missing_info$cond && !missing_info$name %in% names(wb)) {
      addWorksheet(wb, missing_info$name)
      writeData(wb, missing_info$name, data.frame(Note = missing_info$msg))
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
