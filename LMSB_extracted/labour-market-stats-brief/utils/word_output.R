# word_output.r
# fills lmsb_placeholders_v2 placeholders and conditional-colour tokens.
# run from: project root directory

suppressPackageStartupMessages({
  library(officer)
  library(scales)
})

# fallback formatters (uses helpers.r versions if present)
fmt_one_dec <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  if (x == 0) return(format(0, nsmall = 1, trim = TRUE))
  for (d in 1:4) {
    vr <- round(x, d)
    if (vr != 0) return(format(vr, nsmall = d, trim = TRUE))
  }
  format(round(x, 4), nsmall = 4, trim = TRUE)
}

.format_int <- function(x) {
  # IMPORTANT: current/dashboard values should not show a leading '+' sign.
  # helpers.R's format_int() is signed; prefer unsigned if available.
  if (exists("format_int_unsigned", inherits = TRUE)) return(get("format_int_unsigned", inherits = TRUE)(x))
  if (exists("format_int", inherits = TRUE)) return(get("format_int", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  scales::comma(round(x), accuracy = 1)
}

.format_pct <- function(x) {
  if (exists("format_pct", inherits = TRUE)) return(get("format_pct", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  paste0(fmt_one_dec(x), "%")
}

.format_pp <- function(x) {
  if (exists("format_pp", inherits = TRUE)) return(get("format_pp", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  sign <- if (x > 0) "+" else if (x < 0) "-" else ""
  paste0(sign, fmt_one_dec(abs(x)), "pp")
}

.format_gbp_signed0 <- function(x) {
  if (exists("format_gbp_signed0", inherits = TRUE)) return(get("format_gbp_signed0", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  sign <- if (x > 0) "+" else if (x < 0) "-" else ""
  paste0(sign, "\u00A3", scales::comma(round(abs(x)), accuracy = 1))
}

fmt_int_signed <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  s <- scales::comma(abs(round(x)), accuracy = 1)
  if (x > 0) paste0("+", s) else if (x < 0) paste0("-", s) else "0"
}

# counts stored as persons; displayed in 000s
fmt_count_000s_current <- function(x) .format_int(x / 1000)
fmt_count_000s_change <- function(x) fmt_int_signed(x / 1000)

# payroll/vacancies stored in 000s
fmt_exempt_current <- function(x) .format_int(x)
fmt_exempt_change <- function(x) fmt_int_signed(x)

# manual_month like "oct2025" or "2025-10" -> "october 2025"
manual_month_to_label <- function(x) {
  if (is.null(x) || length(x) == 0 || is.na(x)) return("")
  x <- tolower(as.character(x))
  if (grepl("^[0-9]{4}-[0-9]{2}$", x)) {
    parts <- strsplit(x, "-", fixed = TRUE)[[1]]
    d <- as.Date(sprintf("%s-%s-01", parts[1], parts[2]))
    return(format(d, "%B %Y"))
  }
  if (grepl("^[a-z]{3}[0-9]{4}$", x)) {
    mon <- substr(x, 1, 3)
    yr <- substr(x, 4, 7)
    month_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)
    if (mon %in% names(month_map)) {
      d <- as.Date(sprintf("%s-%02d-01", yr, month_map[[mon]]))
      return(format(d, "%B %Y"))
    }
  }
  paste0(toupper(substr(x, 1, 1)), substr(x, 2, nchar(x)))
}

# word replacement helpers
replace_all <- function(doc, key, val) {
  if (is.null(val) || length(val) == 0 || is.na(val)) val <- ""
  val <- as.character(val)
  doc <- body_replace_all_text(doc, key, val, fixed = TRUE)
  doc <- tryCatch(headers_replace_all_text(doc, key, val, fixed = TRUE), error = function(e) doc)
  doc <- tryCatch(footers_replace_all_text(doc, key, val, fixed = TRUE), error = function(e) doc)
  doc
}

fill_conditional <- function(doc, base, value_text, value_num, invert = FALSE, neutral = FALSE) {
  value_num <- suppressWarnings(as.numeric(value_num))
  if (is.na(value_num)) value_num <- 0
  
  p <- n <- z <- ""
  
  if (isTRUE(neutral)) {
    z <- value_text
  } else {
    if (value_num > 0) p <- value_text
    if (value_num < 0) n <- value_text
    if (value_num == 0) z <- value_text
    if (isTRUE(invert)) {
      tmp <- p; p <- n; n <- tmp
    }
  }
  
  doc <- replace_all(doc, paste0(base, "_p"), p)
  doc <- replace_all(doc, paste0(base, "_n"), n)
  doc <- replace_all(doc, paste0(base, "_z"), z)
  doc
}

# main

generate_word_output <- function(template_path = "utils/DB.docx",
                                 output_path = "utils/DBoutput.docx",
                                 calculations_path = "utils/calculations.R",
                                 config_path = "utils/config.R",
                                 summary_path = "sheets/summary.R",
                                 top_ten_path = "sheets/top_ten_stats.R",
                                 manual_month_override = NULL,
                                 vacancies_mode_override = NULL,
                                 payroll_mode_override = NULL,
                                 vac_payroll_mode_override = NULL,
                                 verbose = TRUE) {

  # source config first (auto-detects manual_month from database)
  source(config_path, local = FALSE)

  if (!is.null(manual_month_override)) manual_month <<- tolower(manual_month_override)
  
  # optional: allow vacancies/payroll to be aligned with reference quarter (or latest)
  # - pass both together via vac_payroll_mode_override (legacy)
  # - or pass independently via vacancies_mode_override / payroll_mode_override
  if (!is.null(vac_payroll_mode_override) && is.null(vacancies_mode_override) && is.null(payroll_mode_override)) {
    mode <- tolower(as.character(vac_payroll_mode_override))
    mode <- if (mode %in% c("latest", "aligned")) mode else "latest"
    vacancies_mode <<- mode
    payroll_mode <<- mode
  }
  
  if (!is.null(vacancies_mode_override)) {
    mode <- tolower(as.character(vacancies_mode_override))
    vacancies_mode <<- if (mode %in% c("latest", "aligned")) mode else "latest"
  }
  
  if (!is.null(payroll_mode_override)) {
    mode <- tolower(as.character(payroll_mode_override))
    payroll_mode <<- if (mode %in% c("latest", "aligned")) mode else "latest"
  }
  
  if (verbose && exists("manual_month", inherits = TRUE)) message("[word_output] manual_month = ", manual_month)
  
  # source calculations (this sources helpers.r and all sheets)
  source(calculations_path, local = FALSE)
  
  # source summary and top ten
  source(summary_path, local = FALSE)
  source(top_ten_path, local = FALSE)
  fallback_lines <- function() {
    stats <- list()
    for (i in 1:10) stats[[paste0("line", i)]] <- "(Data unavailable)"
    stats
  }
  summary <- tryCatch(generate_summary(), error = function(e) {
    warning("generate_summary() failed: ", e$message)
    fallback_lines()
  })
  top10 <- tryCatch(generate_top_ten(), error = function(e) {
    warning("generate_top_ten() failed: ", e$message)
    fallback_lines()
  })
  
  doc <- read_docx(template_path)
  
  title_label <- if (exists("manual_month", inherits = TRUE)) manual_month_to_label(manual_month) else ""
  doc <- replace_all(doc, "Z1", title_label)
  if (exists("lfs_period_label", inherits = TRUE)) doc <- replace_all(doc, "LFS_PERIOD_LABEL", lfs_period_label)
  if (exists("lfs_period_short_label", inherits = TRUE)) doc <- replace_all(doc, "LFS_QUARTER_LABEL", lfs_period_short_label)
  if (exists("vacancies_period_short_label", inherits = TRUE)) doc <- replace_all(doc, "VACANCIES_QUARTER_LABEL", vacancies_period_short_label)
  
  for (i in 10:1) doc <- replace_all(doc, paste0("sl", i), summary[[paste0("line", i)]])
  for (i in 10:1) doc <- replace_all(doc, paste0("tt", i), top10[[paste0("line", i)]])
  
  doc <- replace_all(doc, "RENDER_DATE", format(Sys.Date(), "%d %B %Y"))
  
  # current (no conditional except vacancies)
  doc <- replace_all(doc, "B1", fmt_count_000s_current(emp16_cur))
  doc <- replace_all(doc, "C1", .format_pct(emp_rt_cur))
  
  doc <- replace_all(doc, "D1", fmt_count_000s_current(unemp16_cur))
  doc <- replace_all(doc, "E1", .format_pct(unemp_rt_cur))
  
  doc <- replace_all(doc, "F1", fmt_count_000s_current(inact_cur))
  doc <- replace_all(doc, "G1", fmt_count_000s_current(inact5064_cur))
  doc <- replace_all(doc, "H1", .format_pct(inact_rt_cur))
  doc <- replace_all(doc, "I1", .format_pct(inact5064_rt_cur))
  
  doc <- replace_all(doc, "K1", fmt_exempt_current(payroll_cur))
  
  # vacancies current is conditional tokens (neutral)
  doc <- fill_conditional(doc, "J1", fmt_exempt_current(vac_cur), 0, neutral = TRUE)
  
  doc <- replace_all(doc, "L1", .format_pct(latest_wages))
  doc <- replace_all(doc, "M1", .format_pct(latest_wages_cpi))
  
  # dq/dy/dc conditional
  doc <- fill_conditional(doc, "B2", fmt_count_000s_change(emp16_dq), emp16_dq, invert = FALSE)
  doc <- fill_conditional(doc, "B3", fmt_count_000s_change(emp16_dy), emp16_dy, invert = FALSE)
  doc <- fill_conditional(doc, "B4", fmt_count_000s_change(emp16_dc), emp16_dc, invert = FALSE)
  
  doc <- fill_conditional(doc, "C2", .format_pp(emp_rt_dq), emp_rt_dq, invert = FALSE)
  doc <- fill_conditional(doc, "C3", .format_pp(emp_rt_dy), emp_rt_dy, invert = FALSE)
  doc <- fill_conditional(doc, "C4", .format_pp(emp_rt_dc), emp_rt_dc, invert = FALSE)
  
  doc <- fill_conditional(doc, "D2", fmt_count_000s_change(unemp16_dq), unemp16_dq, invert = TRUE)
  doc <- fill_conditional(doc, "D3", fmt_count_000s_change(unemp16_dy), unemp16_dy, invert = TRUE)
  doc <- fill_conditional(doc, "D4", fmt_count_000s_change(unemp16_dc), unemp16_dc, invert = TRUE)
  
  doc <- fill_conditional(doc, "E2", .format_pp(unemp_rt_dq), unemp_rt_dq, invert = TRUE)
  doc <- fill_conditional(doc, "E3", .format_pp(unemp_rt_dy), unemp_rt_dy, invert = TRUE)
  doc <- fill_conditional(doc, "E4", .format_pp(unemp_rt_dc), unemp_rt_dc, invert = TRUE)
  
  doc <- fill_conditional(doc, "F2", fmt_count_000s_change(inact_dq), inact_dq, invert = TRUE)
  doc <- fill_conditional(doc, "F3", fmt_count_000s_change(inact_dy), inact_dy, invert = TRUE)
  doc <- fill_conditional(doc, "F4", fmt_count_000s_change(inact_dc), inact_dc, invert = TRUE)
  
  doc <- fill_conditional(doc, "G2", fmt_count_000s_change(inact5064_dq), inact5064_dq, invert = TRUE)
  doc <- fill_conditional(doc, "G3", fmt_count_000s_change(inact5064_dy), inact5064_dy, invert = TRUE)
  doc <- fill_conditional(doc, "G4", fmt_count_000s_change(inact5064_dc), inact5064_dc, invert = TRUE)
  
  doc <- fill_conditional(doc, "H2", .format_pp(inact_rt_dq), inact_rt_dq, invert = TRUE)
  doc <- fill_conditional(doc, "H3", .format_pp(inact_rt_dy), inact_rt_dy, invert = TRUE)
  doc <- fill_conditional(doc, "H4", .format_pp(inact_rt_dc), inact_rt_dc, invert = TRUE)
  
  doc <- fill_conditional(doc, "I2", .format_pp(inact5064_rt_dq), inact5064_rt_dq, invert = TRUE)
  doc <- fill_conditional(doc, "I3", .format_pp(inact5064_rt_dy), inact5064_rt_dy, invert = TRUE)
  doc <- fill_conditional(doc, "I4", .format_pp(inact5064_rt_dc), inact5064_rt_dc, invert = TRUE)
  
  doc <- fill_conditional(doc, "K2", fmt_exempt_change(payroll_dq), payroll_dq, invert = FALSE)
  doc <- fill_conditional(doc, "K3", fmt_exempt_change(payroll_dy), payroll_dy, invert = FALSE)
  doc <- fill_conditional(doc, "K4", fmt_exempt_change(payroll_dc), payroll_dc, invert = FALSE)
  
  # vacancies neutral across changes
  doc <- fill_conditional(doc, "J2", fmt_exempt_change(vac_dq), 0, neutral = TRUE)
  doc <- fill_conditional(doc, "J3", fmt_exempt_change(vac_dy), 0, neutral = TRUE)
  doc <- fill_conditional(doc, "J4", fmt_exempt_change(vac_dc), 0, neutral = TRUE)
  
  doc <- fill_conditional(doc, "L2", .format_gbp_signed0(wages_change_q), wages_change_q, invert = FALSE)
  doc <- fill_conditional(doc, "L3", .format_gbp_signed0(wages_change_y), wages_change_y, invert = FALSE)
  doc <- fill_conditional(doc, "L4", .format_gbp_signed0(wages_change_covid), wages_change_covid, invert = FALSE)
  
  doc <- fill_conditional(doc, "M2", .format_gbp_signed0(wages_cpi_change_q), wages_cpi_change_q, invert = FALSE)
  doc <- fill_conditional(doc, "M3", .format_gbp_signed0(wages_cpi_change_y), wages_cpi_change_y, invert = FALSE)
  doc <- fill_conditional(doc, "M4", .format_gbp_signed0(wages_cpi_change_covid), wages_cpi_change_covid, invert = FALSE)
  
  # election column
  if (exists("emp16_de", inherits = TRUE)) {
    doc <- fill_conditional(doc, "B5", fmt_count_000s_change(emp16_de), emp16_de, invert = FALSE)
    doc <- fill_conditional(doc, "C5", .format_pp(emp_rt_de), emp_rt_de, invert = FALSE)
    
    doc <- fill_conditional(doc, "D5", fmt_count_000s_change(unemp16_de), unemp16_de, invert = TRUE)
    doc <- fill_conditional(doc, "E5", .format_pp(unemp_rt_de), unemp_rt_de, invert = TRUE)
    
    doc <- fill_conditional(doc, "F5", fmt_count_000s_change(inact_de), inact_de, invert = TRUE)
    doc <- fill_conditional(doc, "G5", fmt_count_000s_change(inact5064_de), inact5064_de, invert = TRUE)
    doc <- fill_conditional(doc, "H5", .format_pp(inact_rt_de), inact_rt_de, invert = TRUE)
    doc <- fill_conditional(doc, "I5", .format_pp(inact5064_rt_de), inact5064_rt_de, invert = TRUE)
    
    doc <- fill_conditional(doc, "K5", fmt_exempt_change(payroll_de), payroll_de, invert = FALSE)
    doc <- fill_conditional(doc, "J5", fmt_exempt_change(vac_de), 0, neutral = TRUE)
    
    if (exists("wages_change_election", inherits = TRUE)) {
      doc <- fill_conditional(doc, "L5", .format_gbp_signed0(wages_change_election), wages_change_election, invert = FALSE)
    }
    if (exists("wages_cpi_change_election", inherits = TRUE)) {
      doc <- fill_conditional(doc, "M5", .format_gbp_signed0(wages_cpi_change_election), wages_cpi_change_election, invert = FALSE)
    }
  }
  

  # extra sections: workforce jobs / unemployment by age / payroll by age
  # (if placeholders are not present in the template, these replacements do nothing.)

  # workforce jobs
  wfj_period <- if (exists("workforce_jobs", inherits = TRUE)) workforce_jobs$period else ""
  wfj_tbl <- if (exists("workforce_jobs", inherits = TRUE)) workforce_jobs$data else NULL
  wfj_top_ind <- ""
  wfj_top_val <- ""
  if (!is.null(wfj_tbl) && nrow(wfj_tbl) > 0) {
    wfj_top_ind <- as.character(wfj_tbl$industry[1])
    wfj_top_val <- .format_int(wfj_tbl$value[1])
  }

  doc <- replace_all(doc, "WFJ_PERIOD", wfj_period)
  doc <- replace_all(doc, "WORKFORCE_JOBS_PERIOD", wfj_period)
  doc <- replace_all(doc, "WFJ_TOP_INDUSTRY", wfj_top_ind)
  doc <- replace_all(doc, "WORKFORCE_JOBS_TOP_INDUSTRY", wfj_top_ind)
  doc <- replace_all(doc, "WFJ_TOP_VALUE", wfj_top_val)
  doc <- replace_all(doc, "WORKFORCE_JOBS_TOP_VALUE", wfj_top_val)

  # optional top-5 lines (industry: value)
  if (!is.null(wfj_tbl) && nrow(wfj_tbl) > 0) {
    for (i in 1:min(5, nrow(wfj_tbl))) {
      line <- paste0(as.character(wfj_tbl$industry[i]), ": ", .format_int(wfj_tbl$value[i]))
      doc <- replace_all(doc, paste0("WFJ_LINE", i), line)
      doc <- replace_all(doc, paste0("WORKFORCE_JOBS_LINE", i), line)
    }
  }

  # unemployment by age
  uage_period <- if (exists("unemployment_by_age", inherits = TRUE)) unemployment_by_age$period else ""
  uage_level <- if (exists("unemployment_by_age", inherits = TRUE)) unemployment_by_age$level else NULL
  uage_rate <- if (exists("unemployment_by_age", inherits = TRUE)) unemployment_by_age$rate else NULL

  doc <- replace_all(doc, "UNEMP_AGE_PERIOD", uage_period)
  doc <- replace_all(doc, "UNEMPLOYMENT_BY_AGE_PERIOD", uage_period)

  if (!is.null(uage_level) && nrow(uage_level) > 0) {
    doc <- replace_all(doc, "UNEMP_AGE_TOP_LEVEL_AGE", as.character(uage_level$age_group[1]))
    doc <- replace_all(doc, "UNEMP_AGE_TOP_LEVEL_VALUE", .format_int(uage_level$value[1]))
    for (i in 1:min(5, nrow(uage_level))) {
      line <- paste0(as.character(uage_level$age_group[i]), ": ", .format_int(uage_level$value[i]))
      doc <- replace_all(doc, paste0("UNEMP_AGE_LEVEL_LINE", i), line)
    }
  }

  if (!is.null(uage_rate) && nrow(uage_rate) > 0) {
    doc <- replace_all(doc, "UNEMP_AGE_TOP_RATE_AGE", as.character(uage_rate$age_group[1]))
    doc <- replace_all(doc, "UNEMP_AGE_TOP_RATE_VALUE", .format_pct(uage_rate$value[1]))
    for (i in 1:min(5, nrow(uage_rate))) {
      line <- paste0(as.character(uage_rate$age_group[i]), ": ", .format_pct(uage_rate$value[i]))
      doc <- replace_all(doc, paste0("UNEMP_AGE_RATE_LINE", i), line)
    }
  }

  # payroll by age
  pba_period <- if (exists("payroll_by_age", inherits = TRUE)) payroll_by_age$period else ""
  pba_tbl <- if (exists("payroll_by_age", inherits = TRUE)) payroll_by_age$data else NULL

  doc <- replace_all(doc, "PAYROLL_AGE_PERIOD", pba_period)
  doc <- replace_all(doc, "PAYROLLED_EMPLOYEES_BY_AGE_PERIOD", pba_period)

  if (!is.null(pba_tbl) && nrow(pba_tbl) > 0) {
    doc <- replace_all(doc, "PAYROLL_AGE_TOP_AGE", as.character(pba_tbl$age_group[1]))
    doc <- replace_all(doc, "PAYROLL_AGE_TOP_VALUE", .format_int(pba_tbl$value[1]))
    for (i in 1:min(5, nrow(pba_tbl))) {
      line <- paste0(as.character(pba_tbl$age_group[i]), ": ", .format_int(pba_tbl$value[i]))
      doc <- replace_all(doc, paste0("PAYROLL_AGE_LINE", i), line)
    }
  }

  print(doc, target = output_path)
  invisible(output_path)
}

# run - from project root directory
generate_word_output()
