# manual_word_output.R
# Standalone Word output pipeline for Excel/manual mode.
# Uses ManualDB.docx template + calculations_from_excel.R — no database needed.
# Run from: project root directory

suppressPackageStartupMessages({
  library(officer)
  library(xml2)
  library(scales)
})

# ---------- formatters ----------

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
fmt_count_000s_change  <- function(x) fmt_int_signed(x / 1000)

# payroll/vacancies stored in 000s
fmt_exempt_current <- function(x) .format_int(x)
fmt_exempt_change  <- function(x) fmt_int_signed(x)

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
    yr  <- substr(x, 4, 7)
    month_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)
    if (mon %in% names(month_map)) {
      d <- as.Date(sprintf("%s-%02d-01", yr, month_map[[mon]]))
      return(format(d, "%B %Y"))
    }
  }
  paste0(toupper(substr(x, 1, 1)), substr(x, 2, nchar(x)))
}

# ---------- Word replacement helpers (direct XML for table cells) ----------

replace_all <- function(doc, key, val) {
  if (is.null(val) || length(val) == 0 || is.na(val)) val <- ""
  val <- as.character(val)

  body_xml   <- doc$doc_obj$get()
  ns         <- xml2::xml_ns(body_xml)
  text_nodes <- xml2::xml_find_all(body_xml, ".//w:t", ns = ns)
  for (node in text_nodes) {
    txt <- xml2::xml_text(node)
    if (grepl(key, txt, fixed = TRUE)) {
      new_txt <- gsub(key, val, txt, fixed = TRUE)
      xml2::xml_text(node) <- new_txt
      xml2::xml_attr(node, "xml:space") <- "preserve"
    }
  }

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
    if (isTRUE(invert)) { tmp <- p; p <- n; n <- tmp }
  }

  doc <- replace_all(doc, paste0(base, "_p"), p)
  doc <- replace_all(doc, paste0(base, "_n"), n)
  doc <- replace_all(doc, paste0(base, "_z"), z)
  doc
}

# ---------- safe value accessor ----------

sv <- function(name, default = NA_real_) {
  if (exists(name, inherits = TRUE)) get(name, inherits = TRUE) else default
}

# ---------- main ----------

generate_manual_word_output <- function(
    manual_month,
    file_a01  = NULL,
    file_x09  = NULL,
    file_rtisa = NULL,
    file_hr1  = NULL,
    template_path = "utils/ManualDB.docx",
    output_path   = "utils/ManualDBoutput.docx",
    verbose = TRUE
) {

  manual_month <- tolower(manual_month)

  # source helpers and run Excel calculations
  source("utils/helpers.R", local = FALSE)
  source("utils/calculations_from_excel.R", local = FALSE)
  run_calculations_from_excel(manual_month,
                              file_a01 = file_a01, file_hr1 = file_hr1,
                              file_x09 = file_x09, file_rtisa = file_rtisa)

  if (verbose) message("[manual] Calculations complete for ", manual_month)

  # generate summary and top-ten narrative lines
  source("sheets/summary.R", local = FALSE)
  source("sheets/top_ten_stats.R", local = FALSE)

  fallback_lines <- function() {
    stats <- list()
    for (i in 1:10) stats[[paste0("line", i)]] <- ""
    stats
  }

  summary <- tryCatch(generate_summary(), error = function(e) {
    if (verbose) warning("generate_summary() failed: ", e$message)
    fallback_lines()
  })
  top10 <- tryCatch(generate_top_ten(), error = function(e) {
    if (verbose) warning("generate_top_ten() failed: ", e$message)
    fallback_lines()
  })

  # open template
  doc <- read_docx(template_path)

  # ---- header / labels ----
  doc <- replace_all(doc, "Z1", manual_month_to_label(manual_month))
  doc <- replace_all(doc, "RENDER_DATE", format(Sys.Date(), "%d %B %Y"))
  doc <- replace_all(doc, "LFS_PERIOD_LABEL",    sv("lfs_period_label", ""))
  doc <- replace_all(doc, "LFS_QUARTER_LABEL",   sv("lfs_period_short_label", ""))
  doc <- replace_all(doc, "VACANCIES_QUARTER_LABEL", sv("vacancies_period_short_label", ""))

  # ---- summary + top ten lines ----
  for (i in 10:1) doc <- replace_all(doc, paste0("sl", i), summary[[paste0("line", i)]])
  for (i in 10:1) doc <- replace_all(doc, paste0("tt", i), top10[[paste0("line", i)]])

  # ---- Stats Dashboard: Current column ----
  doc <- replace_all(doc, "B1", fmt_count_000s_current(sv("emp16_cur")))
  doc <- replace_all(doc, "C1", .format_pct(sv("emp_rt_cur")))
  doc <- replace_all(doc, "D1", fmt_count_000s_current(sv("unemp16_cur")))
  doc <- replace_all(doc, "E1", .format_pct(sv("unemp_rt_cur")))
  doc <- replace_all(doc, "F1", fmt_count_000s_current(sv("inact_cur")))
  doc <- replace_all(doc, "G1", fmt_count_000s_current(sv("inact5064_cur")))
  doc <- replace_all(doc, "H1", .format_pct(sv("inact_rt_cur")))
  doc <- replace_all(doc, "I1", .format_pct(sv("inact5064_rt_cur")))
  doc <- replace_all(doc, "K1", fmt_exempt_current(sv("payroll_cur")))
  doc <- fill_conditional(doc, "J1", fmt_exempt_current(sv("vac_cur")), 0, neutral = TRUE)
  doc <- replace_all(doc, "L1", .format_pct(sv("latest_wages")))
  doc <- replace_all(doc, "M1", .format_pct(sv("latest_wages_cpi")))

  # ---- Stats Dashboard: Change on quarter ----
  doc <- fill_conditional(doc, "B2", fmt_count_000s_change(sv("emp16_dq")),  sv("emp16_dq"),  invert = FALSE)
  doc <- fill_conditional(doc, "C2", .format_pp(sv("emp_rt_dq")),           sv("emp_rt_dq"),  invert = FALSE)
  doc <- fill_conditional(doc, "D2", fmt_count_000s_change(sv("unemp16_dq")), sv("unemp16_dq"), invert = TRUE)
  doc <- fill_conditional(doc, "E2", .format_pp(sv("unemp_rt_dq")),         sv("unemp_rt_dq"), invert = TRUE)
  doc <- fill_conditional(doc, "F2", fmt_count_000s_change(sv("inact_dq")),  sv("inact_dq"),  invert = TRUE)
  doc <- fill_conditional(doc, "G2", fmt_count_000s_change(sv("inact5064_dq")), sv("inact5064_dq"), invert = TRUE)
  doc <- fill_conditional(doc, "H2", .format_pp(sv("inact_rt_dq")),         sv("inact_rt_dq"), invert = TRUE)
  doc <- fill_conditional(doc, "I2", .format_pp(sv("inact5064_rt_dq")),     sv("inact5064_rt_dq"), invert = TRUE)
  doc <- fill_conditional(doc, "K2", fmt_exempt_change(sv("payroll_dq")),    sv("payroll_dq"), invert = FALSE)
  doc <- fill_conditional(doc, "J2", fmt_exempt_change(sv("vac_dq")),        0, neutral = TRUE)
  doc <- fill_conditional(doc, "L2", .format_gbp_signed0(sv("wages_change_q")),     sv("wages_change_q"),     invert = FALSE)
  doc <- fill_conditional(doc, "M2", .format_gbp_signed0(sv("wages_cpi_change_q")), sv("wages_cpi_change_q"), invert = FALSE)

  # ---- Stats Dashboard: Change on year ----
  doc <- fill_conditional(doc, "B3", fmt_count_000s_change(sv("emp16_dy")),  sv("emp16_dy"),  invert = FALSE)
  doc <- fill_conditional(doc, "C3", .format_pp(sv("emp_rt_dy")),           sv("emp_rt_dy"),  invert = FALSE)
  doc <- fill_conditional(doc, "D3", fmt_count_000s_change(sv("unemp16_dy")), sv("unemp16_dy"), invert = TRUE)
  doc <- fill_conditional(doc, "E3", .format_pp(sv("unemp_rt_dy")),         sv("unemp_rt_dy"), invert = TRUE)
  doc <- fill_conditional(doc, "F3", fmt_count_000s_change(sv("inact_dy")),  sv("inact_dy"),  invert = TRUE)
  doc <- fill_conditional(doc, "G3", fmt_count_000s_change(sv("inact5064_dy")), sv("inact5064_dy"), invert = TRUE)
  doc <- fill_conditional(doc, "H3", .format_pp(sv("inact_rt_dy")),         sv("inact_rt_dy"), invert = TRUE)
  doc <- fill_conditional(doc, "I3", .format_pp(sv("inact5064_rt_dy")),     sv("inact5064_rt_dy"), invert = TRUE)
  doc <- fill_conditional(doc, "K3", fmt_exempt_change(sv("payroll_dy")),    sv("payroll_dy"), invert = FALSE)
  doc <- fill_conditional(doc, "J3", fmt_exempt_change(sv("vac_dy")),        0, neutral = TRUE)
  doc <- fill_conditional(doc, "L3", .format_gbp_signed0(sv("wages_change_y")),     sv("wages_change_y"),     invert = FALSE)
  doc <- fill_conditional(doc, "M3", .format_gbp_signed0(sv("wages_cpi_change_y")), sv("wages_cpi_change_y"), invert = FALSE)

  # ---- Stats Dashboard: Change since Covid-19 ----
  doc <- fill_conditional(doc, "B4", fmt_count_000s_change(sv("emp16_dc")),  sv("emp16_dc"),  invert = FALSE)
  doc <- fill_conditional(doc, "C4", .format_pp(sv("emp_rt_dc")),           sv("emp_rt_dc"),  invert = FALSE)
  doc <- fill_conditional(doc, "D4", fmt_count_000s_change(sv("unemp16_dc")), sv("unemp16_dc"), invert = TRUE)
  doc <- fill_conditional(doc, "E4", .format_pp(sv("unemp_rt_dc")),         sv("unemp_rt_dc"), invert = TRUE)
  doc <- fill_conditional(doc, "F4", fmt_count_000s_change(sv("inact_dc")),  sv("inact_dc"),  invert = TRUE)
  doc <- fill_conditional(doc, "G4", fmt_count_000s_change(sv("inact5064_dc")), sv("inact5064_dc"), invert = TRUE)
  doc <- fill_conditional(doc, "H4", .format_pp(sv("inact_rt_dc")),         sv("inact_rt_dc"), invert = TRUE)
  doc <- fill_conditional(doc, "I4", .format_pp(sv("inact5064_rt_dc")),     sv("inact5064_rt_dc"), invert = TRUE)
  doc <- fill_conditional(doc, "K4", fmt_exempt_change(sv("payroll_dc")),    sv("payroll_dc"), invert = FALSE)
  doc <- fill_conditional(doc, "J4", fmt_exempt_change(sv("vac_dc")),        0, neutral = TRUE)
  doc <- fill_conditional(doc, "L4", .format_gbp_signed0(sv("wages_change_covid")),     sv("wages_change_covid"),     invert = FALSE)
  doc <- fill_conditional(doc, "M4", .format_gbp_signed0(sv("wages_cpi_change_covid")), sv("wages_cpi_change_covid"), invert = FALSE)

  # ---- Stats Dashboard: Change since 2024 election ----
  doc <- fill_conditional(doc, "B5", fmt_count_000s_change(sv("emp16_de")),  sv("emp16_de"),  invert = FALSE)
  doc <- fill_conditional(doc, "C5", .format_pp(sv("emp_rt_de")),           sv("emp_rt_de"),  invert = FALSE)
  doc <- fill_conditional(doc, "D5", fmt_count_000s_change(sv("unemp16_de")), sv("unemp16_de"), invert = TRUE)
  doc <- fill_conditional(doc, "E5", .format_pp(sv("unemp_rt_de")),         sv("unemp_rt_de"), invert = TRUE)
  doc <- fill_conditional(doc, "F5", fmt_count_000s_change(sv("inact_de")),  sv("inact_de"),  invert = TRUE)
  doc <- fill_conditional(doc, "G5", fmt_count_000s_change(sv("inact5064_de")), sv("inact5064_de"), invert = TRUE)
  doc <- fill_conditional(doc, "H5", .format_pp(sv("inact_rt_de")),         sv("inact_rt_de"), invert = TRUE)
  doc <- fill_conditional(doc, "I5", .format_pp(sv("inact5064_rt_de")),     sv("inact5064_rt_de"), invert = TRUE)
  doc <- fill_conditional(doc, "K5", fmt_exempt_change(sv("payroll_de")),    sv("payroll_de"), invert = FALSE)
  doc <- fill_conditional(doc, "J5", fmt_exempt_change(sv("vac_de")),        0, neutral = TRUE)
  doc <- fill_conditional(doc, "L5", .format_gbp_signed0(sv("wages_change_election")),     sv("wages_change_election"),     invert = FALSE)
  doc <- fill_conditional(doc, "M5", .format_gbp_signed0(sv("wages_cpi_change_election")), sv("wages_cpi_change_election"), invert = FALSE)

  # ---- write output ----
  print(doc, target = output_path)
  if (verbose) message("[manual] Written to ", output_path)
  invisible(output_path)
}

# Example usage:
#   source("utils/manual_word_output.R")
#   generate_manual_word_output(
#     manual_month = "feb2026",
#     file_a01     = "path/to/a01feb2026.xlsx",
#     file_x09     = "path/to/x09feb2026.xlsx",
#     file_rtisa   = "path/to/rtisafeb2026.xlsx",
#     file_hr1     = "path/to/hr1feb2026.xlsx"
#   )
