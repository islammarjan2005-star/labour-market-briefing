# manual_word_output.R
# Standalone Word output pipeline for Excel/manual mode.
# Uses ManualDB.docx template + calculations_from_excel.R — no database needed.
# Run from: project root directory

suppressPackageStartupMessages({
  library(officer)
  library(xml2)
  library(scales)
  library(flextable)
  library(readxl)
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

  p <- n <- z <- ""

  if (is.na(value_num)) {
    z <- "\u2014"
  } else if (isTRUE(neutral)) {
    z <- value_text
  } else {
    if (value_num > 0) p <- value_text
    if (value_num < 0) n <- value_text
    if (value_num == 0) z <- value_text
    if (isTRUE(invert)) { tmp <- p; p <- n; n <- tmp }
  }

  doc <- replace_all(doc, paste0(base, "p"), p)
  doc <- replace_all(doc, paste0(base, "n"), n)
  doc <- replace_all(doc, paste0(base, "z"), z)
  doc
}

# ---------- safe value accessor ----------

sv <- function(name, default = NA_real_) {
  if (exists(name, inherits = TRUE)) get(name, inherits = TRUE) else default
}

# ---------- OECD data extraction ----------

# Target countries for the international comparison table (order matters)
.oecd_countries <- c(
  "United Kingdom", "United States", "France", "Germany",
  "Italy", "Spain", "Canada", "Japan", "Euro area"
)

# Alternative labels that may appear in OECD files
.oecd_country_aliases <- list(
  "United Kingdom" = c("United Kingdom", "GBR"),
  "United States"  = c("United States", "USA"),
  "France"         = c("France", "FRA"),
  "Germany"        = c("Germany", "DEU"),
  "Italy"          = c("Italy", "ITA"),
  "Spain"          = c("Spain", "ESP"),
  "Canada"         = c("Canada", "CAN"),
  "Japan"          = c("Japan", "JPN"),
  "Euro area"      = c("Euro area", "Euro area (20 countries)",
                        "Euro area (19 countries)", "EA20", "EA19", "EA")
)

# Read an OECD Excel/CSV file and return a data.frame with columns:
#   country, period, value
# Extracts the latest available value per country.
.read_oecd_latest <- function(path) {
  if (is.null(path) || !file.exists(path)) return(NULL)

  ext <- tolower(tools::file_ext(path))

  # --- Excel with "Table" sheet (wide format: countries as rows, periods as cols) ---
  if (ext %in% c("xlsx", "xls")) {
    sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
    tbl_sheet <- if ("Table" %in% sheets) "Table" else NULL

    if (!is.null(tbl_sheet)) {
      raw <- suppressMessages(readxl::read_excel(path, sheet = tbl_sheet, col_names = FALSE))
      # The Table sheet typically has:
      #   - First few rows: metadata (title, measure, etc.)
      #   - A row that starts with "Country" or country names, followed by period columns
      #   - Data rows: country name + numeric values per period

      # Find the header row (contains period-like strings such as "Q1-2025", "2024-Q4", etc.)
      header_row <- NULL
      for (ri in 1:min(15, nrow(raw))) {
        row_text <- as.character(unlist(raw[ri, ]))
        # Periods typically look like "Q1-2025", "2024-Q4", "Jan 2025", "2024-01", etc.
        period_hits <- sum(grepl("(Q[1-4]|20[0-9]{2}|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", row_text, ignore.case = TRUE))
        if (period_hits >= 3) {
          header_row <- ri
          break
        }
      }
      if (is.null(header_row)) return(NULL)

      headers <- as.character(unlist(raw[header_row, ]))
      data_rows <- raw[(header_row + 1):nrow(raw), , drop = FALSE]

      results <- data.frame(country = character(), period = character(),
                            value = numeric(), stringsAsFactors = FALSE)

      for (target in names(.oecd_country_aliases)) {
        aliases <- .oecd_country_aliases[[target]]
        for (ri in seq_len(nrow(data_rows))) {
          row_country <- trimws(as.character(data_rows[[1]][ri]))
          if (tolower(row_country) %in% tolower(aliases)) {
            # Find latest non-NA numeric value (rightmost column)
            for (ci in ncol(data_rows):2) {
              val <- suppressWarnings(as.numeric(as.character(data_rows[[ci]][ri])))
              if (!is.na(val)) {
                period_label <- trimws(headers[ci])
                results <- rbind(results, data.frame(
                  country = target, period = period_label, value = val,
                  stringsAsFactors = FALSE
                ))
                break
              }
            }
            break
          }
        }
      }
      if (nrow(results) > 0) return(results)
    }

    # Fallback: try SDMX-like sheet
    if (length(sheets) > 0) {
      raw <- suppressMessages(readxl::read_excel(path, sheet = sheets[1], n_max = 5000))
      return(.parse_oecd_sdmx(raw))
    }
  }

  # --- CSV (SDMX long format) ---
  if (ext == "csv") {
    raw <- tryCatch(read.csv(path, stringsAsFactors = FALSE), error = function(e) data.frame())
    if (nrow(raw) > 0) return(.parse_oecd_sdmx(raw))
  }

  NULL
}

# Parse SDMX long-format data (REF_AREA, TIME_PERIOD, OBS_VALUE columns)
.parse_oecd_sdmx <- function(df) {
  ref_col  <- intersect(c("Reference.area", "REF_AREA"), names(df))[1]
  time_col <- intersect(c("Time.period", "TIME_PERIOD"), names(df))[1]
  val_col  <- intersect(c("OBS_VALUE", "Observation.value"), names(df))[1]

  if (is.na(ref_col) || is.na(time_col) || is.na(val_col)) return(NULL)

  results <- data.frame(country = character(), period = character(),
                        value = numeric(), stringsAsFactors = FALSE)

  for (target in names(.oecd_country_aliases)) {
    aliases <- tolower(.oecd_country_aliases[[target]])
    rows <- df[tolower(trimws(df[[ref_col]])) %in% aliases, , drop = FALSE]
    if (nrow(rows) == 0) next
    # Sort by period descending, take latest
    rows <- rows[order(rows[[time_col]], decreasing = TRUE), ]
    val <- suppressWarnings(as.numeric(rows[[val_col]][1]))
    if (!is.na(val)) {
      results <- rbind(results, data.frame(
        country = target, period = trimws(rows[[time_col]][1]), value = val,
        stringsAsFactors = FALSE
      ))
    }
  }
  if (nrow(results) > 0) return(results) else NULL
}

# Build the OECD comparison table as a flextable ready for Word insertion.
# Returns NULL if no data available.
.build_oecd_flextable <- function(unemp_data, emp_data, inact_data) {
  if (is.null(unemp_data) && is.null(emp_data) && is.null(inact_data)) return(NULL)

  # Merge all three metrics into one table
  tbl <- data.frame(country = .oecd_countries, stringsAsFactors = FALSE)

  .merge_metric <- function(metric_data, tbl, val_name, period_name) {
    if (is.null(metric_data)) {
      tbl[[val_name]] <- ""
      tbl[[period_name]] <- ""
      return(tbl)
    }
    idx <- match(tbl$country, metric_data$country)
    tbl[[val_name]] <- ifelse(is.na(idx), "",
      sapply(metric_data$value[idx], function(v) if (is.na(v)) "" else fmt_one_dec(v)))
    tbl[[period_name]] <- ifelse(is.na(idx), "", metric_data$period[idx])
    tbl
  }

  # We need: Country | Time Period | Unemployment Rate | Employment Rate | Inactivity Rate
  # The time period may differ per metric, so pick the most relevant one
  # In the published briefing, there's a single "Time Period" column
  # We'll use unemployment period as primary, fall back to employment, then inactivity
  tbl <- .merge_metric(unemp_data, tbl, "unemp_val", "unemp_period")
  tbl <- .merge_metric(emp_data, tbl, "emp_val", "emp_period")
  tbl <- .merge_metric(inact_data, tbl, "inact_val", "inact_period")

  # Build the display period (take first non-empty from unemp > emp > inact)
  tbl$period <- ifelse(nzchar(tbl$unemp_period), tbl$unemp_period,
                  ifelse(nzchar(tbl$emp_period), tbl$emp_period, tbl$inact_period))

  # Add percentage signs to values
  tbl$unemp_display <- ifelse(nzchar(tbl$unemp_val), paste0(tbl$unemp_val, "%"), "")
  tbl$emp_display   <- ifelse(nzchar(tbl$emp_val), paste0(tbl$emp_val, "%"), "")
  tbl$inact_display <- ifelse(nzchar(tbl$inact_val), paste0(tbl$inact_val, "%"), "")

  # Add asterisk for UK
  tbl$country[tbl$country == "United Kingdom"] <- "United Kingdom*"

  # Final display table
  display_df <- data.frame(
    `Country / Region` = tbl$country,
    `Time Period`      = tbl$period,
    `Unemployment Rate\n(15+, %)` = tbl$unemp_display,
    `Employment Rate\n(15-64, %)`  = tbl$emp_display,
    `Inactivity Rate\n(15-64, %)`  = tbl$inact_display,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )

  ft <- flextable(display_df)
  ft <- set_header_labels(ft,
    `Country / Region` = "Country / Region",
    `Time Period` = "Time Period",
    `Unemployment Rate\n(15+, %)` = "Unemployment Rate\n(15+, %)",
    `Employment Rate\n(15-64, %)` = "Employment Rate\n(15-64, %)",
    `Inactivity Rate\n(15-64, %)` = "Inactivity Rate\n(15-64, %)"
  )
  ft <- theme_box(ft)
  ft <- fontsize(ft, size = 9, part = "all")
  ft <- font(ft, fontname = "Arial", part = "all")
  ft <- bold(ft, part = "header")
  ft <- bg(ft, bg = "#2F5496", part = "header")
  ft <- color(ft, color = "white", part = "header")
  ft <- align(ft, j = 3:5, align = "center", part = "all")
  ft <- align(ft, j = 1:2, align = "left", part = "all")
  ft <- width(ft, j = 1, width = 1.5)
  ft <- width(ft, j = 2, width = 1.2)
  ft <- width(ft, j = 3:5, width = 1.3)
  # Bold UK row
  uk_row <- which(grepl("United Kingdom", display_df$`Country / Region`))
  if (length(uk_row) > 0) ft <- bold(ft, i = uk_row, part = "body")

  ft
}

# ---------- main ----------

generate_manual_word_output <- function(
    manual_month = NULL,
    file_a01  = NULL,
    file_x09  = NULL,
    file_rtisa = NULL,
    file_hr1  = NULL,
    file_oecd_unemp = NULL,
    file_oecd_emp   = NULL,
    file_oecd_inact = NULL,
    template_path = "utils/ManualDB.docx",
    output_path   = "utils/ManualDBoutput.docx",
    verbose = TRUE
) {

  if (!is.null(manual_month)) manual_month <- tolower(manual_month)

  # source helpers and run Excel calculations (auto-detects month from A01 if NULL)
  source("utils/helpers.R", local = FALSE)
  source("utils/calculations_from_excel.R", local = FALSE)
  manual_month <- run_calculations_from_excel(manual_month,
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
  # placeholder key: qvz prefix + lowercase descriptor, no underscores
  doc <- replace_all(doc, "qvzmonthlabel", manual_month_to_label(manual_month))
  doc <- replace_all(doc, "qvzrenderdate", format(Sys.Date(), "%d %B %Y"))
  doc <- replace_all(doc, "qvzlfsperiod",  sv("lfs_period_label", ""))
  doc <- replace_all(doc, "qvzlfsquarter", sv("lfs_period_short_label", ""))
  doc <- replace_all(doc, "qvzvacquarter", sv("vacancies_period_short_label", ""))

  # ---- summary + top ten lines ----
  for (i in 1:10) doc <- replace_all(doc, sprintf("qvzsl%02d", i), summary[[paste0("line", i)]])
  for (i in 1:10) doc <- replace_all(doc, sprintf("qvztt%02d", i), top10[[paste0("line", i)]])

  # ---- Stats Dashboard: Current column ----
  # stat codes: emp ert une urt ina irt ife ifr pay vac wno wcp
  doc <- replace_all(doc, "qvzempcur",  fmt_count_000s_current(sv("emp16_cur")))
  doc <- replace_all(doc, "qvzertcur",  .format_pct(sv("emp_rt_cur")))
  doc <- replace_all(doc, "qvzunecur",  fmt_count_000s_current(sv("unemp16_cur")))
  doc <- replace_all(doc, "qvzurtcur",  .format_pct(sv("unemp_rt_cur")))
  doc <- replace_all(doc, "qvzinacur",  fmt_count_000s_current(sv("inact_cur")))
  doc <- replace_all(doc, "qvzifecur",  fmt_count_000s_current(sv("inact5064_cur")))
  doc <- replace_all(doc, "qvzirtcur",  .format_pct(sv("inact_rt_cur")))
  doc <- replace_all(doc, "qvzifrcur",  .format_pct(sv("inact5064_rt_cur")))
  doc <- replace_all(doc, "qvzpaycur",  fmt_exempt_current(sv("payroll_cur")))
  doc <- fill_conditional(doc, "qvzvaccur", fmt_exempt_current(sv("vac_cur")), 0, neutral = TRUE)
  doc <- replace_all(doc, "qvzwnocur",  .format_pct(sv("latest_wages")))
  doc <- replace_all(doc, "qvzwcpcur",  .format_pct(sv("latest_wages_cpi")))

  # ---- Stats Dashboard: Change on quarter ----
  doc <- fill_conditional(doc, "qvzempdq",  fmt_count_000s_change(sv("emp16_dq")),        sv("emp16_dq"))
  doc <- fill_conditional(doc, "qvzertdq",  .format_pp(sv("emp_rt_dq")),                  sv("emp_rt_dq"))
  doc <- fill_conditional(doc, "qvzunedq",  fmt_count_000s_change(sv("unemp16_dq")),      sv("unemp16_dq"),      invert = TRUE)
  doc <- fill_conditional(doc, "qvzurtdq",  .format_pp(sv("unemp_rt_dq")),                sv("unemp_rt_dq"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzinadq",  fmt_count_000s_change(sv("inact_dq")),         sv("inact_dq"),        invert = TRUE)
  doc <- fill_conditional(doc, "qvzifedq",  fmt_count_000s_change(sv("inact5064_dq")),     sv("inact5064_dq"),    invert = TRUE)
  doc <- fill_conditional(doc, "qvzirtdq",  .format_pp(sv("inact_rt_dq")),                 sv("inact_rt_dq"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzifrdq",  .format_pp(sv("inact5064_rt_dq")),             sv("inact5064_rt_dq"), invert = TRUE)
  doc <- fill_conditional(doc, "qvzpaydq",  fmt_exempt_change(sv("payroll_dq")),            sv("payroll_dq"))
  doc <- fill_conditional(doc, "qvzvacdq",  fmt_exempt_change(sv("vac_dq")),               0, neutral = TRUE)
  doc <- fill_conditional(doc, "qvzwnodq",  .format_gbp_signed0(sv("wages_change_q")),     sv("wages_change_q"))
  doc <- fill_conditional(doc, "qvzwcpdq",  .format_gbp_signed0(sv("wages_cpi_change_q")), sv("wages_cpi_change_q"))

  # ---- Stats Dashboard: Change on year ----
  doc <- fill_conditional(doc, "qvzempdy",  fmt_count_000s_change(sv("emp16_dy")),        sv("emp16_dy"))
  doc <- fill_conditional(doc, "qvzertdy",  .format_pp(sv("emp_rt_dy")),                  sv("emp_rt_dy"))
  doc <- fill_conditional(doc, "qvzunedy",  fmt_count_000s_change(sv("unemp16_dy")),      sv("unemp16_dy"),      invert = TRUE)
  doc <- fill_conditional(doc, "qvzurtdy",  .format_pp(sv("unemp_rt_dy")),                sv("unemp_rt_dy"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzinady",  fmt_count_000s_change(sv("inact_dy")),         sv("inact_dy"),        invert = TRUE)
  doc <- fill_conditional(doc, "qvzifedy",  fmt_count_000s_change(sv("inact5064_dy")),     sv("inact5064_dy"),    invert = TRUE)
  doc <- fill_conditional(doc, "qvzirtdy",  .format_pp(sv("inact_rt_dy")),                 sv("inact_rt_dy"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzifrdy",  .format_pp(sv("inact5064_rt_dy")),             sv("inact5064_rt_dy"), invert = TRUE)
  doc <- fill_conditional(doc, "qvzpaydy",  fmt_exempt_change(sv("payroll_dy")),            sv("payroll_dy"))
  doc <- fill_conditional(doc, "qvzvacdy",  fmt_exempt_change(sv("vac_dy")),               0, neutral = TRUE)
  doc <- fill_conditional(doc, "qvzwnody",  .format_gbp_signed0(sv("wages_change_y")),     sv("wages_change_y"))
  doc <- fill_conditional(doc, "qvzwcpdy",  .format_gbp_signed0(sv("wages_cpi_change_y")), sv("wages_cpi_change_y"))

  # ---- Stats Dashboard: Change since Covid-19 ----
  doc <- fill_conditional(doc, "qvzempdc",  fmt_count_000s_change(sv("emp16_dc")),        sv("emp16_dc"))
  doc <- fill_conditional(doc, "qvzertdc",  .format_pp(sv("emp_rt_dc")),                  sv("emp_rt_dc"))
  doc <- fill_conditional(doc, "qvzunedc",  fmt_count_000s_change(sv("unemp16_dc")),      sv("unemp16_dc"),      invert = TRUE)
  doc <- fill_conditional(doc, "qvzurtdc",  .format_pp(sv("unemp_rt_dc")),                sv("unemp_rt_dc"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzinadc",  fmt_count_000s_change(sv("inact_dc")),         sv("inact_dc"),        invert = TRUE)
  doc <- fill_conditional(doc, "qvzifedc",  fmt_count_000s_change(sv("inact5064_dc")),     sv("inact5064_dc"),    invert = TRUE)
  doc <- fill_conditional(doc, "qvzirtdc",  .format_pp(sv("inact_rt_dc")),                 sv("inact_rt_dc"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzifrdc",  .format_pp(sv("inact5064_rt_dc")),             sv("inact5064_rt_dc"), invert = TRUE)
  doc <- fill_conditional(doc, "qvzpaydc",  fmt_exempt_change(sv("payroll_dc")),            sv("payroll_dc"))
  doc <- fill_conditional(doc, "qvzvacdc",  fmt_exempt_change(sv("vac_dc")),               0, neutral = TRUE)
  doc <- fill_conditional(doc, "qvzwnodc",  .format_gbp_signed0(sv("wages_change_covid")),     sv("wages_change_covid"))
  doc <- fill_conditional(doc, "qvzwcpdc",  .format_gbp_signed0(sv("wages_cpi_change_covid")), sv("wages_cpi_change_covid"))

  # ---- Stats Dashboard: Change since 2024 election ----
  doc <- fill_conditional(doc, "qvzempde",  fmt_count_000s_change(sv("emp16_de")),        sv("emp16_de"))
  doc <- fill_conditional(doc, "qvzertde",  .format_pp(sv("emp_rt_de")),                  sv("emp_rt_de"))
  doc <- fill_conditional(doc, "qvzunede",  fmt_count_000s_change(sv("unemp16_de")),      sv("unemp16_de"),      invert = TRUE)
  doc <- fill_conditional(doc, "qvzurtde",  .format_pp(sv("unemp_rt_de")),                sv("unemp_rt_de"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzinade",  fmt_count_000s_change(sv("inact_de")),         sv("inact_de"),        invert = TRUE)
  doc <- fill_conditional(doc, "qvzifede",  fmt_count_000s_change(sv("inact5064_de")),     sv("inact5064_de"),    invert = TRUE)
  doc <- fill_conditional(doc, "qvzirtde",  .format_pp(sv("inact_rt_de")),                 sv("inact_rt_de"),     invert = TRUE)
  doc <- fill_conditional(doc, "qvzifrde",  .format_pp(sv("inact5064_rt_de")),             sv("inact5064_rt_de"), invert = TRUE)
  doc <- fill_conditional(doc, "qvzpayde",  fmt_exempt_change(sv("payroll_de")),            sv("payroll_de"))
  doc <- fill_conditional(doc, "qvzvacde",  fmt_exempt_change(sv("vac_de")),               0, neutral = TRUE)
  doc <- fill_conditional(doc, "qvzwnode",  .format_gbp_signed0(sv("wages_change_election")),     sv("wages_change_election"))
  doc <- fill_conditional(doc, "qvzwcpde",  .format_gbp_signed0(sv("wages_cpi_change_election")), sv("wages_cpi_change_election"))

  # ---- OECD International Comparisons table ----
  oecd_unemp_data <- tryCatch(.read_oecd_latest(file_oecd_unemp), error = function(e) NULL)
  oecd_emp_data   <- tryCatch(.read_oecd_latest(file_oecd_emp),   error = function(e) NULL)
  oecd_inact_data <- tryCatch(.read_oecd_latest(file_oecd_inact), error = function(e) NULL)
  oecd_ft <- .build_oecd_flextable(oecd_unemp_data, oecd_emp_data, oecd_inact_data)

  if (!is.null(oecd_ft)) {
    # Insert OECD section before "External Commentary" paragraph
    # Use officer cursor to find position
    content_summary <- docx_summary(doc)
    ext_comm_idx <- which(grepl("External Commentary", content_summary$text, fixed = TRUE))[1]

    if (!is.na(ext_comm_idx)) {
      # Position cursor just before External Commentary
      doc <- cursor_reach(doc, keyword = "External Commentary")
      # Insert OECD section heading + table + footnotes before External Commentary
      doc <- body_add_par(doc, "", style = "Normal", pos = "before")
      doc <- body_add_par(doc, "OECD International Comparisons", style = "heading 2", pos = "before")
      doc <- body_add_par(doc, "", style = "Normal", pos = "before")
      doc <- body_add_flextable(doc, oecd_ft, pos = "before")
      doc <- body_add_par(doc, "", style = "Normal", pos = "before")
      doc <- body_add_par(doc,
        "*Latest UK data from ONS Labour Force Survey. OECD data schedules may vary by country.",
        style = "Normal", pos = "before")
      doc <- body_add_par(doc,
        "Source: OECD Infra-annual labour statistics. Unemployment rate: aged 15+. Employment and inactivity rates: aged 15-64.",
        style = "Normal", pos = "before")
      doc <- body_add_par(doc, "", style = "Normal", pos = "before")
    } else {
      # Fallback: append at end
      doc <- body_add_par(doc, "")
      doc <- body_add_par(doc, "OECD International Comparisons", style = "heading 2")
      doc <- body_add_par(doc, "")
      doc <- body_add_flextable(doc, oecd_ft)
      doc <- body_add_par(doc, "")
      doc <- body_add_par(doc,
        "*Latest UK data from ONS Labour Force Survey. OECD data schedules may vary by country.")
      doc <- body_add_par(doc,
        "Source: OECD Infra-annual labour statistics. Unemployment rate: aged 15+. Employment and inactivity rates: aged 15-64.")
    }
    if (verbose) message("[manual] OECD comparison table inserted")
  }

  # ---- clean up any unreplaced qvz placeholders ----
  body_xml   <- doc$doc_obj$get()
  ns         <- xml2::xml_ns(body_xml)
  text_nodes <- xml2::xml_find_all(body_xml, ".//w:t", ns = ns)
  for (node in text_nodes) {
    txt <- xml2::xml_text(node)
    if (grepl("qvz", txt, fixed = TRUE)) {
      cleaned <- gsub("qvz[a-z0-9_]+", "\u2014", txt)
      xml2::xml_text(node) <- cleaned
    }
  }

  # ---- write output ----
  print(doc, target = output_path)
  if (verbose) message("[manual] Written to ", output_path)
  invisible(output_path)
}

# Example usage:
#   source("utils/manual_word_output.R")
#   generate_manual_word_output(
#     manual_month    = "feb2026",
#     file_a01        = "path/to/a01feb2026.xlsx",
#     file_x09        = "path/to/x09feb2026.xlsx",
#     file_rtisa      = "path/to/rtisafeb2026.xlsx",
#     file_hr1        = "path/to/hr1feb2026.xlsx",
#     file_oecd_unemp = "path/to/oecd_unemployment.xlsx",
#     file_oecd_emp   = "path/to/oecd_employment.xlsx",
#     file_oecd_inact = "path/to/oecd_inactivity.xlsx"
#   )
