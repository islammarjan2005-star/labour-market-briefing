# manual_word_output.R


suppressPackageStartupMessages({
  library(officer)
  library(xml2)
  library(scales)
  library(readxl)
})

# -- formatters

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

# counts stored 
fmt_count_000s_current <- function(x) .format_int(x / 1000)
fmt_count_000s_change  <- function(x) fmt_int_signed(x / 1000)

# payroll/vacancies 
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

# ----- Word  helpers 

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


sv <- function(name, default = NA_real_) {
  if (exists(name, inherits = TRUE)) get(name, inherits = TRUE) else default
}

# OECD data

# target countries
.oecd_countries <- c(
  "United Kingdom", "United States", "France", "Germany",
  "Italy", "Spain", "Canada", "Japan", "Euro area"
)

# alt labels 
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

#   latest available value per country.
.read_oecd_latest <- function(path) {
  if (is.null(path) || !file.exists(path)) return(NULL)
  
  ext <- tolower(tools::file_ext(path))
  
  # --- Excel with "Table" sheet (wide format: countries as rows, periods as cols) ---
  if (ext %in% c("xlsx", "xls")) {
    sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
    tbl_sheet <- if ("Table" %in% sheets) "Table" else NULL
    
    if (!is.null(tbl_sheet)) {
      raw <- suppressMessages(readxl::read_excel(path, sheet = tbl_sheet, col_names = FALSE))

      #   header row (contains - strings such as "Q1-2025", "2024-Q4", etc.)
      header_row <- NULL
      for (ri in 1:min(15, nrow(raw))) {
        row_text <- as.character(unlist(raw[ri, ]))
        #  is  like "Q1-2025", "2024-Q4", "Jan 2025", "2024-01", etc.
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
    
    #   fallback
    if (length(sheets) > 0) {
      raw <- suppressMessages(readxl::read_excel(path, sheet = sheets[1], n_max = 5000))
      return(.parse_oecd_sdmx(raw))
    }
  }
  
  # CSV long
  if (ext == "csv") {
    raw <- tryCatch(read.csv(path, stringsAsFactors = FALSE), error = function(e) data.frame())
    if (nrow(raw) > 0) return(.parse_oecd_sdmx(raw))
  }
  
  NULL
}

# Parse SDMX 
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
    # Sort  period desc   and take latest
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

# Fill OECD  cells in the temp
.fill_oecd_placeholders <- function(doc, unemp_data, emp_data, inact_data) {
  country_codes <- c("United Kingdom" = "uk", "United States" = "us",
                     "France" = "fr", "Germany" = "de", "Italy" = "it",
                     "Spain" = "es", "Canada" = "ca", "Japan" = "jp",
                     "Euro area" = "ea")
  
  .get_val <- function(data, country) {
    if (is.null(data)) return(list(period = "", value = ""))
    idx <- match(country, data$country)
    if (is.na(idx)) return(list(period = "", value = ""))
    list(period = data$period[idx],
         value = paste0(fmt_one_dec(data$value[idx]), "%"))
  }
  
  for (country in names(country_codes)) {
    cc <- country_codes[[country]]
    u <- .get_val(unemp_data, country)
    e <- .get_val(emp_data, country)
    ia <- .get_val(inact_data, country)
    
    # Time period: prefer unemployment, then employment, then inactivity
    tp <- if (nzchar(u$period)) u$period else if (nzchar(e$period)) e$period else ia$period
    # UK gets asterisk on time period
    if (cc == "uk" && nzchar(tp)) tp <- paste0(tp, "*")
    
    doc <- replace_all(doc, paste0("qvzoecd", cc, "tp"), tp)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "ur"), u$value)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "er"), e$value)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "ir"), ia$value)
  }
  doc
}

#  main

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
  
  # source helpers and run ecxel calculations (auto-detects month from A01 )
  source("utils/helpers.R", local = FALSE)
  source("utils/calculations_from_excel.R", local = FALSE)
  manual_month <- run_calculations_from_excel(manual_month,
                                              file_a01 = file_a01, file_hr1 = file_hr1,
                                              file_x09 = file_x09, file_rtisa = file_rtisa)
  
  if (verbose) message("[manual] Calculations complete for ", manual_month)
  
  # generate summary and topten  lines
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
  
  # - header / labels
  doc <- replace_all(doc, "qvzmonthlabel", manual_month_to_label(manual_month))
  doc <- replace_all(doc, "qvzrenderdate", format(Sys.Date(), "%d %B %Y"))
  doc <- replace_all(doc, "qvzlfsperiod",  sv("lfs_period_label", ""))
  doc <- replace_all(doc, "qvzlfsquarter", sv("lfs_period_short_label", ""))
  doc <- replace_all(doc, "qvzvacquarter", sv("vacancies_period_short_label", ""))
  
  # -summary + top ten lines
  for (i in 1:10) doc <- replace_all(doc, sprintf("qvzsl%02d", i), summary[[paste0("line", i)]])
  for (i in 1:10) doc <- replace_all(doc, sprintf("qvztt%02d", i), top10[[paste0("line", i)]])
  
  # -Stats db Current column
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
  
  # The OECD page is already in ManualDB.docx 
  oecd_unemp_data <- tryCatch(.read_oecd_latest(file_oecd_unemp), error = function(e) NULL)
  oecd_emp_data   <- tryCatch(.read_oecd_latest(file_oecd_emp),   error = function(e) NULL)
  oecd_inact_data <- tryCatch(.read_oecd_latest(file_oecd_inact), error = function(e) NULL)
  
  if (!is.null(oecd_unemp_data) || !is.null(oecd_emp_data) || !is.null(oecd_inact_data)) {
    doc <- .fill_oecd_placeholders(doc, oecd_unemp_data, oecd_emp_data, oecd_inact_data)
    if (verbose) message("[manual] OECD comparison table populated")
  }
  
  # clean up any  nopt used qvz placeholders
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

# Exampl     :
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