# labour market statistics brief - shiny application

library(shiny)

# ui - gov.uk design system

ui <- fluidPage(
  
  # page title
  tags$head(
    tags$title("Labour Market Statistics Brief"),
    tags$meta(charset = "utf-8"),
    tags$meta(name = "viewport", content = "width=device-width, initial-scale=1"),
    
    # gov.uk design system css
    tags$style(HTML("
      @import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@400;600;700&display=swap');

      *, *::before, *::after { box-sizing: border-box; }
      html, body { margin: 0; padding: 0; min-height: 100vh; }

      body {
        font-family: 'Source Sans Pro', Arial, sans-serif;
        font-size: 19px;
        line-height: 1.31579;
        color: #0b0c0c;
        background-color: #f3f2f1;
      }

      .govuk-header {
        border-bottom: 10px solid #1d70b8;
        color: #ffffff;
        background: #0b0c0c;
      }

      .govuk-header__container {
        position: relative;
        margin-bottom: -10px;
        padding-top: 10px;
        border-bottom: 10px solid #1d70b8;
        max-width: 960px;
        margin-left: auto;
        margin-right: auto;
        padding-left: 15px;
        padding-right: 15px;
      }

      .govuk-header__logotype-text {
        font-weight: 700;
        font-size: 30px;
        line-height: 1;
        color: #ffffff;
      }

      .govuk-header__link {
        text-decoration: none;
        color: #ffffff;
      }

      .govuk-header__service-name {
        display: inline-block;
        margin-bottom: 10px;
        font-weight: 700;
        font-size: 24px;
        color: #ffffff;
      }

      .govuk-width-container {
        max-width: 960px;
        margin: 0 auto;
        padding: 0 15px;
      }

      .govuk-main-wrapper {
        padding: 40px 0;
      }

      .govuk-heading-xl {
        font-weight: 700;
        font-size: 48px;
        margin: 0 0 50px 0;
      }

      .govuk-heading-m {
        font-weight: 700;
        font-size: 24px;
        margin: 0 0 20px 0;
      }

      .govuk-body {
        font-size: 19px;
        margin: 0 0 20px 0;
      }

      .govuk-button {
        font-weight: 400;
        font-size: 19px;
        line-height: 1;
        display: inline-block;
        position: relative;
        padding: 8px 10px 7px;
        border: 2px solid transparent;
        border-radius: 0;
        color: #ffffff;
        background-color: #00703c;
        box-shadow: 0 2px 0 #002d18;
        text-align: center;
        cursor: pointer;
        margin-right: 15px;
        margin-bottom: 15px;
      }

      .govuk-button:hover { background-color: #005a30; }
      .govuk-button:focus {
        border-color: #ffdd00;
        outline: 3px solid transparent;
        box-shadow: inset 0 0 0 1px #ffdd00;
        background-color: #ffdd00;
        color: #0b0c0c;
      }

      .govuk-button--blue {
        background-color: #1d70b8;
        box-shadow: 0 2px 0 #003078;
      }
      .govuk-button--blue:hover { background-color: #003078; }

      .govuk-button--secondary {
        background-color: #f3f2f1;
        box-shadow: 0 2px 0 #929191;
        color: #0b0c0c;
      }
      .govuk-button--secondary:hover { background-color: #dbdad9; }

      .govuk-button--warning {
        background-color: #d4351c;
        box-shadow: 0 2px 0 #6e1509;
      }
      .govuk-button--warning:hover { background-color: #aa2a16; }

      .govuk-button.shiny-download-link { text-decoration: none; }

      .govuk-form-group { margin-bottom: 30px; }

      .govuk-label {
        font-weight: 400;
        font-size: 19px;
        display: block;
        margin-bottom: 5px;
      }

      .govuk-hint {
        font-size: 19px;
        margin-bottom: 15px;
        color: #505a5f;
      }

      .govuk-input {
        font-size: 19px;
        width: 100%;
        max-width: 200px;
        height: 40px;
        padding: 5px;
        border: 2px solid #0b0c0c;
        border-radius: 0;
      }

      .govuk-input:focus {
        outline: 3px solid #ffdd00;
        outline-offset: 0;
        box-shadow: inset 0 0 0 2px;
      }

      .govuk-select {
        font-size: 19px;
        width: 100%;
        max-width: 320px;
        height: 40px;
        padding: 5px;
        border: 2px solid #0b0c0c;
        border-radius: 0;
        background-color: #ffffff;
      }

      .govuk-select:focus {
        outline: 3px solid #ffdd00;
        outline-offset: 0;
        box-shadow: inset 0 0 0 2px;
      }

      .govuk-width-container--wide {
        max-width: 1400px;
      }

      .preview-scroll {
        max-height: 460px;
        overflow: auto;
      }

      .govuk-phase-banner {
        padding: 10px 0;
        border-bottom: 1px solid #b1b4b6;
      }

      .govuk-tag {
        font-weight: 700;
        font-size: 16px;
        display: inline-block;
        padding: 5px 8px 4px;
        color: #ffffff;
        background-color: #1d70b8;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-right: 10px;
      }

      .govuk-tag--green {
        background-color: #00703c;
      }

      .dashboard-card {
        background-color: #ffffff;
        border: 1px solid #b1b4b6;
        margin-bottom: 20px;
      }

      .dashboard-card__header {
        background-color: #1d70b8;
        color: #ffffff;
        padding: 15px 20px;
        font-weight: 700;
        font-size: 19px;
      }

      .dashboard-card__content { padding: 20px; }

      .govuk-section-break {
        margin: 30px 0;
        border: 0;
        border-bottom: 1px solid #b1b4b6;
      }

      .govuk-grid-row {
        display: flex;
        flex-wrap: wrap;
        margin: 0 -15px;
      }

      .govuk-grid-column-one-half {
        width: 50%;
        padding: 0 15px;
      }

      @media (max-width: 768px) {
        .govuk-grid-column-one-half { width: 100%; }
      }

      .govuk-footer {
        padding: 25px 0;
        border-top: 1px solid #b1b4b6;
        background: #f3f2f1;
        text-align: center;
        color: #505a5f;
      }

      .container-fluid { padding: 0 !important; margin: 0 !important; max-width: none !important; }

      /* Month confirmation status */
      .month-status {
        display: inline-block;
        padding: 5px 10px;
        margin-left: 10px;
        font-size: 16px;
        border-radius: 3px;
      }

      .month-status--confirmed {
        background-color: #00703c;
        color: #ffffff;
      }

      .month-status--pending {
        background-color: #f47738;
        color: #ffffff;
      }

      /* Input row with button */
      .input-row {
        display: flex;
        align-items: flex-end;
        gap: 15px;
        flex-wrap: wrap;
      }

      .input-row .govuk-form-group {
        margin-bottom: 0;
      }

      /* Stats table */
      .stats-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
      }

      .stats-table th {
        background-color: #0b0c0c;
        color: #ffffff;
        font-weight: 700;
        padding: 10px 8px;
        text-align: left;
        border: 1px solid #0b0c0c;
        font-size: 12px;
      }

      .stats-table td {
        padding: 8px;
        border: 1px solid #b1b4b6;
        background-color: #ffffff;
      }

      .stats-table tr:nth-child(even) td { background-color: #f8f8f8; }
      .stats-table tr:hover td { background-color: #f3f2f1; }

      .stat-positive { color: #00703c; font-weight: 700; }
      .stat-negative { color: #d4351c; font-weight: 700; }
      .stat-neutral { color: #505a5f; }

      /* Top Ten List */
      .top-ten-list {
        list-style: none;
        padding: 0;
        margin: 0;
        counter-reset: item;
      }

      .top-ten-list li {
        padding: 12px 12px 12px 50px;
        margin-bottom: 8px;
        background-color: #ffffff;
        border-left: 4px solid #1d70b8;
        position: relative;
        font-size: 15px;
        line-height: 1.4;
      }

      .top-ten-list li::before {
        counter-increment: item;
        content: counter(item);
        position: absolute;
        left: 12px;
        top: 12px;
        font-weight: 700;
        font-size: 18px;
        color: #1d70b8;
      }

      .govuk-list { padding-left: 20px; }
      .govuk-list li { margin-bottom: 5px; }

      /* Shiny progress bar customization */
      .shiny-notification {
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 400px;
        background: #ffffff;
        border: 3px solid #1d70b8;
        border-radius: 0;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        padding: 20px;
        z-index: 99999;
      }

      .shiny-notification-message {
        font-family: 'Source Sans Pro', Arial, sans-serif;
        font-size: 16px;
        color: #0b0c0c;
        margin-bottom: 15px;
      }

      .shiny-notification .progress {
        height: 10px;
        background-color: #f3f2f1;
        border-radius: 0;
        margin-top: 10px;
      }

      .shiny-notification .progress-bar {
        background-color: #00703c;
        border-radius: 0;
      }

      .shiny-notification-close {
        display: none;
      }

      /* Loading spinner (used during auto reference-month detection) */
      .loader {
        border: 4px solid #f3f2f1;
        border-top: 4px solid #1d70b8;
        border-radius: 50%;
        width: 28px;
        height: 28px;
        animation: spin 0.9s linear infinite;
        display: inline-block;
        vertical-align: middle;
        margin-right: 12px;
      }
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }

      /* Ensure Shiny selectInput matches GOV.UK style */
      #vacancies_period, #payroll_period,
      #manual_vacancies_period, #manual_payroll_period {
        font-size: 19px;
        width: 100%;
        max-width: 320px;
        height: 40px;
        padding: 5px;
        border: 2px solid #0b0c0c;
        border-radius: 0;
        background-color: #ffffff;
      }
      #vacancies_period:focus, #payroll_period:focus,
      #manual_vacancies_period:focus, #manual_payroll_period:focus {
        outline: 3px solid #ffdd00;
        outline-offset: 0;
        box-shadow: inset 0 0 0 2px;
      }

      /* Two-column mode layout */
      .mode-row {
        display: flex;
        gap: 20px;
        margin-bottom: 20px;
      }
      .mode-col {
        flex: 1;
        min-width: 0;
      }
      .mode-col .dashboard-card {
        height: 100%;
        display: flex;
        flex-direction: column;
      }
      .mode-col .dashboard-card__content {
        flex: 1;
      }
      .mode-col--manual {
        border-left: 3px solid #b1b4b6;
        padding-left: 20px;
      }
      .govuk-tag--amber {
        background-color: #f47738;
      }
      @media (max-width: 768px) {
        .mode-row { flex-direction: column; }
        .mode-col--manual { border-left: none; padding-left: 0; border-top: 3px solid #b1b4b6; padding-top: 20px; }
      }
    "))
  ),
  
  # header
  tags$header(class = "govuk-header",
              div(class = "govuk-header__container",
                  div(style = "margin-bottom: 10px;",
                      a(href = "#", class = "govuk-header__link",
                        span(class = "govuk-header__logotype-text", "GOV.UK")
                      )
                  ),
                  span(class = "govuk-header__service-name", "Labour Market Statistics Brief")
              )
  ),
  
  # main content
  div(class = "govuk-width-container",
      
      div(class = "govuk-phase-banner",
          span(class = "govuk-tag", "BETA"),
          span("This is a new service.")
      ),
      
      tags$main(class = "govuk-main-wrapper",
                
                h1(class = "govuk-heading-xl", "Labour Market Statistics Brief Generator"),
                
                # two-column layout: auto (left) vs manual (right)
                div(class = "mode-row",
                    
                    # ---- LEFT COLUMN: Automatic (Database) ----
                    div(class = "mode-col",
                        div(class = "dashboard-card",
                            div(class = "dashboard-card__header", "Automatic (Database)"),
                            div(class = "dashboard-card__content",
                                
                                div(class = "govuk-form-group",
                                    tags$label(class = "govuk-label", "Reference month"),
                                    div(class = "govuk-hint", "Auto-selected from latest available data"),
                                    uiOutput("month_status")
                                ),
                                
                                div(class = "input-row",
                                    div(class = "govuk-form-group",
                                        tags$label(class = "govuk-label", `for` = "vacancies_period", "Vacancies"),
                                        selectInput("vacancies_period", label = NULL, choices = c("Loading" = "Loading"), selected = "Loading")
                                    ),
                                    div(class = "govuk-form-group",
                                        tags$label(class = "govuk-label", `for` = "payroll_period", "Payroll employees"),
                                        selectInput("payroll_period", label = NULL, choices = c("Loading" = "Loading"), selected = "Loading")
                                    )
                                ),
                                
                                tags$hr(class = "govuk-section-break"),
                                
                                h2(class = "govuk-heading-m", "Preview"),
                                actionButton("preview_dashboard", "Preview Dashboard", class = "govuk-button govuk-button--blue"),
                                actionButton("preview_topten", "Preview Top Ten Stats", class = "govuk-button govuk-button--blue"),
                                
                                tags$hr(class = "govuk-section-break"),
                                
                                h2(class = "govuk-heading-m", "Download"),
                                downloadButton("download_word", "Download Word", class = "govuk-button"),
                                downloadButton("download_excel", "Download Excel", class = "govuk-button govuk-button--secondary")
                            )
                        )
                    ),
                    
                    # ---- RIGHT COLUMN: Manual (Excel Upload) ----
                    div(class = "mode-col mode-col--manual",
                        div(class = "dashboard-card",
                            div(class = "dashboard-card__header", "Manual (Excel Upload)"),
                            div(class = "dashboard-card__content",
                                
                                p(class = "govuk-body", "Upload ONS Excel files to generate from local data. A01 is required; others are optional."),
                                tags$ul(class = "govuk-list",
                                        tags$li(tags$a(href = "https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/employmentandemployeetypes/datasets/summaryoflabourmarketstatistics/current",
                                                       target = "_blank", "A01: Summary of labour market statistics")),
                                        tags$li(tags$a(href = "https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/employmentandemployeetypes/datasets/hr1potentialredundancies",
                                                       target = "_blank", "HR1: Potential redundancies")),
                                        tags$li(tags$a(href = "https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/earningsandworkinghours/datasets/x09realaverageweeklyearningsusingconsumerpriceinflationseasonallyadjusted",
                                                       target = "_blank", "X09: Real average weekly earnings (CPI)")),
                                        tags$li(tags$a(href = "https://www.ons.gov.uk/employmentandlabourmarket/peopleinwork/earningsandworkinghours/datasets/realtimeinformationstatisticsreferencetableseasonallyadjusted",
                                                       target = "_blank", "RTISA: Payrolled employees (SA)"))
                                ),
                                
                                div(class = "govuk-form-group",
                                    fileInput("upload_files", "Upload core Excel files (.xlsx)",
                                              accept = ".xlsx", multiple = TRUE, width = "100%"),
                                    div(class = "govuk-hint",
                                        "Select up to 4 files at once. Files are auto-detected by name (e.g. a01, hr1, x09, rtisa).")
                                ),
                                
                                uiOutput("upload_status"),
                                
                                div(class = "input-row",
                                    div(class = "govuk-form-group",
                                        tags$label(class = "govuk-label", `for` = "manual_vacancies_period", "Vacancies"),
                                        selectInput("manual_vacancies_period", label = NULL, choices = c("Upload A01 first" = ""), selected = "")
                                    ),
                                    div(class = "govuk-form-group",
                                        tags$label(class = "govuk-label", `for` = "manual_payroll_period", "Payroll employees"),
                                        selectInput("manual_payroll_period", label = NULL, choices = c("Upload RTISA first" = ""), selected = "")
                                    )
                                ),
                                
                                tags$hr(class = "govuk-section-break"),
                                
                                h2(class = "govuk-heading-m", "Preview"),
                                actionButton("manual_preview_dashboard", "Preview Dashboard", class = "govuk-button govuk-button--blue"),
                                actionButton("manual_preview_topten", "Preview Top Ten Stats", class = "govuk-button govuk-button--blue"),
                                
                                tags$hr(class = "govuk-section-break"),
                                
                                h2(class = "govuk-heading-m", "Download"),
                                downloadButton("manual_download_word", "Download Word", class = "govuk-button"),
                                downloadButton("manual_download_excel", "Download Excel", class = "govuk-button govuk-button--secondary"),
                                
                                tags$hr(class = "govuk-section-break govuk-section-break--visible", style = "margin-top: 20px; margin-bottom: 15px;"),
                                
                                h2(class = "govuk-heading-s", "Supplementary files"),
                                p(class = "govuk-body-s govuk-!-margin-bottom-2",
                                  "Optional — for the full Excel audit workbook (CLA01, X02, OECD)."),
                                
                                fileInput("upload_extra_files", NULL,
                                          accept = c(".xlsx", ".csv"), multiple = TRUE, width = "100%"),
                                
                                div(class = "govuk-hint govuk-body-s",
                                    "Name files with their type prefix (e.g. cla01_feb2026.xlsx, x02_feb2026.xlsx,",
                                    "oecd_unemployment.csv). OECD files can be .csv or .xlsx."),
                                
                                uiOutput("extra_upload_status")
                            )
                        )
                    )
                )
      )
  ),
  
  # full-width preview area (stacked so it is readable)
  div(class = "govuk-width-container govuk-width-container--wide",
      tags$main(class = "govuk-main-wrapper", style = "padding-top: 0;",
                div(class = "dashboard-card",
                    div(class = "dashboard-card__header", "Dashboard Preview"),
                    div(class = "dashboard-card__content preview-scroll", uiOutput("dashboard_preview"))
                ),
                div(class = "dashboard-card",
                    div(class = "dashboard-card__header", "Top Ten Statistics Preview"),
                    div(class = "dashboard-card__content", uiOutput("topten_preview"))
                )
      )
  ),
  
  # footer
  tags$footer(class = "govuk-footer",
              div(class = "govuk-width-container",
                  "Labour Market Statistics Brief Generator | Department for Business and Trade"
              )
  )
)

# server

server <- function(input, output, session) {
  
  # file paths
  config_path       <- "utils/config.R"
  calculations_path <- "utils/calculations.R"
  excel_calc_path   <- "utils/calculations_from_excel.R"
  word_script_path  <- "utils/word_output.R"
  excel_script_path <- "sheets/excel_audit_workbook.R"
  summary_path      <- "sheets/summary.R"
  top_ten_path      <- "sheets/top_ten_stats.R"
  template_path     <- "utils/DB.docx"
  
  # helper: check if A01 (minimum required) has been uploaded
  has_uploads <- function() {
    !is.null(uploaded_files$a01)
  }
  
  has_any_upload <- function() {
    !is.null(uploaded_files$a01) || !is.null(uploaded_files$hr1) ||
      !is.null(uploaded_files$x09) || !is.null(uploaded_files$rtisa) ||
      !is.null(uploaded_files$cla01) || !is.null(uploaded_files$x02) ||
      !is.null(uploaded_files$oecd_unemp)
  }
  
  # reactive values
  dashboard_data <- reactiveVal(NULL)
  topten_data <- reactiveVal(NULL)
  
  # uploaded file paths (NULL = not uploaded, use DB)
  uploaded_files <- reactiveValues(
    a01 = NULL,
    hr1 = NULL,
    x09 = NULL,
    rtisa = NULL,
    cla01 = NULL,
    x02 = NULL,
    oecd_unemp = NULL,
    oecd_emp = NULL,
    oecd_inact = NULL
  )
  
  # Validate Excel sheets on upload and warn if wrong file
  .validate_excel <- function(path, expected_sheets, file_label) {
    sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
    missing <- setdiff(expected_sheets, sheets)
    if (length(missing) > 0) {
      showNotification(
        paste0("This doesn't look like a ", file_label, " file. Missing sheets: ",
               paste(missing, collapse = ", ")),
        type = "warning", duration = 10
      )
    }
  }
  
  # auto-detect uploaded files by filename and sheet contents
  .detect_file_type <- function(name, path) {
    nm <- tolower(name)
    # match by filename — check longer/more-specific names FIRST
    if (grepl("cla01", nm)) return("cla01")
    if (grepl("rtisa", nm)) return("rtisa")
    # OECD files — detect by keyword in filename, then inspect contents
    if (grepl("oecd", nm)) return(.detect_oecd_type(nm, path))
    if (grepl("a01", nm)) return("a01")
    if (grepl("hr1", nm)) return("hr1")
    if (grepl("x09", nm)) return("x09")
    if (grepl("x02", nm)) return("x02")
    # fallback: match by sheet names
    sheets <- tryCatch(tolower(readxl::excel_sheets(path)), error = function(e) character(0))
    if (any(sheets %in% c("1", "10", "13", "15", "19"))) return("a01")
    if (any(sheets %in% c("1a"))) return("hr1")
    if (any(grepl("awe real_cpi", sheets))) return("x09")
    if (any(grepl("payrolled employees", sheets))) return("rtisa")
    if (any(grepl("claimant", sheets)) || any(grepl("people sa", sheets))) return("cla01")
    if (any(grepl("labour market flows", sheets))) return("x02")
    # Last resort: check if this is an OECD file by inspecting contents
    # (catches files with non-standard names like OECD.SDD.TPS.DSD_LFS@...)
    oecd_result <- .detect_oecd_metric_from_content(path)
    if (!is.null(oecd_result)) return(oecd_result)
    NULL
  }
  
  # OECD content-based metric detection — inspects file contents to determine
  # if it's unemployment, employment, or inactivity data (since OECD downloads
  # all have identical filenames differing only by timestamp).
  .detect_oecd_metric_from_content <- function(path) {
    tryCatch({
      ext <- tolower(tools::file_ext(path))
      if (ext %in% c("xlsx", "xls")) {
        sheets <- readxl::excel_sheets(path)
        # Try "Table" sheet firsthas 
        tbl_sheet <- if ("Table" %in% sheets) "Table" else NULL
        if (!is.null(tbl_sheet)) {
          tbl <- suppressMessages(readxl::read_excel(path, sheet = tbl_sheet,
                                                     col_names = FALSE, n_max = 10))
          all_text <- tolower(paste(unlist(tbl), collapse = " "))
          if (grepl("unemployment", all_text)) return("oecd_unemp")
          if (grepl("inactivity", all_text))   return("oecd_inact")
          if (grepl("employment", all_text))   return("oecd_emp")
        }
        # Fallback:  SDMX   
        check_sheet <- if (length(sheets) > 0) sheets[1] else NULL
        if (!is.null(check_sheet)) {
          raw <- suppressMessages(readxl::read_excel(path, sheet = check_sheet, n_max = 5))
          measure_col <- intersect(c("MEASURE", "Measure"), names(raw))
          if (length(measure_col) > 0) {
            vals <- tolower(paste(raw[[measure_col[1]]], collapse = " "))
            if (grepl("une", vals))   return("oecd_unemp")
            if (grepl("inact", vals)) return("oecd_inact")
            if (grepl("emp", vals))   return("oecd_emp")
          }
        }
      } else if (ext == "csv") {
        raw <- read.csv(path, nrows = 50, stringsAsFactors = FALSE)
        measure_col <- intersect(c("MEASURE", "Measure"), names(raw))
        if (length(measure_col) > 0) {
          vals <- tolower(paste(unique(raw[[measure_col[1]]]), collapse = " "))
          if (grepl("une", vals))   return("oecd_unemp")
          if (grepl("inact", vals)) return("oecd_inact")
          if (grepl("emp", vals))   return("oecd_emp")
        }
        # Fallback: scan all column names and values for keywords
        all_text <- tolower(paste(c(names(raw), unlist(raw[1:min(5, nrow(raw)), ])), collapse = " "))
        if (grepl("unemployment", all_text)) return("oecd_unemp")
        if (grepl("inactivity", all_text))   return("oecd_inact")
        if (grepl("employment", all_text))   return("oecd_emp")
      }
      NULL
    }, error = function(e) NULL)
  }
  
  # OECD sub-type detection: try filename keywords first, then content inspection
  .detect_oecd_type <- function(nm, path = NULL) {
    if (grepl("unemp", nm))                      return("oecd_unemp")
    if (grepl("emp", nm) && !grepl("unemp", nm)) return("oecd_emp")
    if (grepl("inact", nm))                      return("oecd_inact")
    # Filename no't help — inspect file contents
    if (!is.null(path)) {
      result <- .detect_oecd_metric_from_content(path)
      if (!is.null(result)) return(result)
    }
    NULL
  }
  
  # track uploads 
  observeEvent(input$upload_files, {
    files <- input$upload_files
    if (is.null(files)) return()
    
    detected_types <- character(0)
    
    for (i in seq_len(nrow(files))) {
      ftype <- .detect_file_type(files$name[i], files$datapath[i])
      if (is.null(ftype)) {
        showNotification(
          paste0("Could not identify file: ", files$name[i],
                 ". OECD files are auto-detected by content; other files need ",
                 "recognisable names (A01, HR1, X09, RTISA)."),
          type = "warning", duration = 10
        )
        next
      }
      
      # validate and store
      if (ftype == "a01") {
        .validate_excel(files$datapath[i], c("1", "10", "13", "15", "19"), "A01")
        uploaded_files$a01 <- files$datapath[i]
        # detect reference month from A01
        tryCatch({
          if (!exists(".detect_manual_month_from_a01", inherits = TRUE)) {
            source("utils/calculations_from_excel.R", local = FALSE)
          }
          detected <- .detect_manual_month_from_a01(files$datapath[i])
          if (!is.null(detected)) .update_ref_month(detected)
        }, error = function(e) NULL)
      } else if (ftype == "hr1") {
        .validate_excel(files$datapath[i], c("1a"), "HR1")
        uploaded_files$hr1 <- files$datapath[i]
      } else if (ftype == "x09") {
        .validate_excel(files$datapath[i], c("AWE Real_CPI"), "X09")
        uploaded_files$x09 <- files$datapath[i]
      } else if (ftype == "rtisa") {
        .validate_excel(files$datapath[i], c("1. Payrolled employees (UK)", "23. Employees (Industry)"), "RTISA")
        uploaded_files$rtisa <- files$datapath[i]
      } else if (ftype == "cla01") {
        uploaded_files$cla01 <- files$datapath[i]
      } else if (ftype == "x02") {
        uploaded_files$x02 <- files$datapath[i]
      } else if (ftype == "oecd_unemp") {
        uploaded_files$oecd_unemp <- files$datapath[i]
      } else if (ftype == "oecd_emp") {
        uploaded_files$oecd_emp <- files$datapath[i]
      } else if (ftype == "oecd_inact") {
        uploaded_files$oecd_inact <- files$datapath[i]
      }
      
      .oecd_friendly <- c(oecd_unemp = "OECD Unemployment Rate",
                          oecd_emp = "OECD Employment Rate",
                          oecd_inact = "OECD Inactivity Rate")
      display <- if (ftype %in% names(.oecd_friendly)) .oecd_friendly[[ftype]] else toupper(ftype)
      detected_types <- c(detected_types, display)
    }
    
    if (length(detected_types) > 0) {
      showNotification(
        paste0("Detected: ", paste(detected_types, collapse = ", ")),
        type = "message", duration = 5
      )
    }
  })
  
  # handle supplementary file uploads (CLA01, X02, OECD)
  observeEvent(input$upload_extra_files, {
    files <- input$upload_extra_files
    if (is.null(files)) return()
    
    detected_types <- character(0)
    
    for (i in seq_len(nrow(files))) {
      ftype <- .detect_file_type(files$name[i], files$datapath[i])
      if (is.null(ftype)) {
        showNotification(
          paste0("Could not identify file: ", files$name[i],
                 ". OECD files are auto-detected by content; other files need ",
                 "recognisable names (CLA01, X02, etc.)."),
          type = "warning", duration = 10
        )
        next
      }
      
      # store recognised types
      if (ftype == "cla01") {
        uploaded_files$cla01 <- files$datapath[i]
      } else if (ftype == "x02") {
        uploaded_files$x02 <- files$datapath[i]
      } else if (ftype == "oecd_unemp") {
        uploaded_files$oecd_unemp <- files$datapath[i]
      } else if (ftype == "oecd_emp") {
        uploaded_files$oecd_emp <- files$datapath[i]
      } else if (ftype == "oecd_inact") {
        uploaded_files$oecd_inact <- files$datapath[i]
      } else {
        # core files can also be uploaded via supplementary uploader
        if (ftype == "a01") uploaded_files$a01 <- files$datapath[i]
        if (ftype == "hr1") uploaded_files$hr1 <- files$datapath[i]
        if (ftype == "x09") uploaded_files$x09 <- files$datapath[i]
        if (ftype == "rtisa") uploaded_files$rtisa <- files$datapath[i]
      }
      
      .oecd_friendly <- c(oecd_unemp = "OECD Unemployment Rate",
                          oecd_emp = "OECD Employment Rate",
                          oecd_inact = "OECD Inactivity Rate")
      display <- if (ftype %in% names(.oecd_friendly)) .oecd_friendly[[ftype]] else toupper(ftype)
      detected_types <- c(detected_types, display)
    }
    
    if (length(detected_types) > 0) {
      showNotification(
        paste0("Detected: ", paste(detected_types, collapse = ", ")),
        type = "message", duration = 5
      )
    }
  })
  
  # upload status display — core files with detected month
  output$upload_status <- renderUI({
    core_files <- list(
      A01   = uploaded_files$a01,
      HR1   = uploaded_files$hr1,
      X09   = uploaded_files$x09,
      RTISA = uploaded_files$rtisa
    )
    
    n_up <- sum(!vapply(core_files, is.null, logical(1)))
    if (n_up == 0) return(NULL)
    
    core_tags <- lapply(names(core_files), function(nm) {
      if (!is.null(core_files[[nm]])) {
        span(class = "govuk-tag govuk-tag--green", style = "margin-right: 4px;",
             paste0(nm, " \u2713"))
      } else {
        span(class = "govuk-tag govuk-tag--grey", style = "margin-right: 4px;", nm)
      }
    })
    
    mm <- reference_manual_month()
    month_line <- if (!is.null(mm) && nzchar(mm) && !is.null(uploaded_files$a01)) {
      div(style = "margin-top: 6px; font-weight: 600;",
          paste0("Reference month: ", manual_month_to_display(mm)))
    }
    
    div(style = "margin-top: 10px;",
        tagList(core_tags),
        month_line)
  })
  
  # supplementary file status — shown below supplementary upload area
  output$extra_upload_status <- renderUI({
    extra_files <- list(
      CLA01 = uploaded_files$cla01,
      X02   = uploaded_files$x02,
      `OECD Unemp` = uploaded_files$oecd_unemp,
      `OECD Emp`   = uploaded_files$oecd_emp,
      `OECD Inact` = uploaded_files$oecd_inact
    )
    
    n_extra <- sum(!vapply(extra_files, is.null, logical(1)))
    if (n_extra == 0) return(NULL)
    
    extra_tags <- lapply(names(extra_files), function(nm) {
      if (!is.null(extra_files[[nm]])) {
        span(class = "govuk-tag govuk-tag--blue", style = "margin-right: 4px;",
             paste0(nm, " \u2713"))
      }
    })
    
    div(style = "margin-top: 6px;", tagList(extra_tags))
  })
  
  # Detect vacancies periods from A01 and populate manual dropdown
  observeEvent(uploaded_files$a01, {
    path <- uploaded_files$a01
    if (is.null(path)) return()
    
    tryCatch({
      mm <- reference_manual_month()
      if (is.null(mm) || !nzchar(mm)) return()
      
      mm_mon3 <- substr(gsub("[^a-z]", "", mm), 1, 3)
      mm_yr   <- as.integer(substr(gsub("[^0-9]", "", mm), 1, 4))
      mm_m    <- match(mm_mon3, tolower(month.abb))
      mm_date <- as.Date(sprintf("%04d-%02d-01", mm_yr, mm_m))
      lfs_end <- add_months(mm_date, -2)
      
      tbl_19 <- readxl::read_excel(path, sheet = "19", col_names = FALSE, .name_repair = "minimal")
      periods <- trimws(as.character(tbl_19[[1]]))
      # Parse "Mmm-Mmm YYYY" labels to get end dates
      parsed <- vapply(periods, function(x) {
        m <- regmatches(x, gregexpr("[A-Za-z]{3}", x))[[1]]
        y <- regmatches(x, gregexpr("[0-9]{4}", x))[[1]]
        if (length(m) >= 2 && length(y) >= 1) {
          month_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)
          em <- month_map[tolower(m[2])]
          if (!is.na(em)) return(as.numeric(as.Date(sprintf("%s-%02d-01", y[1], em))))
        }
        NA_real_
      }, numeric(1))
      
      ok <- !is.na(parsed) & !is.na(suppressWarnings(as.numeric(tbl_19[[3]])))
      if (!any(ok)) return()
      
      ends <- as.Date(parsed[ok], origin = "1970-01-01")
      end_latest <- max(ends)
      aligned_candidates <- ends[ends <= lfs_end]
      end_aligned <- if (length(aligned_candidates) > 0) max(aligned_candidates) else end_latest
      
      make_label <- function(end_d) {
        start_d <- add_months(end_d, -2)
        paste0(format(start_d, "%b"), "-", format(end_d, "%b"), " ", format(end_d, "%Y"))
      }
      
      lab_aligned <- make_label(end_aligned)
      lab_latest  <- make_label(end_latest)
      
      manual_period_labels$vac_aligned <- lab_aligned
      manual_period_labels$vac_latest  <- lab_latest
      
      if (identical(lab_aligned, lab_latest)) {
        choices <- setNames(lab_aligned, paste0(lab_aligned, " (only period)"))
      } else {
        choices <- setNames(
          c(lab_aligned, lab_latest),
          c(paste0(lab_aligned, " (default)"), lab_latest)
        )
      }
      updateSelectInput(session, "manual_vacancies_period", choices = choices, selected = choices[1])
    }, error = function(e) NULL)
  })
  
  # Detect payroll periods from RTISA and populate manual dropdown
  observeEvent(uploaded_files$rtisa, {
    path <- uploaded_files$rtisa
    if (is.null(path)) return()
    
    tryCatch({
      mm <- reference_manual_month()
      if (is.null(mm) || !nzchar(mm)) return()
      
      mm_mon3 <- substr(gsub("[^a-z]", "", mm), 1, 3)
      mm_yr   <- as.integer(substr(gsub("[^0-9]", "", mm), 1, 4))
      mm_m    <- match(mm_mon3, tolower(month.abb))
      mm_date <- as.Date(sprintf("%04d-%02d-01", mm_yr, mm_m))
      lfs_end <- add_months(mm_date, -2)
      
      rtisa_pay <- readxl::read_excel(path, sheet = "1. Payrolled employees (UK)", col_names = FALSE, .name_repair = "minimal")
      text_col <- trimws(as.character(rtisa_pay[[1]]))
      parsed <- suppressWarnings(lubridate::parse_date_time(text_col, orders = c("B Y", "bY", "BY")))
      months_parsed <- lubridate::floor_date(as.Date(parsed), "month")
      vals <- suppressWarnings(as.numeric(gsub("[^0-9.-]", "", as.character(rtisa_pay[[2]]))))
      
      ok <- !is.na(months_parsed) & !is.na(vals)
      if (!any(ok)) return()
      
      avail <- months_parsed[ok]
      end_latest <- max(avail)
      aligned_candidates <- avail[avail <= lfs_end]
      end_aligned <- if (length(aligned_candidates) > 0) max(aligned_candidates) else end_latest
      
      make_label <- function(end_d) {
        start_d <- add_months(end_d, -2)
        paste0(format(start_d, "%b"), "-", format(end_d, "%b"), " ", format(end_d, "%Y"))
      }
      
      lab_aligned <- make_label(end_aligned)
      lab_latest  <- make_label(end_latest)
      
      manual_period_labels$pay_aligned <- lab_aligned
      manual_period_labels$pay_latest  <- lab_latest
      
      if (identical(lab_aligned, lab_latest)) {
        choices <- setNames(lab_latest, paste0(lab_latest, " (only period)"))
      } else {
        choices <- setNames(
          c(lab_aligned, lab_latest),
          c(paste0(lab_aligned, " (default)"), lab_latest)
        )
      }
      updateSelectInput(session, "manual_payroll_period", choices = choices, selected = choices[1])
    }, error = function(e) NULL)
  })
  
  reference_manual_month <- reactiveVal(NULL)
  manual_period_labels <- reactiveValues(
    vac_aligned = NULL, vac_latest = NULL,
    pay_aligned = NULL, pay_latest = NULL
  )
  period_labels <- reactiveVal(list(
    vac = list(aligned = NULL, latest = NULL),
    payroll = list(aligned = NULL, latest = NULL)
  ))
  
  # small date helpers (no extra packages)
  add_months <- function(d, n) {
    d <- as.Date(d)
    if (is.na(d)) return(as.Date(NA))
    if (n == 0) return(d)
    if (n > 0) return(as.Date(seq(d, by = "month", length.out = n + 1)[n + 1]))
    as.Date(seq(d, by = paste0(n, " months"), length.out = 2)[2])
  }
  
  parse_lfs_end <- function(label) {
    x <- trimws(as.character(label))
    month_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)
    months_found <- regmatches(x, gregexpr("[A-Za-z]{3}", x))[[1]]
    year_found <- regmatches(x, gregexpr("[0-9]{4}", x))[[1]]
    if (length(months_found) >= 2 && length(year_found) >= 1) {
      end_month <- month_map[tolower(months_found[2])]
      yr <- as.integer(year_found[1])
      if (!is.na(end_month) && !is.na(yr)) return(as.Date(sprintf("%04d-%02d-01", yr, end_month)))
    }
    as.Date(NA)
  }
  
  manual_month_from_date <- function(d) {
    tolower(paste0(format(d, "%b"), format(d, "%Y")))
  }
  
  manual_month_to_display <- function(mm) {
    #  like "dec2025" -> "december 2025"
    mm <- tolower(gsub("[[:space:]]+", "", as.character(mm)))
    mon3 <- substr(gsub("[^a-z]", "", mm), 1, 3)
    yr <- as.integer(substr(gsub("[^0-9]", "", mm), 1, 4))
    m <- match(mon3, tolower(month.abb))
    if (is.na(m) || is.na(yr)) return(mm)
    format(as.Date(sprintf("%04d-%02d-01", yr, m)), "%B %Y")
  }
  
  mode_from_choice <- function(choice, labs) {
    if (!is.null(labs$latest) && identical(choice, labs$latest)) "latest" else "aligned"
  }
  
  # Parse a manual dropdown label ("Mmm-Mmm YYYY") to its end-month Date, or NULL
  .parse_period_end <- function(label) {
    if (is.null(label) || !nzchar(label)) return(NULL)
    d <- parse_lfs_end(label)
    if (is.na(d)) return(NULL)
    d
  }
  
  # Update reference month  from auto-detected value
  .update_ref_month <- function(detected_mm) {
    old_mm <- reference_manual_month()
    reference_manual_month(detected_mm)
    if (!is.null(old_mm) && nzchar(old_mm) && !identical(tolower(old_mm), tolower(detected_mm))) {
      showNotification(
        paste0("Reference month auto-detected: ", manual_month_to_display(detected_mm),
               " (from uploaded A01)"),
        type = "message", duration = 8
      )
    }
  }
  
  # auto detect ref month + dropdownn
  session$onFlushed(function() {
    
    showModal(modalDialog(
      div(
        div(class = "loader"),
        strong("Loading…"),
        div(style = "margin-top: 8px; color: #505a5f;", "Detecting latest reference month and periods")
      ),
      footer = NULL, easyClose = FALSE
    ))
    
    mm <- NULL
    
    # 1) try latest lfs period
    if (requireNamespace("DBI", quietly = TRUE) && requireNamespace("RPostgres", quietly = TRUE)) {
      conn <- NULL
      tryCatch({
        conn <- DBI::dbConnect(RPostgres::Postgres())
        res <- DBI::dbGetQuery(conn, 'SELECT DISTINCT time_period FROM "ons"."labour_market__age_group"')
        if (nrow(res) > 0) {
          ends <- as.Date(vapply(res$time_period, parse_lfs_end, as.Date(NA)), origin = "1970-01-01")
          if (any(!is.na(ends))) {
            end_latest <- max(ends, na.rm = TRUE)
            mm_date <- add_months(end_latest, 2)
            mm <- manual_month_from_date(mm_date)
          }
        }
      }, error = function(e) NULL, finally = {
        if (!is.null(conn)) try(DBI::dbDisconnect(conn), silent = TRUE)
      })
    }
    
    # 2) fallback
    if (is.null(mm) && file.exists(config_path)) {
      env <- new.env()
      tryCatch({
        source(config_path, local = env)
        if (exists("manual_month", envir = env)) mm <- tolower(env$manual_month)
      }, error = function(e) NULL)
    }
    
    if (is.null(mm) || !nzchar(mm)) {
      mm <- manual_month_from_date(Sys.Date())
    }
    
    reference_manual_month(mm)
    
    #  dashboard quarter end (manual_month - 2 months)
    # manual_month is always the 1st of month
    mm_mon3 <- substr(gsub("[^a-z]", "", mm), 1, 3)
    mm_yr <- as.integer(substr(gsub("[^0-9]", "", mm), 1, 4))
    mm_m <- match(mm_mon3, tolower(month.abb))
    mm_date <- as.Date(sprintf("%04d-%02d-01", mm_yr, mm_m))
    lfs_end <- add_months(mm_date, -2)
    
    # vacancies labels
    vac_lab_aligned <- ""
    vac_lab_latest <- ""
    if (requireNamespace("DBI", quietly = TRUE) && requireNamespace("RPostgres", quietly = TRUE)) {
      conn <- NULL
      tryCatch({
        conn <- DBI::dbConnect(RPostgres::Postgres())
        res <- DBI::dbGetQuery(conn, 'SELECT DISTINCT time_period FROM "ons"."labour_market__vacancies_business"')
        if (nrow(res) > 0) {
          ends <- as.Date(vapply(res$time_period, parse_lfs_end, as.Date(NA)), origin = "1970-01-01")
          ok <- !is.na(ends)
          if (any(ok)) {
            end_latest <- max(ends[ok], na.rm = TRUE)
            end_aligned_candidates <- ends[ok & ends <= lfs_end]
            end_aligned <- if (length(end_aligned_candidates) >= 1) max(end_aligned_candidates) else end_latest
            
            # recreate labels (ensure they exist in db format)
            make_lfs_label_local <- function(end_date) {
              start_date <- add_months(end_date, -2)
              paste0(format(start_date, "%b"), "-", format(end_date, "%b"), " ", format(end_date, "%Y"))
            }
            vac_lab_aligned <- make_lfs_label_local(end_aligned)
            vac_lab_latest  <- make_lfs_label_local(end_latest)
          }
        }
      }, error = function(e) NULL, finally = {
        if (!is.null(conn)) try(DBI::dbDisconnect(conn), silent = TRUE)
      })
    }
    
    # payroll labels (3-month window)
    pay_lab_aligned <- ""
    pay_lab_latest <- ""
    if (requireNamespace("DBI", quietly = TRUE) && requireNamespace("RPostgres", quietly = TRUE)) {
      conn <- NULL
      tryCatch({
        conn <- DBI::dbConnect(RPostgres::Postgres())
        res <- DBI::dbGetQuery(conn, 'SELECT DISTINCT time_period FROM "ons"."labour_market__payrolled_employees"')
        if (nrow(res) > 0) {
          months <- suppressWarnings(as.Date(paste0("01 ", res$time_period), format = "%d %B %Y"))
          ok <- !is.na(months)
          if (any(ok)) {
            end_latest <- max(months[ok], na.rm = TRUE)
            end_aligned_candidates <- months[ok & months <= lfs_end]
            end_aligned <- if (length(end_aligned_candidates) >= 1) max(end_aligned_candidates) else end_latest
            
            make_lfs_label_local <- function(end_date) {
              start_date <- add_months(end_date, -2)
              paste0(format(start_date, "%b"), "-", format(end_date, "%b"), " ", format(end_date, "%Y"))
            }
            pay_lab_aligned <- make_lfs_label_local(end_aligned)
            pay_lab_latest  <- make_lfs_label_local(end_latest)
          }
        }
      }, error = function(e) NULL, finally = {
        if (!is.null(conn)) try(DBI::dbDisconnect(conn), silent = TRUE)
      })
    }
    
    # store + update dropdowns 
    period_labels(list(
      vac = list(aligned = vac_lab_aligned, latest = vac_lab_latest),
      payroll = list(aligned = pay_lab_aligned, latest = pay_lab_latest)
    ))
    
    if (nzchar(vac_lab_aligned) && nzchar(vac_lab_latest)) {
      vac_choices <- setNames(
        c(vac_lab_aligned, vac_lab_latest),
        c(paste0(vac_lab_aligned, " (default)"), vac_lab_latest)
      )
      updateSelectInput(session, "vacancies_period",
                        choices = vac_choices,
                        selected = vac_lab_aligned)
    }
    if (nzchar(pay_lab_aligned) && nzchar(pay_lab_latest)) {
      pay_choices <- setNames(
        c(pay_lab_aligned, pay_lab_latest),
        c(paste0(pay_lab_aligned, " (default)"), pay_lab_latest)
      )
      updateSelectInput(session, "payroll_period",
                        choices = pay_choices,
                        selected = pay_lab_aligned)
    }
    
    removeModal()
  }, once = TRUE)
  
  # reference month display
  output$month_status <- renderUI({
    mm <- reference_manual_month()
    if (is.null(mm) || !nzchar(mm)) {
      return(div(style = "margin-top: 10px;", div(class = "loader")))
    }
    div(style = "margin-top: 10px; font-weight: 600;", manual_month_to_display(mm))
  })
  
  
  # preview: dashboard
  
  observeEvent(input$preview_dashboard, {
    
    withProgress(message = "Loading Dashboard Data", value = 0, {
      
      incProgress(0.1, detail = "Step 1/6: Checking configuration files...")
      Sys.sleep(0.3)
      
      if (!file.exists(calculations_path)) {
        showNotification("Error: calculations.R not found", type = "error")
        return()
      }
      
      incProgress(0.15, detail = "Step 2/6: Loading configuration...")
      Sys.sleep(0.2)
      
      calc_env <- new.env(parent = globalenv())
      
      if (file.exists(config_path)) {
        source(config_path, local = calc_env)
      }
      
      incProgress(0.15, detail = "Step 3/6: Setting reference month...")
      Sys.sleep(0.2)
      
      mm <- reference_manual_month()
      if (!is.null(mm) && nzchar(mm)) {
        calc_env$manual_month <- tolower(mm)
      }
      
      # vacancies & payroll choices
      labs <- period_labels()
      calc_env$vacancies_mode <- mode_from_choice(input$vacancies_period, labs$vac)
      calc_env$payroll_mode <- mode_from_choice(input$payroll_period, labs$payroll)
      
      incProgress(0.2, detail = "Step 4/6: Running calculations...")
      
      tryCatch({
        source(calculations_path, local = calc_env)
      }, error = function(e) {
        showNotification(paste("Calculation error:", e$message), type = "error", duration = 5)
        return()
      })
      
      incProgress(0.2, detail = "Step 5/6: Building metrics table...")
      Sys.sleep(0.2)
      
      gv <- function(name) {
        if (exists(name, envir = calc_env)) {
          val <- get(name, envir = calc_env)
          if (is.numeric(val)) return(val)
        }
        NA_real_
      }
      
      metrics <- list(
        list(name = "Employment 16+ (000s)", cur = gv("emp16_cur") / 1000, dq = gv("emp16_dq") / 1000, dy = gv("emp16_dy") / 1000, dc = gv("emp16_dc") / 1000, de = gv("emp16_de") / 1000, invert = FALSE, type = "count"),
        list(name = "Employment rate 16-64 (%)", cur = gv("emp_rt_cur"), dq = gv("emp_rt_dq"), dy = gv("emp_rt_dy"), dc = gv("emp_rt_dc"), de = gv("emp_rt_de"), invert = FALSE, type = "rate"),
        list(name = "Unemployment 16+ (000s)", cur = gv("unemp16_cur") / 1000, dq = gv("unemp16_dq") / 1000, dy = gv("unemp16_dy") / 1000, dc = gv("unemp16_dc") / 1000, de = gv("unemp16_de") / 1000, invert = TRUE, type = "count"),
        list(name = "Unemployment rate 16+ (%)", cur = gv("unemp_rt_cur"), dq = gv("unemp_rt_dq"), dy = gv("unemp_rt_dy"), dc = gv("unemp_rt_dc"), de = gv("unemp_rt_de"), invert = TRUE, type = "rate"),
        list(name = "Inactivity 16-64 (000s)", cur = gv("inact_cur") / 1000, dq = gv("inact_dq") / 1000, dy = gv("inact_dy") / 1000, dc = gv("inact_dc") / 1000, de = gv("inact_de") / 1000, invert = TRUE, type = "count"),
        list(name = "Inactivity 50-64 (000s)", cur = gv("inact5064_cur") / 1000, dq = gv("inact5064_dq") / 1000, dy = gv("inact5064_dy") / 1000, dc = gv("inact5064_dc") / 1000, de = gv("inact5064_de") / 1000, invert = TRUE, type = "count"),
        list(name = "Inactivity rate 16-64 (%)", cur = gv("inact_rt_cur"), dq = gv("inact_rt_dq"), dy = gv("inact_rt_dy"), dc = gv("inact_rt_dc"), de = gv("inact_rt_de"), invert = TRUE, type = "rate"),
        list(name = "Inactivity rate 50-64 (%)", cur = gv("inact5064_rt_cur"), dq = gv("inact5064_rt_dq"), dy = gv("inact5064_rt_dy"), dc = gv("inact5064_rt_dc"), de = gv("inact5064_rt_de"), invert = TRUE, type = "rate"),
        list(name = "Vacancies (000s)", cur = gv("vac_cur"), dq = gv("vac_dq"), dy = gv("vac_dy"), dc = gv("vac_dc"), de = gv("vac_de"), invert = NA, type = "exempt"),
        list(name = "Payroll employees (000s)", cur = gv("payroll_cur"), dq = gv("payroll_dq"), dy = gv("payroll_dy"), dc = gv("payroll_dc"), de = gv("payroll_de"), invert = FALSE, type = "exempt"),
        list(name = "Wages total pay (%)", cur = gv("latest_wages"), dq = gv("wages_change_q"), dy = gv("wages_change_y"), dc = gv("wages_change_covid"), de = gv("wages_change_election"), invert = FALSE, type = "wages"),
        list(name = "Wages CPI-adjusted (%)", cur = gv("latest_wages_cpi"), dq = gv("wages_cpi_change_q"), dy = gv("wages_cpi_change_y"), dc = gv("wages_cpi_change_covid"), de = gv("wages_cpi_change_election"), invert = FALSE, type = "wages")
      )
      
      incProgress(0.2, detail = "Step 6/6: Finalizing dashboard...")
      Sys.sleep(0.2)
      
      dashboard_data(metrics)
    })
    
    showNotification("Dashboard loaded successfully!", type = "message", duration = 3)
  })
  
  # preview: top ten
  
  observeEvent(input$preview_topten, {
    
    withProgress(message = "Loading Top Ten Statistics", value = 0, {
      
      incProgress(0.1, detail = "Step 1/6: Checking required files...")
      Sys.sleep(0.3)
      
      if (!file.exists(calculations_path)) {
        showNotification("Error: calculations.R not found", type = "error")
        return()
      }
      
      if (!file.exists(top_ten_path)) {
        showNotification("Error: top_ten_stats.R not found", type = "error")
        return()
      }
      
      incProgress(0.15, detail = "Step 2/6: Loading configuration...")
      Sys.sleep(0.2)
      
      if (file.exists(config_path)) {
        source(config_path, local = FALSE)
      }
      
      incProgress(0.15, detail = "Step 3/6: Setting reference month...")
      Sys.sleep(0.2)
      
      mm <- reference_manual_month()
      if (!is.null(mm) && nzchar(mm)) {
        manual_month <<- tolower(mm)
      }
      
      # vacancies & payroll choices
      labs <- period_labels()
      vacancies_mode <<- mode_from_choice(input$vacancies_period, labs$vac)
      payroll_mode <<- mode_from_choice(input$payroll_period, labs$payroll)
      
      incProgress(0.2, detail = "Step 4/6: Running calculations...")
      
      tryCatch({
        source(calculations_path, local = FALSE)
      }, error = function(e) {
        showNotification(paste("Calculation error:", e$message), type = "error", duration = 5)
        return()
      })
      
      incProgress(0.2, detail = "Step 5/6: Loading top ten generator...")
      
      source(top_ten_path, local = FALSE)
      
      incProgress(0.2, detail = "Step 6/6: Generating statistics...")
      
      if (exists("generate_top_ten")) {
        top10 <- tryCatch(generate_top_ten(), error = function(e) {
          showNotification(paste("Top ten generation error:", e$message), type = "error")
          NULL
        })
        if (!is.null(top10)) topten_data(top10)
      } else {
        showNotification("Error: generate_top_ten function not found", type = "error")
        return()
      }
    })
    
    showNotification("Top Ten statistics loaded successfully!", type = "message", duration = 3)
  })
  
  # download: word
  
  output$download_word <- downloadHandler(
    filename = function() {
      mm <- reference_manual_month()
      label <- if (!is.null(mm) && nzchar(mm)) manual_month_to_display(mm) else format(Sys.Date(), "%B %Y")
      paste0("Labour Market Stats Briefing - ", label, ".docx")
    },
    content = function(file) {
      
      tryCatch({
        withProgress(message = "Generating Word Document", value = 0, {
          
          incProgress(0.15, detail = "Step 1/6: Checking officer package...")
          Sys.sleep(0.2)
          
          if (!requireNamespace("officer", quietly = TRUE)) {
            stop("officer package not installed")
          }
          
          incProgress(0.15, detail = "Step 2/6: Locating template file...")
          Sys.sleep(0.2)
          
          if (!file.exists(template_path)) {
            incProgress(0.7, detail = "Creating basic document (no template)...")
            
            doc <- officer::read_docx()
            doc <- officer::body_add_par(doc, "Labour Market Statistics Brief", style = "heading 1")
            doc <- officer::body_add_par(doc, paste("Generated:", format(Sys.Date(), "%d %B %Y")))
            doc <- officer::body_add_par(doc, "Note: Template file (utils/DB.docx) not found.")
            print(doc, target = file)
            
            showNotification("Word document created (basic - no template)", type = "warning", duration = 3)
            return()
          }
          
          incProgress(0.2, detail = "Step 3/6: Loading word output script...")
          
          source(word_script_path, local = FALSE)
          
          mm <- reference_manual_month()
          labs <- period_labels()
          vac_mode <- mode_from_choice(input$vacancies_period, labs$vac)
          pay_mode <- mode_from_choice(input$payroll_period, labs$payroll)
          month_override <- mm
          
          incProgress(0.2, detail = "Step 4/6: Running calculations...")
          incProgress(0.15, detail = "Step 5/6: Generating document content...")
          incProgress(0.15, detail = "Step 6/6: Writing Word file...")
          
          generate_word_output(
            template_path = template_path,
            output_path = file,
            calculations_path = calculations_path,
            config_path = config_path,
            summary_path = summary_path,
            top_ten_path = top_ten_path,
            manual_month_override = month_override,
            vacancies_mode_override = vac_mode,
            payroll_mode_override = pay_mode
          )
        })
        
        showNotification("Word document generated!", type = "message", duration = 3)
        
      }, error = function(e) {
        message("Word download error: ", e$message)
        showNotification(paste("Word error:", e$message), type = "error", duration = 5)
        
        # write a fallback docx so the download doesn't fail as .htm
        if (requireNamespace("officer", quietly = TRUE)) {
          doc <- officer::read_docx()
          doc <- officer::body_add_par(doc, "Error Generating Brief", style = "heading 1")
          doc <- officer::body_add_par(doc, paste("Error:", e$message))
          doc <- officer::body_add_par(doc, "Please check the R console for details.")
          print(doc, target = file)
        } else {
          writeLines(paste("Error:", e$message), con = file)
        }
      })
    }
  )
  
  
  output$download_excel <- downloadHandler(
    filename = function() {
      mm <- reference_manual_month()
      label <- if (!is.null(mm) && nzchar(mm)) manual_month_to_display(mm) else format(Sys.Date(), "%B %Y")
      paste0("LM Stats Audit - ", label, ".xlsx")
    },
    content = function(file) {
      # :
      tryCatch({
        withProgress(message = "Generating Excel Workbook", value = 0, {
          
          incProgress(0.1, detail = "Step 1/4: Checking openxlsx...")
          if (!requireNamespace("openxlsx", quietly = TRUE)) {
            stop("openxlsx package not installed")
          }
          
          incProgress(0.2, detail = "Step 2/4: Loading excel_audit_workbook.R...")
          excel_env <- new.env(parent = globalenv())
          source(excel_script_path, local = excel_env)
          
          if (!exists("create_audit_workbook", envir = excel_env)) {
            stop("create_audit_workbook() not found after sourcing excel_audit_workbook.R")
          }
          
          incProgress(0.5, detail = "Step 3/4: Building workbook (this may take a moment)...")
          mm <- reference_manual_month()
          labs <- period_labels()
          vac_mode <- mode_from_choice(input$vacancies_period, labs$vac)
          pay_mode <- mode_from_choice(input$payroll_period, labs$payroll)
          month_override <- mm
          tmp_xlsx <- tempfile(fileext = ".xlsx")
          excel_env$create_audit_workbook(
            output_path = tmp_xlsx,
            file_a01 = uploaded_files$a01,
            file_hr1 = uploaded_files$hr1,
            file_x09 = uploaded_files$x09,
            file_rtisa = uploaded_files$rtisa,
            file_cla01 = uploaded_files$cla01,
            file_x02 = uploaded_files$x02,
            file_oecd_unemp = uploaded_files$oecd_unemp,
            file_oecd_emp = uploaded_files$oecd_emp,
            file_oecd_inact = uploaded_files$oecd_inact,
            calculations_path = calculations_path,
            config_path = config_path,
            vacancies_mode = vac_mode,
            payroll_mode = pay_mode,
            manual_month_override = month_override,
            verbose = FALSE
          )
          
          incProgress(0.15, detail = "Step 4/4: Preparing download...")
          ok <- file.copy(tmp_xlsx, file, overwrite = TRUE)
          if (!isTRUE(ok) || !file.exists(file)) {
            stop("Excel workbook could not be copied to Shiny download location")
          }
        })
        
        showNotification("Excel workbook generated", type = "message", duration = 3)
        
      }, error = function(e) {
        # on error, still return a valid .xlsx to the user.
        message("Excel download error: ", e$message)
        showNotification(paste("Excel error:", e$message), type = "error", duration = 5)
        
        if (requireNamespace("openxlsx", quietly = TRUE)) {
          tmp_xlsx <- tempfile(fileext = ".xlsx")
          wb <- openxlsx::createWorkbook()
          openxlsx::addWorksheet(wb, "Error")
          openxlsx::writeData(wb, "Error", data.frame(
            Error = paste("Failed to generate workbook:", e$message),
            Suggestion = "Check database connectivity / package availability, then try again."
          ))
          openxlsx::saveWorkbook(wb, tmp_xlsx, overwrite = TRUE)
          file.copy(tmp_xlsx, file, overwrite = TRUE)
        } else {
          # worst-case: write a plain-text error so the download isn't empty
          writeLines(paste("Failed to generate Excel workbook:", e$message), con = file)
        }
      })
    }
  )
  

  # MANUAL  

  
  .check_manual_ready <- function() {
    if (is.null(uploaded_files$a01)) {
      showNotification("Upload at least the A01 file", type = "warning", duration = 5)
      return(FALSE)
    }
    missing <- c()
    if (is.null(uploaded_files$hr1))   missing <- c(missing, "HR1 (redundancy notifications)")
    if (is.null(uploaded_files$x09))   missing <- c(missing, "X09 (real wages/CPI)")
    if (is.null(uploaded_files$rtisa)) missing <- c(missing, "RTISA (payroll/sector)")
    if (length(missing) > 0) {
      showNotification(
        paste0("Proceeding without: ", paste(missing, collapse = ", "),
               ". Those sections will show \u2014."),
        type = "message", duration = 8
      )
    }
    TRUE
  }
  
  # manual preview: dashboard
  observeEvent(input$manual_preview_dashboard, {
    if (!.check_manual_ready()) return()
    
    withProgress(message = "Loading Dashboard (Manual Upload)", value = 0, {
      
      incProgress(0.1, detail = "Reading Excel files...")
      
      calc_env <- new.env(parent = globalenv())
      
      incProgress(0.3, detail = "Running calculations...")
      
      source(excel_calc_path, local = calc_env)
      source("utils/helpers.R", local = calc_env)
      detected_mm <- calc_env$run_calculations_from_excel(
        manual_month = NULL,
        file_a01 = uploaded_files$a01, file_hr1 = uploaded_files$hr1,
        file_x09 = uploaded_files$x09, file_rtisa = uploaded_files$rtisa,
        vac_end_override = .parse_period_end(input$manual_vacancies_period),
        payroll_end_override = .parse_period_end(input$manual_payroll_period),
        target_env = calc_env
      )
      .update_ref_month(detected_mm)
      
      incProgress(0.4, detail = "Building metrics table...")
      
      gv <- function(name) {
        if (exists(name, envir = calc_env)) {
          val <- get(name, envir = calc_env)
          if (is.numeric(val)) return(val)
        }
        NA_real_
      }
      
      metrics <- list(
        list(name = "Employment 16+ (000s)", cur = gv("emp16_cur") / 1000, dq = gv("emp16_dq") / 1000, dy = gv("emp16_dy") / 1000, dc = gv("emp16_dc") / 1000, de = gv("emp16_de") / 1000, invert = FALSE, type = "count"),
        list(name = "Employment rate 16-64 (%)", cur = gv("emp_rt_cur"), dq = gv("emp_rt_dq"), dy = gv("emp_rt_dy"), dc = gv("emp_rt_dc"), de = gv("emp_rt_de"), invert = FALSE, type = "rate"),
        list(name = "Unemployment 16+ (000s)", cur = gv("unemp16_cur") / 1000, dq = gv("unemp16_dq") / 1000, dy = gv("unemp16_dy") / 1000, dc = gv("unemp16_dc") / 1000, de = gv("unemp16_de") / 1000, invert = TRUE, type = "count"),
        list(name = "Unemployment rate 16+ (%)", cur = gv("unemp_rt_cur"), dq = gv("unemp_rt_dq"), dy = gv("unemp_rt_dy"), dc = gv("unemp_rt_dc"), de = gv("unemp_rt_de"), invert = TRUE, type = "rate"),
        list(name = "Inactivity 16-64 (000s)", cur = gv("inact_cur") / 1000, dq = gv("inact_dq") / 1000, dy = gv("inact_dy") / 1000, dc = gv("inact_dc") / 1000, de = gv("inact_de") / 1000, invert = TRUE, type = "count"),
        list(name = "Inactivity 50-64 (000s)", cur = gv("inact5064_cur") / 1000, dq = gv("inact5064_dq") / 1000, dy = gv("inact5064_dy") / 1000, dc = gv("inact5064_dc") / 1000, de = gv("inact5064_de") / 1000, invert = TRUE, type = "count"),
        list(name = "Inactivity rate 16-64 (%)", cur = gv("inact_rt_cur"), dq = gv("inact_rt_dq"), dy = gv("inact_rt_dy"), dc = gv("inact_rt_dc"), de = gv("inact_rt_de"), invert = TRUE, type = "rate"),
        list(name = "Inactivity rate 50-64 (%)", cur = gv("inact5064_rt_cur"), dq = gv("inact5064_rt_dq"), dy = gv("inact5064_rt_dy"), dc = gv("inact5064_rt_dc"), de = gv("inact5064_rt_de"), invert = TRUE, type = "rate"),
        list(name = "Vacancies (000s)", cur = gv("vac_cur"), dq = gv("vac_dq"), dy = gv("vac_dy"), dc = gv("vac_dc"), de = gv("vac_de"), invert = NA, type = "exempt"),
        list(name = "Payroll employees (000s)", cur = gv("payroll_cur"), dq = gv("payroll_dq"), dy = gv("payroll_dy"), dc = gv("payroll_dc"), de = gv("payroll_de"), invert = FALSE, type = "exempt"),
        list(name = "Wages total pay (%)", cur = gv("latest_wages"), dq = gv("wages_change_q"), dy = gv("wages_change_y"), dc = gv("wages_change_covid"), de = gv("wages_change_election"), invert = FALSE, type = "wages"),
        list(name = "Wages CPI-adjusted (%)", cur = gv("latest_wages_cpi"), dq = gv("wages_cpi_change_q"), dy = gv("wages_cpi_change_y"), dc = gv("wages_cpi_change_covid"), de = gv("wages_cpi_change_election"), invert = FALSE, type = "wages")
      )
      
      incProgress(0.2, detail = "Done")
      dashboard_data(metrics)
    })
    
    showNotification("Dashboard loaded (Manual Upload)", type = "message", duration = 3)
  })
  
  # manual preview: top ten
  observeEvent(input$manual_preview_topten, {
    if (!.check_manual_ready()) return()
    
    withProgress(message = "Loading Top Ten (Manual Upload)", value = 0, {
      
      incProgress(0.2, detail = "Running Excel calculations...")
      
      if (file.exists(config_path)) source(config_path, local = FALSE)
      source("utils/helpers.R", local = FALSE)
      
      source(excel_calc_path, local = FALSE)
      detected_mm <- run_calculations_from_excel(
        manual_month = NULL,
        file_a01 = uploaded_files$a01, file_hr1 = uploaded_files$hr1,
        file_x09 = uploaded_files$x09, file_rtisa = uploaded_files$rtisa,
        vac_end_override = .parse_period_end(input$manual_vacancies_period),
        payroll_end_override = .parse_period_end(input$manual_payroll_period),
        target_env = globalenv()
      )
      manual_month <<- detected_mm
      .update_ref_month(detected_mm)
      
      incProgress(0.5, detail = "Generating top ten stats...")
      
      source(top_ten_path, local = FALSE)
      
      if (exists("generate_top_ten")) {
        top10 <- tryCatch(generate_top_ten(), error = function(e) {
          showNotification(paste("Top ten error:", e$message), type = "error")
          NULL
        })
        if (!is.null(top10)) topten_data(top10)
      }
      
      incProgress(0.3, detail = "Done")
    })
    
    showNotification("Top Ten loaded (Manual Upload)", type = "message", duration = 3)
  })
  
  # manual download: word
  output$manual_download_word <- downloadHandler(
    filename = function() {
      mm <- reference_manual_month()
      label <- if (!is.null(mm) && nzchar(mm)) manual_month_to_display(mm) else format(Sys.Date(), "%B %Y")
      paste0("Labour Market Stats Briefing - ", label, ".docx")
    },
    content = function(file) {
      if (is.null(uploaded_files$a01)) {
        showNotification("Upload at least the A01 file first", type = "warning")
        doc <- officer::read_docx()
        doc <- officer::body_add_par(doc, "Upload the A01 file to generate.", style = "heading 1")
        print(doc, target = file)
        return()
      }
      
      tryCatch({
        withProgress(message = "Generating Word (Manual Upload)", value = 0, {
          
          incProgress(0.2, detail = "Loading scripts...")
          
          source("utils/helpers.R", local = FALSE)
          source(excel_calc_path, local = FALSE)
          if (file.exists(config_path)) source(config_path, local = FALSE)
          
          incProgress(0.3, detail = "Running calculations...")
          
          detected_mm <- run_calculations_from_excel(
            manual_month = NULL,
            file_a01 = uploaded_files$a01, file_hr1 = uploaded_files$hr1,
            file_x09 = uploaded_files$x09, file_rtisa = uploaded_files$rtisa,
            vac_end_override = .parse_period_end(input$manual_vacancies_period),
            payroll_end_override = .parse_period_end(input$manual_payroll_period),
            target_env = globalenv()
          )
          manual_month <<- detected_mm
          .update_ref_month(detected_mm)
          
          incProgress(0.2, detail = "Generating summary & top ten...")
          
          source(summary_path, local = FALSE)
          source(top_ten_path, local = FALSE)
          
          fallback_lines <- function() {
            stats <- list()
            for (i in 1:10) stats[[paste0("line", i)]] <- "(Data unavailable)"
            stats
          }
          summary_lines <- tryCatch(generate_summary(), error = function(e) fallback_lines())
          top10_lines <- tryCatch(generate_top_ten(), error = function(e) fallback_lines())
          
          incProgress(0.2, detail = "Writing Word file...")
          
          # use ManualDB.docx template with qvz placeholders
          manual_template <- "utils/ManualDB.docx"
          if (file.exists(manual_template)) {
            source("utils/manual_word_output.R", local = FALSE)
            generate_manual_word_output(
              manual_month = manual_month,
              file_a01 = uploaded_files$a01, file_hr1 = uploaded_files$hr1,
              file_x09 = uploaded_files$x09, file_rtisa = uploaded_files$rtisa,
              file_oecd_unemp = uploaded_files$oecd_unemp,
              file_oecd_emp   = uploaded_files$oecd_emp,
              file_oecd_inact = uploaded_files$oecd_inact,
              vac_end_override = .parse_period_end(input$manual_vacancies_period),
              payroll_end_override = .parse_period_end(input$manual_payroll_period),
              template_path = manual_template,
              output_path = file,
              verbose = FALSE
            )
          } else if (file.exists(template_path)) {
            # fallback to DB.docx with old-style placeholders
            source(word_script_path, local = FALSE)
            doc <- officer::read_docx(template_path)
            title_label <- if (exists("manual_month", inherits = TRUE)) manual_month_to_label(manual_month) else ""
            doc <- replace_all(doc, "Z1", title_label)
            if (exists("lfs_period_label", inherits = TRUE)) doc <- replace_all(doc, "LFS_PERIOD_LABEL", lfs_period_label)
            if (exists("lfs_period_short_label", inherits = TRUE)) doc <- replace_all(doc, "LFS_QUARTER_LABEL", lfs_period_short_label)
            if (exists("vacancies_period_short_label", inherits = TRUE)) doc <- replace_all(doc, "VACANCIES_QUARTER_LABEL", vacancies_period_short_label)
            for (i in 10:1) doc <- replace_all(doc, paste0("sl", i), summary_lines[[paste0("line", i)]])
            for (i in 10:1) doc <- replace_all(doc, paste0("tt", i), top10_lines[[paste0("line", i)]])
            doc <- replace_all(doc, "RENDER_DATE", format(Sys.Date(), "%d %B %Y"))
            print(doc, target = file)
          } else {
            stop("No Word template found (ManualDB.docx or DB.docx)")
          }
          
          incProgress(0.1, detail = "Done")
        })
        
        showNotification("Word document generated (Manual Upload)", type = "message", duration = 3)
        
      }, error = function(e) {
        message("Manual Word error: ", e$message)
        showNotification(
          paste("Word generation failed:", e$message),
          type = "error", duration = 10
        )
        # so the browser doesn't hang, but the notication tells the user what happened
        writeLines(paste("Generation failed:", e$message), con = file)
      })
    }
  )
  
  # manual download: excel
  output$manual_download_excel <- downloadHandler(
    filename = function() {
      mm <- reference_manual_month()
      label <- if (!is.null(mm) && nzchar(mm)) manual_month_to_display(mm) else format(Sys.Date(), "%B %Y")
      paste0("LM Stats Audit - ", label, ".xlsx")
    },
    content = function(file) {
      if (is.null(uploaded_files$a01)) {
        showNotification("Upload at least the A01 file first", type = "warning")
        if (requireNamespace("openxlsx", quietly = TRUE)) {
          wb <- openxlsx::createWorkbook()
          openxlsx::addWorksheet(wb, "Error")
          openxlsx::writeData(wb, "Error", data.frame(Error = "Upload the A01 file to generate."))
          openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
        }
        return()
      }
      
      tryCatch({
        withProgress(message = "Generating Excel (Manual Upload)", value = 0, {
          incProgress(0.2, detail = "Loading scripts...")
          source("utils/helpers.R", local = FALSE)
          source(excel_calc_path, local = FALSE)
          if (file.exists(config_path)) source(config_path, local = FALSE)
          
          incProgress(0.4, detail = "Running calculations...")
          detected_mm <- run_calculations_from_excel(
            manual_month = NULL,
            file_a01 = uploaded_files$a01, file_hr1 = uploaded_files$hr1,
            file_x09 = uploaded_files$x09, file_rtisa = uploaded_files$rtisa,
            vac_end_override = .parse_period_end(input$manual_vacancies_period),
            payroll_end_override = .parse_period_end(input$manual_payroll_period),
            target_env = globalenv()
          )
          manual_month <<- detected_mm
          .update_ref_month(detected_mm)
          
          incProgress(0.3, detail = "Building workbook...")
          excel_env <- new.env(parent = globalenv())
          source(excel_script_path, local = excel_env)
          tmp_xlsx <- tempfile(fileext = ".xlsx")
          excel_env$create_audit_workbook(
            output_path = tmp_xlsx,
            file_a01 = uploaded_files$a01,
            file_hr1 = uploaded_files$hr1,
            file_x09 = uploaded_files$x09,
            file_rtisa = uploaded_files$rtisa,
            file_cla01 = uploaded_files$cla01,
            file_x02 = uploaded_files$x02,
            file_oecd_unemp = uploaded_files$oecd_unemp,
            file_oecd_emp = uploaded_files$oecd_emp,
            file_oecd_inact = uploaded_files$oecd_inact,
            calculations_path = calculations_path,
            config_path = config_path,
            verbose = FALSE
          )
          
          incProgress(0.1, detail = "Done")
          file.copy(tmp_xlsx, file, overwrite = TRUE)
        })
        showNotification("Excel workbook generated (manual mode)", type = "message", duration = 3)
      }, error = function(e) {
        message("Manual Excel error: ", e$message)
        showNotification(paste("Excel error:", e$message), type = "error", duration = 5)
        if (requireNamespace("openxlsx", quietly = TRUE)) {
          wb <- openxlsx::createWorkbook()
          openxlsx::addWorksheet(wb, "Error")
          openxlsx::writeData(wb, "Error", data.frame(Error = paste("Failed:", e$message)))
          openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
        }
      })
    }
  )
  
  # render: dashboard preview
  
  output$dashboard_preview <- renderUI({
    metrics <- dashboard_data()
    
    if (is.null(metrics)) {
      return(div(
        p(class = "govuk-body", "Click 'Preview Dashboard' to load statistics."),
        tags$ul(class = "govuk-list",
                tags$li("Employment and unemployment figures"),
                tags$li("Inactivity rates"),
                tags$li("Vacancies and payroll data"),
                tags$li("Wage statistics")
        )
      ))
    }
    
    format_change <- function(val, invert, type) {
      val <- suppressWarnings(as.numeric(gsub("^\\+", "", as.character(val))))
      if (is.na(val)) return(tags$span(class = "stat-neutral", "-"))
      
      css_class <- if (is.na(invert)) "stat-neutral"
      else if (val > 0) { if (invert) "stat-negative" else "stat-positive" }
      else if (val < 0) { if (invert) "stat-positive" else "stat-negative" }
      else "stat-neutral"
      
      sign_str <- if (val > 0) "+" else if (val < 0) "-" else ""
      abs_val <- abs(val)
      
      formatted <- if (type == "rate") paste0(sign_str, round(abs_val, 1), "pp")
      else if (type == "wages") paste0(sign_str, "£", format(round(abs_val), big.mark = ","))
      else paste0(sign_str, format(round(abs_val), big.mark = ","))
      
      tags$span(class = css_class, formatted)
    }
    
    format_current <- function(val, type) {
      val <- suppressWarnings(as.numeric(gsub("^\\+", "", as.character(val))))
      if (is.na(val)) return("-")
      if (type == "rate" || type == "wages") paste0(round(val, 1), "%")
      else format(round(val), big.mark = ",")
    }
    
    rows <- lapply(metrics, function(m) {
      tags$tr(
        tags$td(m$name),
        tags$td(format_current(m$cur, m$type)),
        tags$td(format_change(m$dq, m$invert, m$type)),
        tags$td(format_change(m$dy, m$invert, m$type)),
        tags$td(format_change(m$dc, m$invert, m$type)),
        tags$td(format_change(m$de, m$invert, m$type))
      )
    })
    
    tags$table(class = "stats-table",
               tags$thead(tags$tr(
                 tags$th("Metric"), tags$th("Current"), tags$th("vs Qtr"),
                 tags$th("vs Year"), tags$th("vs Covid"), tags$th("vs Election")
               )),
               tags$tbody(rows)
    )
  })
  
  # render: top ten preview
  
  output$topten_preview <- renderUI({
    top10 <- topten_data()
    
    if (is.null(top10)) {
      return(div(
        p(class = "govuk-body", "Click 'Preview Top Ten Stats' to load statistics."),
        tags$ul(class = "govuk-list",
                tags$li("Wage growth (nominal and CPI-adjusted)"),
                tags$li("Employment and unemployment rates"),
                tags$li("Payroll employment"),
                tags$li("Inactivity trends"),
                tags$li("Vacancies and redundancies")
        )
      ))
    }
    
    items <- lapply(1:10, function(i) {
      line_key <- paste0("line", i)
      line_text <- top10[[line_key]]
      if (is.null(line_text) || line_text == "") line_text <- "(Data not available)"
      tags$li(line_text)
    })
    
    tags$ol(class = "top-ten-list", items)
  })
}

# run application


shinyApp(ui = ui, server = server)