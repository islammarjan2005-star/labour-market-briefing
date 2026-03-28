# payroll module - rtisa payrolled employees
# table: ons.labour_market__payrolled_employees
# time period format: "july 2025"

# config

PAYROLL_UNIT_TYPE <- "Payrolled employees"

# fetch

fetch_payroll <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  
  tryCatch({
    query <- 'SELECT time_period, unit_type, value
              FROM "ons"."labour_market__payrolled_employees"'
    result <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("Failed to fetch payroll data: ", e$message)
    tibble::tibble(
      time_period = character(),
      unit_type = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}

# lookup

val_payroll <- function(pg_data, period_label) {
  if (is.null(pg_data) || nrow(pg_data) == 0) return(NA_real_)
  
  match_row <- pg_data %>%
    filter(
      unit_type == PAYROLL_UNIT_TYPE,
      startsWith(time_period, period_label)
    )
  
  if (nrow(match_row) == 0) return(NA_real_)
  
  suppressWarnings(as.numeric(match_row$value[1]))
}

# helpers

get_payroll_avg <- function(pg_data, dates) {
  vals <- sapply(dates, function(d) val_payroll(pg_data, make_payroll_label(d)))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

# compute

compute_payroll <- function(pg_data, manual_mm, mode = c("latest", "aligned")) {
  
  # find latest 3 months dynamically for dashboard
  
  payroll_data <- pg_data %>%
    filter(unit_type == PAYROLL_UNIT_TYPE) %>%
    mutate(parsed_date = as.Date(paste0("01 ", time_period), format = "%d %B %Y")) %>%
    filter(!is.na(parsed_date)) %>%
    arrange(desc(parsed_date))
  
  if (nrow(payroll_data) < 3) {
    return(list(
      cur = NA_real_, dq = NA_real_, dy = NA_real_, dc = NA_real_, de = NA_real_,
      anchor = NA,
      flash_cur = NA_real_, flash_dy = NA_real_, flash_de = NA_real_, flash_dm = NA_real_,
      flash_anchor = NA
    ))
  }
  
  mode <- match.arg(mode)

  # choose anchor month for the 3-month "quarter" used on the dashboard:
  # - latest: most recent available in the database (current behaviour)
  # - aligned: align to dashboard reference quarter end (manual_mm - 2 months)
  if (mode == "aligned") {
    target_end <- tryCatch(parse_manual_month(manual_mm) %m-% months(2), error = function(e) NA)
    if (!is.na(target_end)) {
      idx_le <- which(payroll_data$parsed_date <= target_end)
      anchor <- if (length(idx_le) >= 1) payroll_data$parsed_date[idx_le[1]] else payroll_data$parsed_date[1]
    } else {
      anchor <- payroll_data$parsed_date[1]
    }
  } else {
    anchor <- payroll_data$parsed_date[1]
  }

  # 3-month window ending at anchor
  win3 <- seq(anchor, by = "-1 month", length.out = 3)
  
  # 3-month windows for comparisons (based on dynamic anchor)
  prev3 <- seq(anchor %m-% months(3), by = "-1 month", length.out = 3)
  yago3 <- seq(anchor %m-% months(12), by = "-1 month", length.out = 3)
  covid3 <- c(as.Date("2020-02-01"), as.Date("2020-01-01"), as.Date("2019-12-01"))
  election3 <- c(as.Date("2024-06-01"), as.Date("2024-05-01"), as.Date("2024-04-01"))
  
  # get 3-month averages for dashboard
  cur <- get_payroll_avg(pg_data, win3)
  val_q <- get_payroll_avg(pg_data, prev3)
  val_y <- get_payroll_avg(pg_data, yago3)
  val_c <- get_payroll_avg(pg_data, covid3)
  val_e <- get_payroll_avg(pg_data, election3)
  
  # calculate changes (in thousands)
  dq <- if (!is.na(cur) && !is.na(val_q)) (cur - val_q) / 1000 else NA_real_
  dy <- if (!is.na(cur) && !is.na(val_y)) (cur - val_y) / 1000 else NA_real_
  dc <- if (!is.na(cur) && !is.na(val_c)) (cur - val_c) / 1000 else NA_real_
  de <- if (!is.na(cur) && !is.na(val_e)) (cur - val_e) / 1000 else NA_real_
  
  # flash estimate - latest single month available
  
  flash_anchor <- payroll_data$parsed_date[1]
  flash_cur <- suppressWarnings(as.numeric(payroll_data$value[1]))
  
  # comparisons for flash (single months, not averages)
  flash_prev_y_label <- make_payroll_label(flash_anchor %m-% months(12))
  flash_election_label <- "June 2024"
  flash_prev_m_label <- make_payroll_label(flash_anchor %m-% months(1))
  
  flash_val_y <- val_payroll(pg_data, flash_prev_y_label)
  flash_val_e <- val_payroll(pg_data, flash_election_label)
  flash_val_m <- val_payroll(pg_data, flash_prev_m_label)
  
  # flash changes in thousands
  flash_dy <- if (!is.na(flash_cur) && !is.na(flash_val_y)) (flash_cur - flash_val_y) / 1000 else NA_real_
  flash_de <- if (!is.na(flash_cur) && !is.na(flash_val_e)) (flash_cur - flash_val_e) / 1000 else NA_real_
  flash_dm <- if (!is.na(flash_cur) && !is.na(flash_val_m)) (flash_cur - flash_val_m) / 1000 else NA_real_
  
  # flash current in millions
  flash_cur_millions <- flash_cur / 1000000
  
  list(
    cur = cur / 1000,  # Dashboard: thousands
    dq = dq,
    dy = dy,
    dc = dc,
    de = de,
    anchor = anchor,
    # flash estimates
    flash_cur = flash_cur_millions,  # Millions for narrative
    flash_dy = flash_dy,
    flash_de = flash_de,
    flash_dm = flash_dm,
    flash_anchor = flash_anchor
  )
}

# calculate payroll

calculate_payroll <- function(manual_mm, mode = c("latest", "aligned")) {
  pg_data <- fetch_payroll()
  compute_payroll(pg_data, manual_mm, mode = mode)
}
