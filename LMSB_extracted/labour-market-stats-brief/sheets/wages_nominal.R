
# wages nominal module - a01 sheet 13 (total) + sheet 15 (regular)
# total: ons.labour_market__weekly_earnings_total


# config

WAGES_NOM_CODES <- list(
  WEEKLY_TOTAL = "KAB9",
  YOY_TOTAL    = "KAC3",
  WEEKLY_REG   = "KAI7",
  YOY_REG      = "KAI9",
  # public/private sector
  YOY_TOTAL_PUBLIC  = "KAC9",
  YOY_TOTAL_PRIVATE = "KAC6",
  YOY_REG_PUBLIC    = "KAJ7",
  YOY_REG_PRIVATE   = "KAJ4"
)

# fetch

fetch_wages_total <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  
  tryCatch({
    query <- 'SELECT time_period, dataset_identifier_code, value
              FROM "ons"."labour_market__weekly_earnings_total"'
    result <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("Failed to fetch total wages: ", e$message)
    tibble::tibble(
      time_period = character(),
      dataset_identifier_code = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}

fetch_wages_regular <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  
  tryCatch({
    query <- 'SELECT time_period, dataset_identifier_code, value
FROM \"ons\".\"labour_market__weekly_earnings_regular\"'
    result <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("Failed to fetch regular wages: ", e$message)
    tibble::tibble(
      time_period = character(),
      dataset_identifier_code = character(),
      value = numeric()
    )
  },
  finally = {
    DBI::dbDisconnect(conn)
  })
}

# helpers

get_nom_val <- function(pg_data, date, code) {
  val_by_code(pg_data, code, make_ymd_label(date))
}

get_nom_avg <- function(pg_data, dates, code) {
  vals <- sapply(dates, function(d) get_nom_val(pg_data, d, code))
  if (any(is.na(vals))) return(NA_real_)
  mean(vals)
}

# compute

compute_wages_nominal <- function(pg_total, pg_regular, manual_mm) {
  cm <- parse_manual_month(manual_mm)
  
  anchor_m <- cm %m-% months(2)
  prev_q_anchor <- anchor_m %m-% months(3)
  
  win3 <- seq(anchor_m, by = "-1 month", length.out = 3)
  prev3 <- seq(anchor_m %m-% months(3), by = "-1 month", length.out = 3)
  yago3 <- seq(anchor_m %m-% months(12), by = "-1 month", length.out = 3)
  covid3 <- c(as.Date("2020-02-01"), as.Date("2020-01-01"), as.Date("2019-12-01"))
  election3 <- c(as.Date("2024-06-01"), as.Date("2024-05-01"), as.Date("2024-04-01"))
  
  # total pay yoy % (from total table)
  latest_total <- get_nom_val(pg_total, anchor_m, WAGES_NOM_CODES$YOY_TOTAL)
  
  # public/private sector yoy % (current)
  total_public <- get_nom_val(pg_total, anchor_m, WAGES_NOM_CODES$YOY_TOTAL_PUBLIC)
  total_private <- get_nom_val(pg_total, anchor_m, WAGES_NOM_CODES$YOY_TOTAL_PRIVATE)
  
  # quarterly comparison for total yoy %
  total_prev_q <- get_nom_val(pg_total, prev_q_anchor, WAGES_NOM_CODES$YOY_TOTAL)
  total_qchange <- if (!is.na(latest_total) && !is.na(total_prev_q)) latest_total - total_prev_q else NA_real_
  
  calc_change_total <- function(dates_a, dates_b, code) {
    a <- get_nom_avg(pg_total, dates_a, code)
    b <- get_nom_avg(pg_total, dates_b, code)
    if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
  }
  
  total_dq <- calc_change_total(win3, prev3, WAGES_NOM_CODES$WEEKLY_TOTAL)
  total_dy <- calc_change_total(win3, yago3, WAGES_NOM_CODES$WEEKLY_TOTAL)
  total_dc <- calc_change_total(win3, covid3, WAGES_NOM_CODES$WEEKLY_TOTAL)
  total_de <- calc_change_total(win3, election3, WAGES_NOM_CODES$WEEKLY_TOTAL)
  
  # regular pay yoy % (from regular table)
  latest_reg <- get_nom_val(pg_regular, anchor_m, WAGES_NOM_CODES$YOY_REG)
  
  # public/private sector regular yoy %
  reg_public <- get_nom_val(pg_regular, anchor_m, WAGES_NOM_CODES$YOY_REG_PUBLIC)
  reg_private <- get_nom_val(pg_regular, anchor_m, WAGES_NOM_CODES$YOY_REG_PRIVATE)
  
  # quarterly comparison for regular yoy %
  reg_prev_q <- get_nom_val(pg_regular, prev_q_anchor, WAGES_NOM_CODES$YOY_REG)
  reg_qchange <- if (!is.na(latest_reg) && !is.na(reg_prev_q)) latest_reg - reg_prev_q else NA_real_
  
  calc_change_reg <- function(dates_a, dates_b, code) {
    a <- get_nom_avg(pg_regular, dates_a, code)
    b <- get_nom_avg(pg_regular, dates_b, code)
    if (is.na(a) || is.na(b)) NA_real_ else (a - b) * 52
  }
  
  reg_dq <- calc_change_reg(win3, prev3, WAGES_NOM_CODES$WEEKLY_REG)
  reg_dy <- calc_change_reg(win3, yago3, WAGES_NOM_CODES$WEEKLY_REG)
  reg_dc <- calc_change_reg(win3, covid3, WAGES_NOM_CODES$WEEKLY_REG)
  reg_de <- calc_change_reg(win3, election3, WAGES_NOM_CODES$WEEKLY_REG)
  
  list(
    total = list(
      cur = latest_total, 
      dq = total_dq, 
      dy = total_dy, 
      dc = total_dc, 
      de = total_de,
      public = total_public,
      private = total_private,
      qchange = total_qchange
    ),
    regular = list(
      cur = latest_reg, 
      dq = reg_dq, 
      dy = reg_dy, 
      dc = reg_dc, 
      de = reg_de,
      public = reg_public,
      private = reg_private,
      qchange = reg_qchange
    ),
    anchor = anchor_m
  )
}

# calculate wages nominal

calculate_wages_nominal <- function(manual_mm) {
  pg_total <- fetch_wages_total()
  pg_regular <- fetch_wages_regular()
  compute_wages_nominal(pg_total, pg_regular, manual_mm)
}
