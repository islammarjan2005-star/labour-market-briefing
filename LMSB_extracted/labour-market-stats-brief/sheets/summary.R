# summary module
# produces 10 lines
# requirescalculations.r to have been sourced

suppressPackageStartupMessages({
  library(glue)
  library(dplyr)
  library(lubridate)
})

# ensure helpers are loaded (formatters like fmt_pct, lfs_label_narrative, etc.)
if (!exists("fmt_pct", inherits = TRUE)) {
  source("utils/helpers.R")
}

# small  helpers

fmt_signed_int <- function(x) {
  if (is.na(x)) return("—")
  s <- if (x > 0) "+" else if (x < 0) "-" else ""
  paste0(s, format(round(abs(x), 0), big.mark = ","))
}

# format integer with commas (absolute)
fmt_int_1k <- function(x) {
  if (is.na(x)) return("\u2014")
  format(round(abs(x)), big.mark = ",")
}

# format integer with commas (signed)
fmt_signed_int_1k <- function(x) {
  if (is.na(x)) return("\u2014")
  s <- if (x > 0) "+" else if (x < 0) "-" else ""
  paste0(s, format(round(abs(x)), big.mark = ","))
}

fmt_signed_pp <- function(x) {
  if (is.na(x)) return("—")
  s <- if (x > 0) "+" else if (x < 0) "-" else ""
  v <- abs(x)
  v1 <- round(v, 1)
  if (v1 == 0 && v != 0) {
    v2 <- round(v, 2)
    paste0(s, format(v2, nsmall = 2), " percentage points")
  } else {
    paste0(s, format(v1, nsmall = 1), " percentage points")
  }
}

fmt_pct_unsigned <- function(x) {
  if (is.na(x)) return("—")
  v <- abs(as.numeric(x))
  v1 <- round(v, 1)
  if (v1 == 0 && v != 0) {
    paste0(format(round(v, 2), nsmall = 2), "%")
  } else {
    paste0(format(v1, nsmall = 1), "%")
  }
}

safe_num <- function(x) {
  if (is.null(x) || length(x) == 0) return(NA_real_)
  suppressWarnings(as.numeric(x[1]))
}

pct_from_delta <- function(cur, delta) {
  # cur and delta in the same units
  cur <- safe_num(cur); delta <- safe_num(delta)
  base <- cur - delta
  if (is.na(cur) || is.na(delta) || is.na(base) || base == 0) return(NA_real_)
  (delta / base) * 100
}

# extra  for new datasets

# workforce jobs:  totals and q/y changes
calc_workforce_jobs_changes <- function() {
  if (!exists("fetch_workforce_jobs", inherits = TRUE)) return(list(cur=NA, dq=NA, dy=NA, pct_q=NA, pct_y=NA))
  df <- tryCatch(fetch_workforce_jobs(), error = function(e) NULL)
  if (is.null(df) || nrow(df) == 0) return(list(cur=NA, dq=NA, dy=NA, pct_q=NA, pct_y=NA))

  # reuse
  parse_fn <- if (exists("parse_wfj_period_to_date", inherits = TRUE)) get("parse_wfj_period_to_date", inherits = TRUE) else function(x) NA

  df2 <- df %>%
    mutate(
      value = safe_num(value),
      period_date = as.Date(vapply(time_period, parse_fn, as.Date(NA)))
    ) %>%
    filter(!is.na(period_date), !is.na(value))

  if (nrow(df2) == 0) return(list(cur=NA, dq=NA, dy=NA, pct_q=NA, pct_y=NA))

  totals <- df2 %>%
    group_by(period_date) %>%
    summarise(total = sum(value, na.rm = TRUE), .groups="drop") %>%
    arrange(period_date)

  end_cur <- max(totals$period_date, na.rm = TRUE)
  cur <- totals$total[totals$period_date == end_cur][1]

  # previous quarter: 
  target_q <- end_cur %m-% months(3)
  prev_q_date <- max(totals$period_date[totals$period_date <= target_q], na.rm = TRUE)
  prev_q <- totals$total[totals$period_date == prev_q_date][1]

  # year ago:
  target_y <- end_cur %m-% months(12)
  prev_y_date <- max(totals$period_date[totals$period_date <= target_y], na.rm = TRUE)
  prev_y <- totals$total[totals$period_date == prev_y_date][1]

  dq <- if (!is.na(cur) && !is.na(prev_q)) cur - prev_q else NA_real_
  dy <- if (!is.na(cur) && !is.na(prev_y)) cur - prev_y else NA_real_

  list(
    cur = cur,
    dq = dq,
    dy = dy,
    pct_q = pct_from_delta(cur, dq),
    pct_y = pct_from_delta(cur, dy)
  )
}

# unemployment (18-24): 
calc_youth_unemp <- function() {
  if (!exists("fetch_unemployment_by_age", inherits = TRUE)) return(list(cur=NA, dq=NA, label=NA, is_largest_since_2022=NA))
  df <- tryCatch(fetch_unemployment_by_age(), error = function(e) NULL)
  if (is.null(df) || nrow(df) == 0) return(list(cur=NA, dq=NA, label=NA, is_largest_since_2022=NA))

  parse_fn <- if (exists("parse_lfs_period_to_end_date", inherits = TRUE)) get("parse_lfs_period_to_end_date", inherits = TRUE) else function(x) NA

  df2 <- df %>%
    mutate(
      value = safe_num(value),
      end_date = as.Date(vapply(time_period, parse_fn, as.Date(NA)))
    ) %>%
    filter(!is.na(end_date), !is.na(value))

  # pick 18-24 t
  youth <- df2 %>%
    filter(
      grepl("^\\s*18\\s*\\D*24", age_group, ignore.case = TRUE) | grepl("18", age_group) & grepl("24", age_group),
      tolower(trimws(value_type)) == "level"
    )

  if (nrow(youth) == 0) return(list(cur=NA, dq=NA, label=NA, is_largest_since_2022=NA))

  if ("All" %in% youth$duration) youth <- youth %>% filter(duration == "All")

  # keep latest period
  end_cur <- max(youth$end_date, na.rm = TRUE)
  cur_label <- youth %>% filter(end_date == end_cur) %>% slice(1) %>% pull(time_period) %>% as.character() %>% trimws()
  cur <- youth %>% filter(end_date == end_cur) %>% summarise(v = sum(value, na.rm=TRUE)) %>% pull(v)

  # previous quarter:
  target_q <- end_cur %m-% months(3)
  prev_q_date <- max(youth$end_date[youth$end_date <= target_q], na.rm = TRUE)
  prev_q <- youth %>% filter(end_date == prev_q_date) %>% summarise(v = sum(value, na.rm=TRUE)) %>% pull(v)

  dq <- if (!is.na(cur) && !is.na(prev_q)) cur - prev_q else NA_real_

  # largest increase since nov 2022 
  s <- youth %>%
    group_by(end_date) %>%
    summarise(v = sum(value, na.rm=TRUE), .groups="drop") %>%
    arrange(end_date) %>%
    mutate(dq = v - lag(v))

  s2 <- s %>% filter(end_date >= as.Date("2022-11-01"))
  max_inc <- if (nrow(s2) > 0) max(s2$dq, na.rm=TRUE) else NA_real_
  is_largest <- if (!is.na(dq) && !is.na(max_inc)) abs(dq - max_inc) < 1e-6 else NA

  list(cur=cur, dq=dq, label=cur_label, is_largest_since_2022=is_largest)
}

# payroll by age, monthly change by age group
calc_payroll_age_drops <- function() {
  if (!exists("fetch_payroll_by_age", inherits = TRUE)) return(list())
  df <- tryCatch(fetch_payroll_by_age(), error=function(e) NULL)
  if (is.null(df) || nrow(df) == 0) return(list())

  parse_fn <- if (exists("parse_month_label_to_date", inherits = TRUE)) get("parse_month_label_to_date", inherits = TRUE) else function(x) NA

  df2 <- df %>%
    mutate(
      value = safe_num(value),
      month_date = as.Date(vapply(time_period, parse_fn, as.Date(NA)))
    ) %>%
    filter(!is.na(month_date), !is.na(value))

  if (nrow(df2) == 0) return(list())

  latest_date <- max(df2$month_date, na.rm = TRUE)
  prev_date <- max(df2$month_date[df2$month_date < latest_date], na.rm = TRUE)

  cur <- df2 %>% filter(month_date == latest_date) %>% select(age_group, value)
  prev <- df2 %>% filter(month_date == prev_date) %>% select(age_group, value)

  chg <- cur %>%
    left_join(prev, by="age_group", suffix=c("_cur","_prev")) %>%
    mutate(delta = value_cur - value_prev) %>%
    arrange(delta) # most negative first

  if (nrow(chg) == 0) return(list())

  top3 <- head(chg, 3)

  get_age <- function(i) {
    if (nrow(top3) >= i) top3$age_group[i] else NA_character_
  }
  get_delta <- function(i) {
    if (nrow(top3) >= i) safe_num(top3$delta[i]) else NA_real_
  }

  # NOTE: return numeric deltas (levels). Formatting happens where used.
  list(
    a1 = get_age(1), d1 = get_delta(1),
    a2 = get_age(2), d2 = get_delta(2),
    a3 = get_age(3), d3 = get_delta(3)
  )
}

# generate summary (10 lines)

generate_summary <- function() {

  # top-level tryCatch: if anything fails, return fallback lines
  tryCatch({

  # NA-safe rate formatter
  .fmt_rate <- function(x) if (is.na(x)) "\u2014" else format(round(x, 1), nsmall = 1)

  mm <- if (exists("manual_month", inherits = TRUE)) get("manual_month", inherits = TRUE) else NA_character_

  # compute aligned/latest for payroll + vacancies 
  payroll_aligned <- tryCatch(calculate_payroll(mm, mode = "aligned"), error = function(e) NULL)
  payroll_latest  <- tryCatch(calculate_payroll(mm, mode = "latest"), error = function(e) NULL)
  vac_aligned     <- tryCatch(calculate_vacancies(mm, mode = "aligned"), error = function(e) NULL)
  vac_latest      <- tryCatch(calculate_vacancies(mm, mode = "latest"), error = function(e) NULL)

  # fallbacks 
  if (is.null(payroll_aligned) && exists("payroll", inherits = TRUE)) payroll_aligned <- get("payroll", inherits = TRUE)
  if (is.null(payroll_latest)  && exists("payroll", inherits = TRUE)) payroll_latest  <- get("payroll", inherits = TRUE)
  if (is.null(vac_latest)      && exists("vac", inherits = TRUE))     vac_latest      <- get("vac", inherits = TRUE)

  # labels
  lfs_lbl <- if (exists("lfs_period_label", inherits = TRUE)) get("lfs_period_label", inherits = TRUE) else ""
  payroll_flash_lbl <- if (exists("payroll_flash_label", inherits = TRUE)) get("payroll_flash_label", inherits = TRUE) else {
    if (!is.null(payroll_latest) && "flash_anchor" %in% names(payroll_latest) && !is.na(payroll_latest$flash_anchor)) {
      format(payroll_latest$flash_anchor, "%B %Y")
    } else ""
  }
  vac_lbl <- if (!is.null(vac_latest) && "end" %in% names(vac_latest) && !is.na(vac_latest$end)) {
    lfs_label_narrative(vac_latest$end)
  } else lfs_lbl

  wages_lbl <- if (exists("wages_nom", inherits = TRUE) && !is.null(get("wages_nom", inherits=TRUE)$anchor)) {
    paste0("three months to ", format(get("wages_nom", inherits=TRUE)$anchor, "%B %Y"))
  } else {
    paste0("three months to ", format(parse_manual_month(mm) %m-% months(2), "%B %Y"))
  }

  # payroll (aligned quarter )
  py_cur <- if (!is.null(payroll_aligned)) safe_num(payroll_aligned$cur) else NA_real_
  py_dy  <- if (!is.null(payroll_aligned)) safe_num(payroll_aligned$dy) else NA_real_
  py_dq  <- if (!is.null(payroll_aligned)) safe_num(payroll_aligned$dq) else NA_real_

  py_pct_y <- pct_from_delta(py_cur, py_dy)
  py_pct_q <- pct_from_delta(py_cur, py_dq)

  line1 <- glue(
    "The number of payrolled employees (PAYE) {ifelse(is.na(py_dy) || py_dy == 0, 'was unchanged', ifelse(py_dy < 0, 'fell', 'rose'))} by {fmt_int_1k(abs(py_dy) * 1000)} ({fmt_pct_unsigned(py_pct_y)}) on the year, ",
    "and {ifelse(is.na(py_dq) || py_dq == 0, 'was unchanged', ifelse(py_dq < 0, 'fell', 'rose'))} by {fmt_int_1k(abs(py_dq) * 1000)} ({fmt_pct_unsigned(py_pct_q)}) on the quarter in {lfs_lbl} (the period comparable with LFS).",
    .comment = ""
  )

  # payroll (flash, latest single month)
  pf_cur_m <- if (!is.null(payroll_latest)) safe_num(payroll_latest$flash_cur) else NA_real_
  pf_dy_k  <- if (!is.null(payroll_latest)) safe_num(payroll_latest$flash_dy) else NA_real_
  pf_dm_k  <- if (!is.null(payroll_latest)) safe_num(payroll_latest$flash_dm) else NA_real_

  # compute pct for flash
  pf_cur_n <- if (!is.na(pf_cur_m)) pf_cur_m * 1e6 else NA_real_
  pf_base_y <- if (!is.na(pf_cur_n) && !is.na(pf_dy_k)) pf_cur_n - pf_dy_k * 1000 else NA_real_
  pf_pct_y <- if (!is.na(pf_base_y) && pf_base_y != 0) (pf_dy_k * 1000) / pf_base_y * 100 else NA_real_
  pf_base_m <- if (!is.na(pf_cur_n) && !is.na(pf_dm_k)) pf_cur_n - pf_dm_k * 1000 else NA_real_
  pf_pct_m <- if (!is.na(pf_base_m) && pf_base_m != 0) (pf_dm_k * 1000) / pf_base_m * 100 else NA_real_

  line2 <- glue(
    "The ‘flash’ estimate for {payroll_flash_lbl} suggests payroll employees {ifelse(is.na(pf_dy_k) || pf_dy_k == 0, ‘were unchanged’, ifelse(pf_dy_k < 0, ‘fell’, ‘rose’))} by {fmt_int_1k(abs(pf_dy_k) * 1000)} ({fmt_pct_unsigned(pf_pct_y)}) on the year, ",
    "and {ifelse(is.na(pf_dm_k) || pf_dm_k == 0, ‘were unchanged’, ifelse(pf_dm_k < 0, ‘fell’, ‘rose’))} by {fmt_int_1k(abs(pf_dm_k) * 1000)} ({fmt_pct_unsigned(pf_pct_m)}) on the month, although this is prone to revision.",
    .comment = ""
  )

  # workforce jobs
  wfj <- calc_workforce_jobs_changes()
  wfj_dq <- wfj$dq; wfj_dy <- wfj$dy
  wfj_pct_q <- wfj$pct_q; wfj_pct_y <- wfj$pct_y

  line3 <- glue(
    "Workforce jobs data shows {ifelse(is.na(wfj_dq) || wfj_dq >= 0, 'an increase', 'a fall')} of {fmt_int_1k(abs(wfj_dq) * 1000)} jobs on the quarter ({ifelse(is.na(wfj_pct_q), '\u2014', paste0(ifelse(is.na(wfj_pct_q) || wfj_pct_q >= 0, '', '-'), fmt_pct_unsigned(wfj_pct_q))) }). ",
    "On the year, jobs {ifelse(is.na(wfj_dy) || wfj_dy >= 0, 'rose', 'fell')} by {fmt_int_1k(abs(wfj_dy) * 1000)} ({ifelse(is.na(wfj_pct_y), '\u2014', paste0(ifelse(is.na(wfj_pct_y) || wfj_pct_y >= 0, '', '-'), fmt_pct_unsigned(wfj_pct_y))) }).",
    .comment = ""
  )

  # vacancies (latest)
  vc_cur <- if (!is.null(vac_latest)) safe_num(vac_latest$cur) else NA_real_
  vc_dq  <- if (!is.null(vac_latest)) safe_num(vac_latest$dq) else NA_real_
  vc_pct_q <- pct_from_delta(vc_cur, vc_dq)

  vac_change_word <- if (is.na(vc_dq) || vc_dq == 0) "were unchanged" else if (vc_dq < 0) "fell" else "rose"
  line4 <- glue(
    "Vacancies {vac_change_word} by {fmt_int_1k(abs(vc_dq) * 1000)} ({fmt_pct_unsigned(vc_pct_q)}) on the quarter to {fmt_int_1k(vc_cur * 1000)} in {vac_lbl}.",
    .comment = ""
  )

  # lfs rates (current + quarter)
  emp_rt_cur <- if (exists("emp_rt_cur", inherits = TRUE)) safe_num(get("emp_rt_cur", inherits=TRUE)) else NA_real_
  emp_rt_dq  <- if (exists("emp_rt_dq", inherits = TRUE)) safe_num(get("emp_rt_dq", inherits=TRUE)) else NA_real_
  unemp_rt_cur <- if (exists("unemp_rt_cur", inherits = TRUE)) safe_num(get("unemp_rt_cur", inherits=TRUE)) else NA_real_
  unemp_rt_dq  <- if (exists("unemp_rt_dq", inherits = TRUE)) safe_num(get("unemp_rt_dq", inherits=TRUE)) else NA_real_
  inact_rt_cur <- if (exists("inact_rt_cur", inherits = TRUE)) safe_num(get("inact_rt_cur", inherits=TRUE)) else NA_real_
  inact_rt_dq  <- if (exists("inact_rt_dq", inherits = TRUE)) safe_num(get("inact_rt_dq", inherits=TRUE)) else NA_real_

  line5 <- glue(
    "Labour Force Survey (LFS) suggests that in {lfs_lbl}, the employment rate {ifelse(is.na(emp_rt_dq) || emp_rt_dq == 0, 'was unchanged', ifelse(emp_rt_dq < 0, 'fell', 'rose'))} to {.fmt_rate(emp_rt_cur)}% ({fmt_signed_pp(emp_rt_dq)} compared to the previous quarter). ",
    "On the quarter, unemployment {ifelse(is.na(unemp_rt_dq) || unemp_rt_dq >= 0, 'rose', 'fell')} to {.fmt_rate(unemp_rt_cur)}% ({fmt_signed_pp(unemp_rt_dq)}) and inactivity {ifelse(is.na(inact_rt_dq) || inact_rt_dq > 0, 'rose', 'fell')} to {.fmt_rate(inact_rt_cur)}% ({fmt_signed_pp(inact_rt_dq)}).",
    .comment = ""
  )

  # youth unemployment (18-24)
  yu <- calc_youth_unemp()
  yu_phrase <- if (isTRUE(yu$is_largest_since_2022)) "the largest increase since November 2022" else "a large increase"
  # SL6 not required
  line6 <- ""

  # payroll by age monthly drops
  pa <- calc_payroll_age_drops()
  if (is.null(pa$a1) || length(pa$a1) == 0 || is.na(pa$a1)) pa$a1 <- "age group"
  if (is.null(pa$a2) || length(pa$a2) == 0 || is.na(pa$a2)) pa$a2 <- "age group"
  if (is.null(pa$a3) || length(pa$a3) == 0 || is.na(pa$a3)) pa$a3 <- "age group"
  # present as fall/rise rather than signed numbers
  fmt_k_abs <- function(x) {
    if (length(x) == 0 || is.null(x) || is.na(x)) return("0")
    format(abs(round(as.numeric(x)/1000)), big.mark = ",")
  }
  line7 <- glue(
    "In contrast, monthly payrolled employees {ifelse(is.null(pa$d1) || length(pa$d1) == 0 || is.na(pa$d1) || pa$d1 == 0, 'were unchanged', ifelse(pa$d1 < 0, 'fell', 'rose'))} by {fmt_k_abs(pa$d1)}k for {pa$a1}, followed by {pa$a2} ({ifelse(is.null(pa$d2) || length(pa$d2) == 0 || is.na(pa$d2) || pa$d2 == 0, 'were unchanged', ifelse(pa$d2 < 0, 'fell', 'rose'))} by {fmt_k_abs(pa$d2)}k) and {pa$a3} ({ifelse(is.null(pa$d3) || length(pa$d3) == 0 || is.na(pa$d3) || pa$d3 == 0, 'were unchanged', ifelse(pa$d3 < 0, 'fell', 'rose'))} by {fmt_k_abs(pa$d3)}k).",
    .comment = ""
  )

  # wages (nominal + real)
  latest_wages <- if (exists("latest_wages", inherits = TRUE)) safe_num(get("latest_wages", inherits=TRUE)) else NA_real_
  latest_regular_cash <- if (exists("latest_regular_cash", inherits = TRUE)) safe_num(get("latest_regular_cash", inherits=TRUE)) else NA_real_
  latest_wages_cpi <- if (exists("latest_wages_cpi", inherits = TRUE)) safe_num(get("latest_wages_cpi", inherits=TRUE)) else NA_real_
  wages_total_qchange <- if (exists("wages_total_qchange", inherits = TRUE)) safe_num(get("wages_total_qchange", inherits=TRUE)) else NA_real_
  wages_reg_qchange <- if (exists("wages_reg_qchange", inherits = TRUE)) safe_num(get("wages_reg_qchange", inherits=TRUE)) else NA_real_

  line8 <- glue(
    "Annual wage growth in average weekly earnings (inc. bonuses) {ifelse(is.na(wages_total_qchange) || wages_total_qchange < 0, 'fell', 'rose')} to {fmt_pct(latest_wages)} in {wages_lbl} ({fmt_signed_pp(wages_total_qchange)} from the previous 3-month period). ",
    "Wage growth excl. bonuses also {ifelse(is.na(wages_reg_qchange) || wages_reg_qchange < 0, 'fell', 'rose')} to {fmt_pct(latest_regular_cash)} ({fmt_signed_pp(wages_reg_qchange)}). Real wage growth (inc. bonuses) {ifelse(is.na(latest_wages_cpi) || latest_wages_cpi < 0, 'dropped', 'rose')} to {fmt_pct(latest_wages_cpi)}.",
    .comment = ""
  )

  # public vs private (yoy)
  wages_total_public <- if (exists("wages_total_public", inherits = TRUE)) safe_num(get("wages_total_public", inherits=TRUE)) else NA_real_
  wages_total_private <- if (exists("wages_total_private", inherits = TRUE)) safe_num(get("wages_total_private", inherits=TRUE)) else NA_real_

  # order sectors by highest growth first
  first_sector <- ifelse(is.na(wages_total_public) || is.na(wages_total_private) || wages_total_public >= wages_total_private, "public sector", "private sector")
  first_val <- ifelse(is.na(wages_total_public) || is.na(wages_total_private) || wages_total_public >= wages_total_private, wages_total_public, wages_total_private)
  second_sector <- ifelse(is.na(wages_total_public) || is.na(wages_total_private) || wages_total_public >= wages_total_private, "private sector", "public sector")
  second_val <- ifelse(is.na(wages_total_public) || is.na(wages_total_private) || wages_total_public >= wages_total_private, wages_total_private, wages_total_public)
  line9 <- glue(
    "Pay growth was driven by the {first_sector} ({fmt_pct(first_val)}), compared to {fmt_pct(second_val)} in the {second_sector}.",
    .comment = ""
  )

  # redundancies + hr1
  redund_cur <- if (exists("redund_cur", inherits = TRUE)) safe_num(get("redund_cur", inherits=TRUE)) else NA_real_
  redund_dq <- if (exists("redund_dq", inherits = TRUE)) safe_num(get("redund_dq", inherits=TRUE)) else NA_real_
  hr1_dm <- if (exists("hr1_dm", inherits = TRUE)) safe_num(get("hr1_dm", inherits=TRUE)) else NA_real_

  line10 <- glue(
    "LFS redundancies {ifelse(is.na(redund_dq) || redund_dq >= 0, 'rose', 'fell')} on the quarter to {fmt_int_1k(redund_cur * 1000)} in {lfs_lbl} ({fmt_signed_int_1k(redund_dq * 1000)} from the previous quarter). ",
    "HR1 redundancies (notifications of redundancies) {ifelse(is.na(hr1_dm) || hr1_dm >= 0, 'rose', 'fell')} by {fmt_int_1k(abs(hr1_dm))} on the month.",
    .comment = ""
  )

  list(
    line1 = as.character(line1),
    line2 = as.character(line2),
    line3 = as.character(line3),
    line4 = as.character(line4),
    line5 = as.character(line5),
    line6 = as.character(line6),
    line7 = as.character(line7),
    line8 = as.character(line8),
    line9 = as.character(line9),
    line10 = as.character(line10)
  )

  }, error = function(e) {
    warning("generate_summary() internal error: ", e$message, "\n", paste(capture.output(traceback()), collapse = "\n"))
    fallback <- list()
    for (i in 1:10) fallback[[paste0("line", i)]] <- paste0("(Data unavailable: ", e$message, ")")
    fallback
  })
}
