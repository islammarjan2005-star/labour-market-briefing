# top ten stats module - narrative generation for briefing

library(glue)

# ensure helpers are loaded (formatters like fmt_pct, fmt_pp, fmt_dir, etc.)
if (!exists("fmt_pct", inherits = TRUE)) {
  source("utils/helpers.R")
}

# helper functions

dir_word <- function(x) {
  if (is.na(x)) "a change of"
  else if (x > 0) "an increase of"
  else if (x < 0) "a decrease of"
  else "no change,"
}

# format number rounded to nearest 1,000
fmt_int_1k <- function(x) {
  if (is.na(x)) return("\u2014")
  format(round(abs(x), -3), big.mark = ",")
}

# format flash change (rounded to nearest 1,000)
fmt_flash_change <- function(x) {
  if (is.na(x)) return("\u2014")
  full_val <- round(x * 1000, -3)
  format(abs(full_val), big.mark = ",")
}

# format rate (NA-safe wrapper for format(round(x,1)))
fmt_rate <- function(x) {
  if (is.na(x)) return("\u2014")
  format(round(x, 1), nsmall = 1)
}

# generate top ten stats

generate_top_ten <- function() {

  # top-level tryCatch: if anything fails, return fallback lines
  tryCatch({

  # safe variable accessor — returns default if variable doesn't exist or is NULL
  sv <- function(name, default = NA_real_) {
    if (exists(name, inherits = TRUE)) {
      val <- get(name, inherits = TRUE)
      if (is.null(val)) default else val
    } else default
  }

  # pull all needed variables safely
  wages_total_public <- sv("wages_total_public")
  wages_total_private <- sv("wages_total_private")
  wages_reg_public <- sv("wages_reg_public")
  wages_reg_private <- sv("wages_reg_private")
  latest_wages <- sv("latest_wages")
  latest_regular_cash <- sv("latest_regular_cash")
  latest_wages_cpi <- sv("latest_wages_cpi")
  latest_regular_cpi <- sv("latest_regular_cpi")
  wages_cpi_total_vs_dec2007 <- sv("wages_cpi_total_vs_dec2007")
  wages_cpi_total_vs_pandemic <- sv("wages_cpi_total_vs_pandemic")
  wages_total_qchange <- sv("wages_total_qchange")
  wages_reg_qchange <- sv("wages_reg_qchange")
  emp_rt_cur <- sv("emp_rt_cur")
  emp_rt_dy <- sv("emp_rt_dy")
  emp_rt_dq <- sv("emp_rt_dq")
  emp_rt_dc <- sv("emp_rt_dc")
  inact_rt_cur <- sv("inact_rt_cur")
  inact_rt_dy <- sv("inact_rt_dy")
  inact_rt_dq <- sv("inact_rt_dq")
  inact_rt_dc <- sv("inact_rt_dc")
  unemp_rt_cur <- sv("unemp_rt_cur")
  unemp_rt_dy <- sv("unemp_rt_dy")
  unemp_rt_dq <- sv("unemp_rt_dq")
  unemp_rt_dc <- sv("unemp_rt_dc")
  payroll_flash_de <- sv("payroll_flash_de")
  payroll_flash_dy <- sv("payroll_flash_dy")
  payroll_flash_dm <- sv("payroll_flash_dm")
  payroll_flash_cur <- sv("payroll_flash_cur")
  payroll_flash_label <- sv("payroll_flash_label", "")
  hosp_dy <- sv("hosp_dy")
  retail_dy <- sv("retail_dy")
  health_dy <- sv("health_dy")
  days_lost_cur <- sv("days_lost_cur")
  vac_dy <- sv("vac_dy")
  vac_cur <- sv("vac_cur")
  vac_dc <- sv("vac_dc")
  redund_dq <- sv("redund_dq")
  redund_dy <- sv("redund_dy")
  redund_cur <- sv("redund_cur")
  hr1_cur <- sv("hr1_cur")
  hr1_month_label <- sv("hr1_month_label", "")
  lfs_period_label <- sv("lfs_period_label", "")

  # calculate public/private dif
  pub_priv_diff_total <- if (!is.na(wages_total_public) && !is.na(wages_total_private)) {
    wages_total_public - wages_total_private
  } else NA_real_

  pub_priv_diff_reg <- if (!is.na(wages_reg_public) && !is.na(wages_reg_private)) {
    wages_reg_public - wages_reg_private
  } else NA_real_

  #  direction words
  pub_priv_dir <- ifelse(!is.na(pub_priv_diff_reg) && pub_priv_diff_reg >= 0, "higher", "lower")
  wages_cpi_total_dir <- ifelse(!is.na(latest_wages_cpi) && latest_wages_cpi >= 0, "grew by", "fell by")
  wages_cpi_reg_dir <- ifelse(!is.na(latest_regular_cpi) && latest_regular_cpi >= 0, "grew by", "fell by")
  vs_dec2007_dir <- ifelse(!is.na(wages_cpi_total_vs_dec2007) && wages_cpi_total_vs_dec2007 >= 0, "higher", "lower")
  vs_pandemic_dir <- ifelse(!is.na(wages_cpi_total_vs_pandemic) && wages_cpi_total_vs_pandemic >= 0, "higher", "lower")
  emp_rt_covid_dir <- ifelse(!is.na(emp_rt_dc) && emp_rt_dc >= 0, "higher", "lower")
  inact_rt_covid_dir <- ifelse(!is.na(inact_rt_dc) && inact_rt_dc >= 0, "higher", "lower")
  unemp_rt_covid_dir <- ifelse(!is.na(unemp_rt_dc) && unemp_rt_dc >= 0, "higher", "lower")

  # line 1: wages nominal
  line1 <- glue(
    'Annual growth in employees\' average earnings was {fmt_pct(latest_wages)} for total pay ',
    '(including bonuses) and {fmt_pct(latest_regular_cash)} for regular pay (excluding bonuses) ',
    'in {lfs_period_label}. Public sector changes of {fmt_pct(wages_total_public)} and ',
    '{fmt_pct(wages_reg_public)} respectively are {fmt_one_dec(abs(pub_priv_diff_reg))}pp {pub_priv_dir} than the private sector. ',
    'Wage growth is {ifelse(is.na(wages_total_qchange) || wages_total_qchange < 0, "easing", "accelerating")}, with this representing a quarterly ',
    '{ifelse(is.na(wages_total_qchange) || wages_total_qchange >= 0, "increase", "decline")} of {fmt_one_dec(abs(wages_total_qchange))}pp ',
    'and a quarterly {ifelse(is.na(wages_reg_qchange) || wages_reg_qchange >= 0, "increase", "decline")} of {fmt_one_dec(abs(wages_reg_qchange))}pp respectively.',
    .comment = ""
  )

  # line 2: wages cpi
  line2 <- glue(
    'Adjusted for CPI inflation, total and regular pay in {lfs_period_label} {wages_cpi_total_dir} ',
    '{fmt_pct(abs(latest_wages_cpi))} and {wages_cpi_reg_dir} {fmt_pct(abs(latest_regular_cpi))} respectively. ',
    'Inflation-adjusted total wages are now around {fmt_one_dec(abs(wages_cpi_total_vs_dec2007))}% {vs_dec2007_dir} than they were in ',
    'December 2007 prior to the global financial crisis and {fmt_one_dec(abs(wages_cpi_total_vs_pandemic))}% {vs_pandemic_dir} than ',
    'before the pandemic (2019 average).',
    .comment = ""
  )

  # line 3: employment rate
  line3 <- glue(
    'The 16-64 employment rate was {fmt_rate(emp_rt_cur)}% in ',
    '{lfs_period_label}, which is {fmt_dir(emp_rt_dy, "up", "down")} ',
    '{fmt_pp(emp_rt_dy)} from a year ago and {fmt_dir(emp_rt_dq, "up", "down")} ',
    '{fmt_pp(emp_rt_dq)} on the last quarter. The employment rate is ',
    '{fmt_pp(emp_rt_dc)} {emp_rt_covid_dir} than before the pandemic.',
    .comment = ""
  )

  # line 4: payroll flash
  flash_de_dir <- if (!is.na(payroll_flash_de) && payroll_flash_de < 0) "a fall of around" else "a rise of around"
  flash_dy_dir <- if (!is.na(payroll_flash_dy) && payroll_flash_dy < 0) "a fall of" else "a rise of"
  flash_dm_dir <- if (!is.na(payroll_flash_dm) && payroll_flash_dm < 0) "a fall of" else "a rise of"

  # sector direction helpers
  hosp_dy_dir <- if (!is.na(hosp_dy) && hosp_dy < 0) "a fall of" else if (!is.na(hosp_dy) && hosp_dy > 0) "a rise of" else "no change of"
  retail_dy_dir <- if (!is.na(retail_dy) && retail_dy < 0) "a fall of" else if (!is.na(retail_dy) && retail_dy > 0) "a rise of" else "no change of"
  health_dy_dir <- if (!is.na(health_dy) && health_dy < 0) "fell by" else if (!is.na(health_dy) && health_dy > 0) "rose by" else "was unchanged by"

  line4 <- glue(
    'Early "flash" estimates for {payroll_flash_label} indicate that ',
    'there were {fmt_rate(payroll_flash_cur)}M payrolled employees. ',
    'This represents {flash_de_dir} ',
    '{fmt_flash_change(payroll_flash_de)} since the 2024 election, {flash_dy_dir} {fmt_flash_change(payroll_flash_dy)} ',
    'compared to the same period a year ago, and {flash_dm_dir} ',
    '{fmt_flash_change(payroll_flash_dm)} from the previous month. This varies ',
    'between sectors \u2013 changes in employee numbers on the year have been ',
    'concentrated in sectors such as hospitality and ',
    'retail which saw {hosp_dy_dir} {fmt_int_1k(hosp_dy * 1000)} and ',
    '{retail_dy_dir} {fmt_int_1k(retail_dy * 1000)} respectively. Employee numbers in the health and ',
    'social work sector {health_dy_dir} {fmt_int_1k(health_dy * 1000)} in the same period. [Note: payroll ',
    'employment data is frequently revised; data here is a "flash" estimate ',
    'and so does not align with the table above].',
    .comment = ""
  )

  # line 5: inactivity rate
  line5 <- glue(
    'The 16-64s economic inactivity rate was {fmt_rate(inact_rt_cur)}% ',
    'in {lfs_period_label}, {fmt_dir(inact_rt_dy, "down", "up")} {fmt_pp(inact_rt_dy)} ',
    'from a year ago, and {fmt_dir(inact_rt_dq, "down", "up")} {fmt_pp(inact_rt_dq)} ',
    'from the previous quarter. The inactivity rate is ',
    '{fmt_pp(inact_rt_dc)} {inact_rt_covid_dir} than before the pandemic.',
    .comment = ""
  )

  # line 6: unemployment rate
  line6 <- glue(
    'The unemployment rate for 16+ was {fmt_rate(unemp_rt_cur)}% ',
    'in {lfs_period_label}, {fmt_dir(unemp_rt_dq, "an increase of", "a decrease of")} ',
    '{fmt_pp(unemp_rt_dq)} from the previous quarter, and ',
    '{fmt_dir(unemp_rt_dy, "an increase of", "a decrease of")} {fmt_pp(unemp_rt_dy)} ',
    'from this time last year. The unemployment rate is {fmt_pp(unemp_rt_dc)} ',
    '{unemp_rt_covid_dir} than before the pandemic.',
    .comment = ""
  )

  # line 7: days lost
  days_lost_lbl <- if (exists("days_lost_label", inherits = TRUE)) as.character(get("days_lost_label", inherits = TRUE)) else ""
  days_lost_month <- if (nzchar(days_lost_lbl) && !is.na(days_lost_lbl) && days_lost_lbl != "NA") {
    paste0("in ", days_lost_lbl, " ")
  } else ""

  line7 <- glue(
    'There were an estimated {fmt_int_1k(days_lost_cur * 1000)} ',
    'working days lost {days_lost_month}because of labour disputes across the UK.',
    .comment = ""
  )

  # line 8: vacancies
  # use the vacancies reference period (can differ from lfs period label)
  vac_period_label <- lfs_period_label
  if (exists("vac", inherits = TRUE)) {
    v <- get("vac", inherits = TRUE)
    if (!is.null(v) && "end" %in% names(v) && !is.na(v$end)) {
      vac_period_label <- make_lfs_label(v$end)
    }
  }
  if (exists("vacancies_period_label", inherits = TRUE)) {
    vpl <- get("vacancies_period_label", inherits = TRUE)
    if (!is.null(vpl) && length(vpl) == 1 && nzchar(as.character(vpl))) {
      vac_period_label <- as.character(vpl)
    }
  }
  vac_dir <- ifelse(is.na(vac_dy) || vac_dy < 0, "decreased", "increased")
  vac_cur_vs_peak <- ifelse(is.na(vac_cur) || vac_cur < 1300, "falling", "rising")
  vac_prepandemic_phrase <- if (!is.na(vac_dc) && vac_dc != 0) {
    paste0("; and are ", fmt_int_1k(vac_dc * 1000), ifelse(is.na(vac_dc) || vac_dc < 0, " below", " above"), " the pre-pandemic level")
  } else ""
  vac_cur_vs_2010avg <- ifelse(is.na(vac_cur) || vac_cur < 660, "below", "above")

  line8 <- glue(
    'The number of job vacancies {vac_dir} on the year by {fmt_int_1k(vac_dy * 1000)} ',
    'to {fmt_int_1k(vac_cur * 1000)} in {vac_period_label}. Vacancies have been ',
    '{vac_cur_vs_peak} since the peak in March to May 2022 (1.3 million)',
    '{vac_prepandemic_phrase}. However, they ',
    'are {vac_cur_vs_2010avg} the average in the 2010s (around 660,000).',
    .comment = ""
  )

  # line 9: redundancies
  redund_q_dir <- dir_word(redund_dq)
  redund_y_dir <- dir_word(redund_dy)
  avg2010 <- if (!is.na(redund_cur) && redund_cur < 4.7) "below" else "above"

  line9 <- glue(
    "The number of people reporting redundancy in {lfs_period_label}, ",
    "according to the Labour Force Survey, was {fmt_one_dec(redund_cur)} ",
    "per thousand employees, {redund_q_dir} ",
    "{fmt_one_dec(abs(redund_dq))} per thousand employees ",
    "from the previous quarter and {redund_y_dir} ",
    "{fmt_one_dec(abs(redund_dy))} per thousand employees ",
    "compared to the same period last year. This is {avg2010} the average ",
    "in the 2010s of 4.7 redundancies per thousand employees.",
    .comment = ""
  )

  # line 10: hr1
  hr1_vs_prepandemic <- if (!is.na(hr1_cur) && hr1_cur < 27600) "below" else "above"

  hr1_in_month <- if (nzchar(hr1_month_label) && !is.na(hr1_month_label) && hr1_month_label != "NA") {
    paste0("in ", hr1_month_label)
  } else ""

  line10 <- glue(
    'The Insolvency Service were notified of {fmt_int(hr1_cur)} ',
    'potential redundancies {hr1_in_month}. This is ',
    '{hr1_vs_prepandemic} the pre-pandemic average of 27,600 (Apr 2019 \u2013 Feb 2020).',
    .comment = ""
  )

  list(
    line1 = line1,
    line2 = line2,
    line3 = line3,
    line4 = line4,
    line5 = line5,
    line6 = line6,
    line7 = line7,
    line8 = line8,
    line9 = line9,
    line10 = line10
  )

  }, error = function(e) {
    warning("generate_top_ten() internal error: ", e$message, "\n", paste(capture.output(traceback()), collapse = "\n"))
    fallback <- list()
    for (i in 1:10) fallback[[paste0("line", i)]] <- paste0("(Data unavailable: ", e$message, ")")
    fallback
  })
}

# print top ten stats

print_top_ten <- function(stats) {
  cat('\n')
  cat('1.', stats$line1, '\n\n')
  cat('2.', stats$line2, '\n\n')
  cat('3.', stats$line3, '\n\n')
  cat('4.', stats$line4, '\n\n')

  cat('5.', stats$line5, '\n\n')
  cat('6.', stats$line6, '\n\n')
  cat('7.', stats$line7, '\n\n')
  cat('8.', stats$line8, '\n\n')
  cat('9.', stats$line9, '\n\n')
  cat('10.', stats$line10, '\n\n')

}
