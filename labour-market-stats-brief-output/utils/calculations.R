# calculations - sources all sheets and computes all metrics

# load common helpers
source("utils/helpers.R")

# load config
if (!exists("manual_month", inherits = TRUE)) {
  source("utils/config.R")
}

# load all sheets
source("sheets/lfs.R")
source("sheets/vacancies.R")
source("sheets/payroll.R")
source("sheets/wages_nominal.R")
source("sheets/wages_cpi.R")
source("sheets/days_lost.R")
source("sheets/redundancy.R")
source("sheets/sector_payroll.R")
source("sheets/hr1.R")
source("sheets/inactivity_reasons.R")
source("sheets/summary.R")
source("sheets/workforce_jobs.R")
source("sheets/unemployment_by_age.R")
source("sheets/payroll_by_age.R")

# significance testing (ONS CI methodology)
source("utils/significance.R")


# safe extraction helper
.safe <- function(obj, ...) {
  fields <- list(...)
  val <- obj
  for (f in fields) {
    if (is.null(val)) return(NA_real_)
    val <- val[[f]]
  }
  if (is.null(val)) NA_real_ else val
}

.safe_chr <- function(obj, ...) {
  fields <- list(...)
  val <- obj
  for (f in fields) {
    if (is.null(val)) return(NA_character_)
    val <- val[[f]]
  }
  if (is.null(val)) NA_character_ else val
}

.safe_date <- function(obj, ...) {
  fields <- list(...)
  val <- obj
  for (f in fields) {
    if (is.null(val)) return(as.Date(NA))
    val <- val[[f]]
  }
  if (is.null(val)) as.Date(NA) else val
}

# calculate all metrics (each wrapped in tryCatch)

# lfs metrics (employment, unemployment, inactivity)
lfs <- tryCatch(calculate_lfs(manual_month), error = function(e) { warning("LFS calculation failed: ", e$message); NULL })

# vacancies
# some releases have vacancies/payroll updated ahead of the  lfs quarter.
vac_mode <- if (exists("vacancies_mode", inherits = TRUE)) get("vacancies_mode", inherits = TRUE) else "latest"
vac <- tryCatch(calculate_vacancies(manual_month, mode = vac_mode), error = function(e) { warning("Vacancies calculation failed: ", e$message); NULL })

# payrolled employees
payroll_mode <- if (exists("payroll_mode", inherits = TRUE)) get("payroll_mode", inherits = TRUE) else "latest"
payroll <- tryCatch(calculate_payroll(manual_month, mode = payroll_mode), error = function(e) { warning("Payroll calculation failed: ", e$message); NULL })

# nominal wages (total + regular)
wages_nom <- tryCatch(calculate_wages_nominal(manual_month), error = function(e) { warning("Wages nominal calculation failed: ", e$message); NULL })

# cpi-adjusted wages (total + regular)
wages_cpi <- tryCatch(calculate_wages_cpi(manual_month), error = function(e) { warning("Wages CPI calculation failed: ", e$message); NULL })

# working days lost
days_lost <- tryCatch(calculate_days_lost(manual_month), error = function(e) { warning("Days lost calculation failed: ", e$message); NULL })

# redundancy (lfs)
redund <- tryCatch(calculate_redundancy(manual_month), error = function(e) { warning("Redundancy calculation failed: ", e$message); NULL })

# sector payroll (hospitality, retail, health)
sectors <- tryCatch(calculate_sector_payroll(manual_month), error = function(e) { warning("Sector payroll calculation failed: ", e$message); NULL })

# hr1 redundancy notifications
hr1 <- tryCatch(calculate_hr1(), error = function(e) { warning("HR1 calculation failed: ", e$message); NULL })

# inactivity by reason
inact_reasons <- tryCatch(calculate_inactivity_reasons(manual_month), error = function(e) { warning("Inactivity reasons failed: ", e$message); NULL })

# workforce jobs (by industry)
workforce_jobs <- tryCatch(calculate_workforce_jobs(manual_month), error = function(e) { warning("Workforce jobs failed: ", e$message); NULL })

# unemployment by age
unemployment_by_age <- tryCatch(calculate_unemployment_by_age(manual_month), error = function(e) { warning("Unemployment by age failed: ", e$message); NULL })

# payroll employees by age
payroll_by_age <- tryCatch(calculate_payroll_by_age(manual_month), error = function(e) { warning("Payroll by age failed: ", e$message); NULL })


# anchor date for labels

cm <- parse_manual_month(manual_month)
anchor_m <- cm %m-% months(2)

#  variables for dashboard (all guarded against NULL)

# employment
emp16_cur <- .safe(lfs, "emp16", "cur")
emp16_dq <- .safe(lfs, "emp16", "dq")
emp16_dy <- .safe(lfs, "emp16", "dy")
emp16_dc <- .safe(lfs, "emp16", "dc")
emp16_de <- .safe(lfs, "emp16", "de")

emp_rt_cur <- .safe(lfs, "emp_rt", "cur")
emp_rt_dq <- .safe(lfs, "emp_rt", "dq")
emp_rt_dy <- .safe(lfs, "emp_rt", "dy")
emp_rt_dc <- .safe(lfs, "emp_rt", "dc")
emp_rt_de <- .safe(lfs, "emp_rt", "de")

# unemployment
unemp16_cur <- .safe(lfs, "unemp16", "cur")
unemp16_dq <- .safe(lfs, "unemp16", "dq")
unemp16_dy <- .safe(lfs, "unemp16", "dy")
unemp16_dc <- .safe(lfs, "unemp16", "dc")
unemp16_de <- .safe(lfs, "unemp16", "de")

unemp_rt_cur <- .safe(lfs, "unemp_rt", "cur")
unemp_rt_dq <- .safe(lfs, "unemp_rt", "dq")
unemp_rt_dy <- .safe(lfs, "unemp_rt", "dy")
unemp_rt_dc <- .safe(lfs, "unemp_rt", "dc")
unemp_rt_de <- .safe(lfs, "unemp_rt", "de")

# inactivity
inact_cur <- .safe(lfs, "inact", "cur")
inact_dq <- .safe(lfs, "inact", "dq")
inact_dy <- .safe(lfs, "inact", "dy")
inact_dc <- .safe(lfs, "inact", "dc")
inact_de <- .safe(lfs, "inact", "de")

inact_rt_cur <- .safe(lfs, "inact_rt", "cur")
inact_rt_dq <- .safe(lfs, "inact_rt", "dq")
inact_rt_dy <- .safe(lfs, "inact_rt", "dy")
inact_rt_dc <- .safe(lfs, "inact_rt", "dc")
inact_rt_de <- .safe(lfs, "inact_rt", "de")

# 50-64 inactivity
inact5064_cur <- .safe(lfs, "inact5064", "cur")
inact5064_dq <- .safe(lfs, "inact5064", "dq")
inact5064_dy <- .safe(lfs, "inact5064", "dy")
inact5064_dc <- .safe(lfs, "inact5064", "dc")
inact5064_de <- .safe(lfs, "inact5064", "de")

inact5064_rt_cur <- .safe(lfs, "inact5064_rt", "cur")
inact5064_rt_dq <- .safe(lfs, "inact5064_rt", "dq")
inact5064_rt_dy <- .safe(lfs, "inact5064_rt", "dy")
inact5064_rt_dc <- .safe(lfs, "inact5064_rt", "dc")
inact5064_rt_de <- .safe(lfs, "inact5064_rt", "de")

# vacancies
vac_cur <- .safe(vac, "cur")
vac_dq <- .safe(vac, "dq")
vac_dy <- .safe(vac, "dy")
vac_dc <- .safe(vac, "dc")
vac_de <- .safe(vac, "de")

# payroll (dashboard - 3 month avg)
payroll_cur <- .safe(payroll, "cur")
payroll_dq <- .safe(payroll, "dq")
payroll_dy <- .safe(payroll, "dy")
payroll_dc <- .safe(payroll, "dc")
payroll_de <- .safe(payroll, "de")



# payroll flash ( - latest single month)
payroll_flash_cur <- .safe(payroll, "flash_cur")
payroll_flash_dy <- .safe(payroll, "flash_dy")
payroll_flash_de <- .safe(payroll, "flash_de")
payroll_flash_dm <- .safe(payroll, "flash_dm")

# nominal wages - total
latest_wages <- .safe(wages_nom, "total", "cur")
wages_change_q <- .safe(wages_nom, "total", "dq")
wages_change_y <- .safe(wages_nom, "total", "dy")
wages_change_covid <- .safe(wages_nom, "total", "dc")
wages_change_election <- .safe(wages_nom, "total", "de")
wages_total_public <- .safe(wages_nom, "total", "public")
wages_total_private <- .safe(wages_nom, "total", "private")
wages_total_qchange <- .safe(wages_nom, "total", "qchange")

# nominal wages - regular
latest_regular_cash <- .safe(wages_nom, "regular", "cur")
wages_reg_change_q <- .safe(wages_nom, "regular", "dq")
wages_reg_change_y <- .safe(wages_nom, "regular", "dy")
wages_reg_change_covid <- .safe(wages_nom, "regular", "dc")
wages_reg_change_election <- .safe(wages_nom, "regular", "de")
wages_reg_public <- .safe(wages_nom, "regular", "public")
wages_reg_private <- .safe(wages_nom, "regular", "private")
wages_reg_qchange <- .safe(wages_nom, "regular", "qchange")

# cpi wages - total
latest_wages_cpi <- .safe(wages_cpi, "total", "cur")
wages_cpi_change_q <- .safe(wages_cpi, "total", "dq")
wages_cpi_change_y <- .safe(wages_cpi, "total", "dy")
wages_cpi_change_covid <- .safe(wages_cpi, "total", "dc")
wages_cpi_change_election <- .safe(wages_cpi, "total", "de")
wages_cpi_total_vs_dec2007 <- .safe(wages_cpi, "total", "pct_vs_dec2007")
wages_cpi_total_vs_pandemic <- .safe(wages_cpi, "total", "pct_vs_pandemic")

# cpi wages - regular
latest_regular_cpi <- .safe(wages_cpi, "regular", "cur")
wages_reg_cpi_change_q <- .safe(wages_cpi, "regular", "dq")
wages_reg_cpi_change_y <- .safe(wages_cpi, "regular", "dy")
wages_reg_cpi_change_covid <- .safe(wages_cpi, "regular", "dc")
wages_reg_cpi_change_election <- .safe(wages_cpi, "regular", "de")
wages_cpi_reg_vs_dec2007 <- .safe(wages_cpi, "regular", "pct_vs_dec2007")
wages_cpi_reg_vs_pandemic <- .safe(wages_cpi, "regular", "pct_vs_pandemic")

# days lost
days_lost_cur <- .safe(days_lost, "cur")
days_lost_label <- .safe_chr(days_lost, "label")

# redundancy
redund_cur <- .safe(redund, "cur")
redund_dq <- .safe(redund, "dq")
redund_dy <- .safe(redund, "dy")
redund_dc <- .safe(redund, "dc")
redund_de <- .safe(redund, "de")

# sectors - hospitality
hosp_cur <- .safe(sectors, "hospitality", "cur")
hosp_dm <- .safe(sectors, "hospitality", "dm")
hosp_dy <- .safe(sectors, "hospitality", "dy")
hosp_dc <- .safe(sectors, "hospitality", "dc")
hosp_de <- .safe(sectors, "hospitality", "de")

# sectors - retail
retail_cur <- .safe(sectors, "retail", "cur")
retail_dm <- .safe(sectors, "retail", "dm")
retail_dy <- .safe(sectors, "retail", "dy")
retail_dc <- .safe(sectors, "retail", "dc")
retail_de <- .safe(sectors, "retail", "de")

# sectors - health
health_cur <- .safe(sectors, "health", "cur")
health_dm <- .safe(sectors, "health", "dm")
health_dy <- .safe(sectors, "health", "dy")
health_dc <- .safe(sectors, "health", "dc")
health_de <- .safe(sectors, "health", "de")

# hr1
hr1_cur <- .safe(hr1, "cur")
hr1_dm <- .safe(hr1, "dm")
hr1_dy <- .safe(hr1, "dy")
hr1_dc <- .safe(hr1, "dc")
hr1_de <- .safe(hr1, "de")

# inactivity driver text
inact_driver_text <- tryCatch(generate_inactivity_driver_text(inact_reasons), error = function(e) "")

#  labels (guarded against NULL results)

lfs_period_label <- if (!is.null(lfs) && !is.null(lfs$emp16) && !is.na(.safe_date(lfs, "emp16", "end"))) {
  lfs_label_narrative(lfs$emp16$end)
} else {
  paste0(format(anchor_m, "%B %Y"), " to ", format(anchor_m %m+% months(2), "%B %Y"))
}

lfs_period_short_label <- if (!is.null(lfs) && !is.null(lfs$emp16) && !is.na(.safe_date(lfs, "emp16", "end"))) {
  make_lfs_label(lfs$emp16$end)
} else {
  make_lfs_label(anchor_m %m+% months(2))
}

vacancies_period_short_label <- if (!is.null(vac) && "end" %in% names(vac) && !is.na(vac$end)) make_lfs_label(vac$end) else lfs_period_short_label

payroll_month_label <- if (!is.null(payroll) && !is.null(payroll$anchor) && !is.na(payroll$anchor)) {
  format(payroll$anchor, "%B %Y")
} else ""

payroll_flash_label <- if (!is.null(payroll) && !is.null(payroll$flash_anchor) && !is.na(payroll$flash_anchor)) {
  format(payroll$flash_anchor, "%B %Y")
} else ""

sector_month_label <- if (!is.null(sectors) && !is.null(sectors$hospitality) && !is.null(sectors$hospitality$anchor) && !is.na(sectors$hospitality$anchor)) {
  format(sectors$hospitality$anchor, "%B %Y")
} else ""

hr1_month_label <- if (!is.null(hr1) && !is.null(hr1$anchor) && !is.na(hr1$anchor)) {
  format(hr1$anchor, "%B %Y")
} else ""

# notable changes (statistically significant LFS movements)
notable_changes <- tryCatch(find_notable_changes(lfs), error = function(e) {
  warning("Notable changes calculation failed: ", e$message)
  data.frame(metric = character(), metric_label = character(),
             comparison = character(), comparison_label = character(),
             diff = numeric(), ci_lower = numeric(), ci_upper = numeric(),
             z_score = numeric(), narrative = character(),
             stringsAsFactors = FALSE)
})
