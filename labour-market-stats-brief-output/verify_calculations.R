# verify_calculations.R
# Verification script for calculations_from_excel.R
#
# Usage:
#   Rscript verify_calculations.R <manual_month> <path_to_a01> [path_to_hr1] [path_to_x09] [path_to_rtisa]
#
# Example:
#   Rscript verify_calculations.R feb2026 data/a01feb2026.xlsx data/hr1.xlsx data/x09.xlsx data/rtisa.xlsx
#
# Or source interactively and call:
#   results <- verify_all("feb2026", file_a01 = "path/to/a01.xlsx", ...)

suppressPackageStartupMessages({
  library(readxl)
  library(lubridate)
})

# Source dependencies
source("utils/helpers.R")
source("utils/calculations_from_excel.R")

# KNOWN-GOOD VALUES
# 

known_good_feb2026 <- list(
  # LFS headline rates (Sheet 1) — from published briefing
  emp_rt_cur       = 75.0,     # Employment rate 16-64
  unemp_rt_cur     = 4.4,      # Unemployment rate 16+
  inact_rt_cur     = 21.6,     # Inactivity rate 16-64

  

  # Wages nominal (Sheet 13/15)
  latest_wages         = 5.9,   # Total pay YoY % (from top ten line 1)
  latest_regular_cash  = 5.6,   # Regular pay YoY % (from top ten line 1)

  # Wages real CPI (X09)
  latest_wages_cpi     = 2.2,   # Real total pay YoY %
  latest_regular_cpi   = 2.0,   # Real regular pay YoY %

  # Public/private wages (Sheet 13/15) — from top ten line 1
  # wages_total_public  = 7.0,  # (uncomment when verified)
  # wages_reg_public    = 7.2,  # (uncomment when verified)

  # Vacancies (Sheet 19, in thousands)
  vac_cur = 818                 # Vacancy level (thousands)
)

# 
# TOLERANCE 

TOLERANCE_RATE  <- 0.15   # ± for percentage rates (allows 75.0 vs 74.9)
TOLERANCE_LEVEL <- 5      # ± for levels in thousands
TOLERANCE_PCT   <- 0.15   # ± for YoY % growth figures

#
# VERIFICATION 

verify_metric <- function(name, actual, expected, tolerance, unit = "") {
  if (is.na(actual)) {
    list(name = name, status = "FAIL", reason = "NA (not extracted)",
         actual = NA, expected = expected, unit = unit)
  } else if (abs(actual - expected) <= tolerance) {
    list(name = name, status = "PASS", reason = "",
         actual = actual, expected = expected, unit = unit)
  } else {
    list(name = name, status = "FAIL",
         reason = sprintf("off by %.4f (tolerance ±%.4f)", actual - expected, tolerance),
         actual = actual, expected = expected, unit = unit)
  }
}

verify_all <- function(manual_month, file_a01 = NULL, file_hr1 = NULL,
                       file_x09 = NULL, file_rtisa = NULL,
                       known_good = known_good_feb2026) {

  # Run calculations into a fresh environment
  calc_env <- new.env(parent = globalenv())
  run_calculations_from_excel(
    manual_month = manual_month,
    file_a01     = file_a01,
    file_hr1     = file_hr1,
    file_x09     = file_x09,
    file_rtisa   = file_rtisa,
    target_env   = calc_env
  )

  results <- list()

  # get variable from calc_env safely
  sv <- function(name) {
    if (exists(name, envir = calc_env)) get(name, envir = calc_env) else NA_real_
  }

  # Check  known-good value
  for (name in names(known_good)) {
    expected <- known_good[[name]]
    actual   <- sv(name)

    #  tolerance based on metric type
    tol <- if (grepl("_rt_|_rate", name)) {
      TOLERANCE_RATE
    } else if (grepl("wages|regular|cpi", name)) {
      TOLERANCE_PCT
    } else if (grepl("_cur$|_dq$|_dy$", name)) {
      TOLERANCE_LEVEL
    } else {
      TOLERANCE_PCT
    }

    unit <- if (grepl("_rt_", name)) "%" else if (grepl("wages|cpi|regular", name)) "%" else ""

    results[[name]] <- verify_metric(name, actual, expected, tol, unit)
  }

  # Print results
  cat("\n========================================\n")
  cat("VERIFICATION REPORT\n")
  cat(sprintf("manual_month: %s\n", manual_month))
  cat(sprintf("Files: A01=%s HR1=%s X09=%s RTISA=%s\n",
              if (is.null(file_a01)) "NONE" else basename(file_a01),
              if (is.null(file_hr1)) "NONE" else basename(file_hr1),
              if (is.null(file_x09)) "NONE" else basename(file_x09),
              if (is.null(file_rtisa)) "NONE" else basename(file_rtisa)))
  cat("========================================\n\n")

  pass_count <- 0
  fail_count <- 0

  for (r in results) {
    status_icon <- if (r$status == "PASS") "OK" else "XX"
    actual_str  <- if (is.na(r$actual)) "NA" else sprintf("%.2f%s", r$actual, r$unit)
    expect_str  <- sprintf("%.2f%s", r$expected, r$unit)

    cat(sprintf("[%s] %-25s  actual=%-12s expected=%-12s %s\n",
                status_icon, r$name, actual_str, expect_str, r$reason))

    if (r$status == "PASS") pass_count <- pass_count + 1
    else fail_count <- fail_count + 1
  }

  cat(sprintf("\n--- %d PASS, %d FAIL out of %d checks ---\n\n",
              pass_count, fail_count, length(results)))

  # Also print all computed variables for manual inspection
  cat("ALL COMPUTED VARIABLES:\n")
  cat("----------------------------------------\n")
  all_vars <- ls(envir = calc_env)
  # Sort and print non-list variables
  for (v in sort(all_vars)) {
    val <- get(v, envir = calc_env)
    if (is.numeric(val) && length(val) == 1) {
      cat(sprintf("  %-30s = %s\n", v, if (is.na(val)) "NA" else format(val, digits = 6)))
    } else if (is.character(val) && length(val) == 1) {
      cat(sprintf("  %-30s = \"%s\"\n", v, val))
    }
  }
  cat("\n")

  invisible(list(results = results, calc_env = calc_env,
                 pass = pass_count, fail = fail_count))
}

# =============================================================================
# CLI ENTRY POINT
# =============================================================================

if (!interactive()) {
  args <- commandArgs(trailingOnly = TRUE)
  if (length(args) < 2) {
    cat("Usage: Rscript verify_calculations.R <manual_month> <a01_path> [hr1_path] [x09_path] [rtisa_path]\n")
    quit(status = 1)
  }

  mm       <- args[1]
  file_a01 <- args[2]
  file_hr1 <- if (length(args) >= 3) args[3] else NULL
  file_x09 <- if (length(args) >= 4) args[4] else NULL
  file_rtisa <- if (length(args) >= 5) args[5] else NULL

  out <- verify_all(mm, file_a01, file_hr1, file_x09, file_rtisa)

  if (out$fail > 0) {
    quit(status = 1)
  }
}
