# significance.R — ONS-standard confidence intervals and notable change detection
#
# ONS gold standard: 95% CIs (z >= 1.96, p < 0.05).
# The definitive test is calculating the CI around the difference itself
# and checking whether it excludes zero.
#
# SEs are approximate values from ONS published sampling variability tables
# for LFS 3-month rolling estimates. For overlapping periods the assumption
# of independence overstates SE_diff slightly (conservative).

# --- Standard error lookup (approximate, from ONS A01 supplementary data) ---

LFS_SE <- list(

  # Levels (thousands)
  emp16      = 100,   # Employment 16+ level
  unemp16    =  40,   # Unemployment 16+ level
  inact      =  80,   # Inactivity 16-64 level
  inact5064  =  45,   # Inactivity 50-64 level


  # Rates (percentage points)
  emp_rt      = 0.2,  # Employment rate 16-64
  unemp_rt    = 0.1,  # Unemployment rate 16+
  inact_rt    = 0.2,  # Inactivity rate 16-64
  inact5064_rt = 0.3  # Inactivity rate 50-64
)

# Human-readable metric names
LFS_METRIC_LABELS <- list(
  emp16        = "Employment level (16+)",
  emp_rt       = "Employment rate (16-64)",
  unemp16      = "Unemployment level (16+)",
  unemp_rt     = "Unemployment rate (16+)",
  inact        = "Inactivity level (16-64)",
  inact_rt     = "Inactivity rate (16-64)",
  inact5064    = "Inactivity level (50-64)",
  inact5064_rt = "Inactivity rate (50-64)"
)

# Comparison period labels
COMPARISON_LABELS <- list(
  dq = "on the quarter",
  dy = "on the year",
  dc = "since COVID (Dec-Feb 2020)",
  de = "since the election (Apr-Jun 2024)"
)

# Whether the metric is a rate (pp) or level (thousands)
is_rate_metric <- function(metric_name) {
  grepl("_rt$", metric_name)
}


# --- Core CI functions ---

#' SE of the difference between two independent estimates
#' For overlapping LFS periods this is conservative (overstates SE)
se_of_difference <- function(se_a, se_b) {
  sqrt(se_a^2 + se_b^2)
}

#' Test whether a change is statistically significant at 95% level
#' Returns a list with: significant (logical), z_score, ci_lower, ci_upper
test_significance <- function(diff, se_a, se_b, z_crit = 1.96) {
  if (is.na(diff) || is.na(se_a) || is.na(se_b)) {
    return(list(significant = FALSE, z_score = NA_real_,
                ci_lower = NA_real_, ci_upper = NA_real_))
  }
  se_diff <- se_of_difference(se_a, se_b)
  if (se_diff == 0) {
    return(list(significant = FALSE, z_score = NA_real_,
                ci_lower = NA_real_, ci_upper = NA_real_))
  }
  z <- abs(diff) / se_diff
  ci_lower <- diff - z_crit * se_diff
  ci_upper <- diff + z_crit * se_diff
  list(
    significant = z >= z_crit,
    z_score     = z,
    ci_lower    = ci_lower,
    ci_upper    = ci_upper
  )
}


# --- Notable change detection ---

#' Scan all LFS metrics x comparison periods, return significant changes
#' ranked by |z-score| (most notable first)
#'
#' @param lfs_results The lfs list from calculate_lfs(), containing sublists
#'   like emp16, emp_rt, etc., each with cur/dq/dy/dc/de
#' @param max_items Maximum number of notable changes to return (default 10)
#' @return A data.frame with columns: metric, metric_label, comparison,
#'   comparison_label, diff, ci_lower, ci_upper, z_score, narrative
find_notable_changes <- function(lfs_results, max_items = 10) {

  if (is.null(lfs_results)) {
    return(data.frame(
      metric = character(), metric_label = character(),
      comparison = character(), comparison_label = character(),
      diff = numeric(), ci_lower = numeric(), ci_upper = numeric(),
      z_score = numeric(), narrative = character(),
      stringsAsFactors = FALSE
    ))
  }

  metric_names <- names(LFS_SE)
  comparisons  <- c("dq", "dy", "dc", "de")

  rows <- list()

  for (m in metric_names) {
    metric_data <- lfs_results[[m]]
    if (is.null(metric_data)) next

    se_single <- LFS_SE[[m]]

    for (comp in comparisons) {
      diff_val <- metric_data[[comp]]
      if (is.null(diff_val) || is.na(diff_val)) next

      # Both estimates have the same SE (same survey design)
      result <- test_significance(diff_val, se_single, se_single)
      if (!result$significant) next

      # Build narrative
      is_rate <- is_rate_metric(m)
      label   <- LFS_METRIC_LABELS[[m]]
      comp_label <- COMPARISON_LABELS[[comp]]

      if (is_rate) {
        dir_word <- if (diff_val > 0) "rose" else "fell"
        diff_fmt <- paste0(format(round(abs(diff_val), 1), nsmall = 1), "pp")
        ci_fmt   <- paste0(
          format(round(result$ci_lower, 1), nsmall = 1), "pp to ",
          format(round(result$ci_upper, 1), nsmall = 1), "pp"
        )
      } else {
        dir_word <- if (diff_val > 0) "rose" else "fell"
        diff_fmt <- paste0(format(round(abs(diff_val), 0), big.mark = ","), " thousand")
        ci_fmt   <- paste0(
          format(round(result$ci_lower, 0), big.mark = ","), " to ",
          format(round(result$ci_upper, 0), big.mark = ","), " thousand"
        )
      }

      narrative <- paste0(
        label, " ", dir_word, " by ", diff_fmt, " ", comp_label,
        " (95% CI: ", ci_fmt, ")"
      )

      rows[[length(rows) + 1]] <- data.frame(
        metric           = m,
        metric_label     = label,
        comparison       = comp,
        comparison_label = comp_label,
        diff             = diff_val,
        ci_lower         = result$ci_lower,
        ci_upper         = result$ci_upper,
        z_score          = result$z_score,
        narrative        = narrative,
        stringsAsFactors = FALSE
      )
    }
  }

  if (length(rows) == 0) {
    return(data.frame(
      metric = character(), metric_label = character(),
      comparison = character(), comparison_label = character(),
      diff = numeric(), ci_lower = numeric(), ci_upper = numeric(),
      z_score = numeric(), narrative = character(),
      stringsAsFactors = FALSE
    ))
  }

  df <- do.call(rbind, rows)
  df <- df[order(-abs(df$z_score)), ]
  if (nrow(df) > max_items) df <- df[seq_len(max_items), ]
  rownames(df) <- NULL
  df
}
