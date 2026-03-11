# lfs module - a01 sheet 2 (employment, unemployment, inactivity)

# config


LFS_CODES <- list(
  EMP16 = "MGRZ",
  EMP_RT = "LF24",
  UNEMP16 = "MGSC",
  UNEMP_RT = "MGSX",
  INACT = "LF2M",
  INACT_RT = "LF2S",
  INACT5064 = "LF2A",
  INACT5064_RT = "LF2W"
)


# fetch


fetch_lfs <- function() {
  conn <- DBI::dbConnect(RPostgres::Postgres())
  
  tryCatch({
    query <- 'SELECT time_period, dataset_identifier_code, value
FROM "ons"."labour_market__age_group"'
    result <- DBI::dbGetQuery(conn, query)
    tibble::as_tibble(result)
  },
  error = function(e) {
    warning("Failed to fetch LFS data: ", e$message)
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


# compute


compute_lfs_metric <- function(pg_data,
                               manual_mm,
                               code,
                               covid_label = COVID_LFS_LABEL,
                               election_label = ELECTION_LABEL) {
  cm <- parse_manual_month(manual_mm)
  
  end_cur <- cm %m-% months(2)
  end_q <- end_cur %m-% months(3)
  end_y <- end_cur %m-% months(12)
  
  lab_cur <- make_lfs_label(end_cur)
  lab_q <- make_lfs_label(end_q)
  lab_y <- make_lfs_label(end_y)
  
  cur <- val_by_code(pg_data, code, lab_cur)
  val_q <- val_by_code(pg_data, code, lab_q)
  val_y <- val_by_code(pg_data, code, lab_y)
  val_c <- val_by_code(pg_data, code, covid_label)
  val_e <- val_by_code(pg_data, code, election_label)
  
  dq <- if (!is.na(cur) && !is.na(val_q)) cur - val_q else NA_real_
  dy <- if (!is.na(cur) && !is.na(val_y)) cur - val_y else NA_real_
  dc <- if (!is.na(cur) && !is.na(val_c)) cur - val_c else NA_real_
  de <- if (!is.na(cur) && !is.na(val_e)) cur - val_e else NA_real_
  
  list(cur = cur, dq = dq, dy = dy, dc = dc, de = de, end = end_cur)
}


# calculate all lfs metrics


calculate_lfs <- function(manual_mm) {
  pg_data <- fetch_lfs()
  
  list(
    emp16 = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$EMP16),
    emp_rt = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$EMP_RT),
    unemp16 = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$UNEMP16),
    unemp_rt = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$UNEMP_RT),
    inact = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$INACT),
    inact_rt = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$INACT_RT),
    inact5064 = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$INACT5064),
    inact5064_rt = compute_lfs_metric(pg_data, manual_mm, LFS_CODES$INACT5064_RT)
  )
}
