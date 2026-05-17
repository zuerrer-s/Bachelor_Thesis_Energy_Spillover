################################################################################
# Volatility stress environment, rolling TCI comparison, and TCI robustness checks
#
# Input:
#   vol_models_all_7models.rds
#
# Main logic:
#   Stress environment:
#     uses volatility levels from benchmark model
#
#   Rolling TCI and robustness:
#     uses transformed volatility inputs, consistent with directional CECI:
#       input_transform = "dlog_var": Δlog(h_t) = Δlog(σ_t^2)
#       input_transform = "dlog_vol": Δlog(σ_t)
#
# Outputs:
#   outputs/stress_environment_DCCfull.png
#   outputs/RollingTCI_ModelComparison_2Panel.png
#   outputs/RollingTCI_SummaryTable.csv
#   outputs/TCI_Robustness_3Panel.png
#   outputs/TCI_Robustness_Summary.csv
#   outputs/TCI_Robustness_Summary.tex
################################################################################

# ------------------------------------------------------------------------------
# 0) Libraries
# ------------------------------------------------------------------------------
suppressPackageStartupMessages({
  library(zoo)
  library(ConnectednessApproach)
})

# ------------------------------------------------------------------------------
# 1) User settings
# ------------------------------------------------------------------------------
base_dir <- "C:/Users/sezue/OneDrive/Desktop/BA"

output_dir <- file.path(base_dir, "outputs")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

vol_models_file <- file.path(output_dir, "vol_models_all_7models.rds")
if (!file.exists(vol_models_file)) {
  vol_models_file <- file.path(base_dir, "vol_models_all_7models.rds")
}

benchmark_model  <- "DCC_full"
robustness_model <- "DCC_full"

models_used <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym",
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage"
)

commodities4 <- c(
  "TTF (Gas)",
  "Brent (Oil)",
  "API2 (Coal)",
  "MO1 (Carbon)"
)

equity_regex <- "^SX"

# TCI input transform, consistent with the new directional CECI script:
#   "dlog_var" = Δlog(h_t) = Δlog(σ_t^2)
#   "dlog_vol" = Δlog(σ_t)
input_transform <- "dlog_var"

# Stress-environment settings
smooth_window_stress <- 5
smooth_window_vol    <- 5
stress_q             <- 0.90

# Rolling TCI settings
nlag_tci   <- 1
nfore_tci  <- 20
window_tci <- 200

# Robustness baseline and grids
baseline_p <- 1
baseline_H <- 20
baseline_W <- 200

H_grid <- c(5, 10, 20, 60)
p_grid <- c(1, 2, 3, 4)
W_grid <- c(100, 150, 200, 250)

# Stress episodes
stress_episodes <- data.frame(
  label = c(
    "2015–16 Equity Stress",
    "COVID-19",
    "Russia–Ukraine War",
    "Liberation Day"
  ),
  start = as.Date(c(
    "2015-08-01",
    "2020-01-01",
    "2021-10-01",
    "2025-02-01"
  )),
  end = as.Date(c(
    "2016-12-31",
    "2020-12-31",
    "2023-12-31",
    "2025-11-10"
  )),
  stringsAsFactors = FALSE
)

episode_shade_col <- gray(0.90, alpha = 0.95)

label_x_shift_days <- c(0, -35, 65, 15)
label_y_frac_from_top <- c(0.025, 0.025, 0.025, 0.025)

# Output files
out_stress_png <- file.path(output_dir, "stress_environment_DCCfull.png")
out_tci_png    <- file.path(output_dir, "RollingTCI_ModelComparison_2Panel.png")
out_tci_csv    <- file.path(output_dir, "RollingTCI_SummaryTable.csv")

out_robust_png <- file.path(output_dir, "TCI_Robustness_3Panel.png")
out_robust_csv <- file.path(output_dir, "TCI_Robustness_Summary.csv")
out_robust_tex <- file.path(output_dir, "TCI_Robustness_Summary.tex")

# ------------------------------------------------------------------------------
# 2) General helpers
# ------------------------------------------------------------------------------
save_png <- function(filename, expr, width = 2600, height = 1700, res = 200) {
  grDevices::png(
    filename = filename,
    width = width,
    height = height,
    res = res,
    type = "cairo"
  )
  
  tryCatch(
    eval.parent(substitute(expr)),
    finally = grDevices::dev.off()
  )
  
  invisible(filename)
}

ensure_date_index <- function(z) {
  stopifnot(inherits(z, "zoo"))
  
  idx <- zoo::index(z)
  
  if (inherits(idx, "Date")) return(z)
  
  if (inherits(idx, c("POSIXct", "POSIXt"))) {
    zoo::index(z) <- as.Date(idx)
    return(z)
  }
  
  if (inherits(idx, c("yearmon", "yearqtr"))) {
    zoo::index(z) <- as.Date(idx)
    return(z)
  }
  
  if (is.numeric(idx)) {
    zoo::index(z) <- as.Date(idx, origin = "1970-01-01")
    return(z)
  }
  
  idx2 <- suppressWarnings(as.Date(idx))
  
  if (all(is.na(idx2))) {
    stop(
      "Could not coerce zoo index to Date. Current class: ",
      paste(class(idx), collapse = ", ")
    )
  }
  
  zoo::index(z) <- idx2
  z
}

ensure_vol_colnames <- function(vol_zoo, model_name = "model") {
  stopifnot(inherits(vol_zoo, "zoo"))
  
  cn <- colnames(vol_zoo)
  
  if (is.null(cn)) {
    stop(model_name, ": volatility object has no column names.")
  }
  
  colnames(vol_zoo) <- ifelse(
    grepl("_vol$", cn),
    cn,
    paste0(cn, "_vol")
  )
  
  vol_zoo
}

strip_vol_suffix <- function(x) {
  sub("_vol$", "", x)
}

normalize_name <- function(x) {
  x <- tolower(x)
  gsub("[^a-z0-9]+", "", x)
}

match_base_name <- function(requested_base, available_base) {
  req_n <- normalize_name(requested_base)
  av_n  <- normalize_name(available_base)
  
  hit <- which(av_n == req_n)
  if (length(hit) >= 1) return(available_base[hit[1]])
  
  hit <- which(
    grepl(req_n, av_n, fixed = TRUE) |
      grepl(av_n, req_n, fixed = TRUE)
  )
  
  if (length(hit) >= 1) return(available_base[hit[1]])
  
  NA_character_
}

roll_mean_safe <- function(z, k = 5) {
  if (k <= 1) return(z)
  
  zoo::zoo(
    zoo::rollapply(
      z,
      width = k,
      FUN = function(x) mean(x, na.rm = TRUE),
      align = "right",
      fill = NA
    ),
    order.by = zoo::index(z)
  )
}

axis_years <- function(x_dates, cex_axis = 1.0) {
  x_dates <- as.Date(x_dates)
  
  d0 <- min(x_dates, na.rm = TRUE)
  d1 <- max(x_dates, na.rm = TRUE)
  
  yrs <- seq(
    from = as.Date(paste0(format(d0, "%Y"), "-01-01")),
    to   = as.Date(paste0(format(d1, "%Y"), "-01-01")),
    by   = "1 year"
  )
  
  yrs <- yrs[yrs >= d0 & yrs <= d1]
  
  if (length(yrs) == 0) {
    axis(1, cex.axis = cex_axis)
  } else {
    axis(1, at = yrs, labels = format(yrs, "%Y"), cex.axis = cex_axis)
  }
}

add_episode_shading <- function(episodes, col = gray(0.90, alpha = 0.95)) {
  usr <- par("usr")
  
  for (i in seq_len(nrow(episodes))) {
    rect(
      xleft   = episodes$start[i],
      ybottom = usr[3],
      xright  = episodes$end[i],
      ytop    = usr[4],
      col     = col,
      border  = NA
    )
  }
}

add_episode_labels_top <- function(episodes,
                                   x_shift_days = NULL,
                                   y_frac_from_top = NULL,
                                   cex = 0.92,
                                   col = "black") {
  n <- nrow(episodes)
  
  if (is.null(x_shift_days)) {
    x_shift_days <- rep(0, n)
  }
  
  if (is.null(y_frac_from_top)) {
    y_frac_from_top <- rep(0.025, n)
  }
  
  usr <- par("usr")
  y_span <- diff(usr[3:4])
  
  for (i in seq_len(n)) {
    x_mid <- episodes$start[i] +
      floor(as.numeric(episodes$end[i] - episodes$start[i]) / 2) +
      x_shift_days[i]
    
    y_lab <- usr[4] - y_frac_from_top[i] * y_span
    
    text(
      x = x_mid,
      y = y_lab,
      labels = episodes$label[i],
      cex = cex,
      col = col
    )
  }
}

align_common <- function(zlist) {
  stopifnot(is.list(zlist), length(zlist) >= 1)
  
  out <- do.call(merge, c(zlist, all = FALSE))
  ensure_date_index(out)
}

build_v_from_sigma <- function(sig_zoo, transform = c("dlog_var", "dlog_vol")) {
  transform <- match.arg(transform)
  
  sig_zoo <- ensure_date_index(sig_zoo)
  
  S <- zoo::coredata(sig_zoo)
  if (!is.matrix(S)) S <- as.matrix(S)
  
  S <- pmax(S, 1e-12)
  
  if (transform == "dlog_var") {
    H <- S^2
    V <- diff(log(H))
  } else {
    V <- diff(log(S))
  }
  
  zoo::zoo(V, order.by = zoo::index(sig_zoo)[-1])
}

# ------------------------------------------------------------------------------
# 3) Load and validate volatility models
# ------------------------------------------------------------------------------
if (!file.exists(vol_models_file)) {
  stop("Could not find volatility model file: ", vol_models_file)
}

vol_models <- readRDS(vol_models_file)

if (!is.list(vol_models) || length(vol_models) == 0) {
  stop("vol_models must be a non-empty named list of zoo objects.")
}

vol_models <- stats::setNames(
  lapply(names(vol_models), function(nm) {
    ensure_date_index(ensure_vol_colnames(vol_models[[nm]], nm))
  }),
  names(vol_models)
)

missing_tci_models <- setdiff(models_used, names(vol_models))

if (length(missing_tci_models) > 0) {
  stop("vol_models is missing: ", paste(missing_tci_models, collapse = ", "))
}

if (!(benchmark_model %in% names(vol_models))) {
  stop("benchmark_model not found in vol_models: ", benchmark_model)
}

if (!(robustness_model %in% names(vol_models))) {
  stop("robustness_model not found in vol_models: ", robustness_model)
}

# ------------------------------------------------------------------------------
# 4) Figure 1: volatility stress environment
# ------------------------------------------------------------------------------
build_stress_environment_data <- function(vol_models,
                                          benchmark_model,
                                          commodities4,
                                          equity_regex = "^SX",
                                          smooth_window_stress = 5,
                                          smooth_window_vol = 5,
                                          stress_q = 0.90) {
  vb <- vol_models[[benchmark_model]]
  vb <- ensure_date_index(ensure_vol_colnames(vb, benchmark_model))
  
  base_names <- strip_vol_suffix(colnames(vb))
  
  commodity_hits <- sapply(
    commodities4,
    match_base_name,
    available_base = base_names
  )
  
  if (any(is.na(commodity_hits))) {
    stop(
      "Could not match all four commodities in benchmark model.\n",
      "Requested: ", paste(commodities4, collapse = ", "), "\n",
      "Matched: ", paste(commodity_hits, collapse = ", ")
    )
  }
  
  commodity_cols <- paste0(commodity_hits, "_vol")
  
  eq_idx <- grep(equity_regex, base_names)
  
  if (length(eq_idx) == 0) {
    stop(
      "No equity columns matched regex ",
      shQuote(equity_regex),
      " in ",
      benchmark_model
    )
  }
  
  equity_cols <- paste0(base_names[eq_idx], "_vol")
  
  eq_mat  <- vb[, equity_cols, drop = FALSE]
  com_mat <- vb[, commodity_cols, drop = FALSE]
  
  eq_avg <- zoo::zoo(
    rowMeans(zoo::coredata(eq_mat), na.rm = TRUE),
    order.by = zoo::index(vb)
  )
  
  com_avg <- zoo::zoo(
    rowMeans(zoo::coredata(com_mat), na.rm = TRUE),
    order.by = zoo::index(vb)
  )
  
  s_tilde <- zoo::zoo(
    log1p(pmax(as.numeric(eq_avg) * as.numeric(com_avg), 0)),
    order.by = zoo::index(vb)
  )
  
  s_tilde_plot <- roll_mean_safe(s_tilde, smooth_window_stress)
  eq_plot      <- roll_mean_safe(eq_avg, smooth_window_vol)
  com_plot     <- roll_mean_safe(com_avg, smooth_window_vol)
  
  keep <- is.finite(as.numeric(s_tilde_plot)) &
    is.finite(as.numeric(eq_plot)) &
    is.finite(as.numeric(com_plot))
  
  s_tilde_plot <- s_tilde_plot[keep]
  eq_plot      <- eq_plot[keep]
  com_plot     <- com_plot[keep]
  
  q_stress <- as.numeric(
    stats::quantile(
      zoo::coredata(s_tilde_plot),
      probs = stress_q,
      na.rm = TRUE
    )
  )
  
  list(
    s_tilde = s_tilde_plot,
    equity_avg = eq_plot,
    commodity_avg = com_plot,
    q_stress = q_stress,
    equity_cols = equity_cols,
    commodity_cols = commodity_cols
  )
}

plot_stress_environment <- function(stress_data,
                                    benchmark_model,
                                    commodities4,
                                    episodes,
                                    label_x_shift_days,
                                    label_y_frac_from_top) {
  x <- zoo::index(stress_data$s_tilde)
  
  op <- par(no.readonly = TRUE)
  
  on.exit({
    layout(1)
    par(op)
  }, add = TRUE)
  
  layout(matrix(c(1, 2), nrow = 2), heights = c(10, 1.55))
  
  par(
    mar = c(3.2, 5.6, 3.8, 5.2) + 0.1,
    cex.axis = 1.18,
    cex.lab = 1.30,
    yaxs = "r"
  )
  
  y_left <- range(
    as.numeric(zoo::coredata(stress_data$s_tilde)),
    na.rm = TRUE
  )
  
  y_right <- range(
    c(
      as.numeric(zoo::coredata(stress_data$equity_avg)),
      as.numeric(zoo::coredata(stress_data$commodity_avg))
    ),
    na.rm = TRUE
  )
  
  plot(
    x,
    as.numeric(zoo::coredata(stress_data$s_tilde)),
    type = "n",
    xaxt = "n",
    xlab = "",
    ylab = expression(paste("Joint Stress, ", tilde(S)[t])),
    ylim = y_left,
    main = paste0("Volatility stress environment — benchmark ", benchmark_model),
    cex.main = 1.25
  )
  
  add_episode_shading(episodes, col = episode_shade_col)
  box()
  
  abline(
    h = stress_data$q_stress,
    lty = 2,
    lwd = 1.2,
    col = "gray45"
  )
  
  lines(
    x,
    as.numeric(zoo::coredata(stress_data$s_tilde)),
    lwd = 2.8,
    col = "black"
  )
  
  add_episode_labels_top(
    episodes = episodes,
    x_shift_days = label_x_shift_days,
    y_frac_from_top = label_y_frac_from_top,
    cex = 0.92,
    col = "black"
  )
  
  par(new = TRUE)
  
  plot(
    x,
    as.numeric(zoo::coredata(stress_data$equity_avg)),
    type = "n",
    axes = FALSE,
    xlab = "",
    ylab = "",
    ylim = y_right
  )
  
  lines(
    x,
    as.numeric(zoo::coredata(stress_data$equity_avg)),
    col = "dodgerblue3",
    lwd = 1.9,
    lty = 2
  )
  
  lines(
    x,
    as.numeric(zoo::coredata(stress_data$commodity_avg)),
    col = "red2",
    lwd = 1.9,
    lty = 1
  )
  
  axis(4, cex.axis = 1.18)
  
  mtext(
    expression(paste("DCC Volatility Benchmark, ", bar(sigma)[i])),
    side = 4,
    line = 3.1,
    cex = 1.30
  )
  
  axis_years(x, cex_axis = 1.18)
  
  mtext(
    paste0(
      "Benchmark source: ",
      benchmark_model,
      " | Commodity average: ",
      paste(commodities4, collapse = ", ")
    ),
    side = 3,
    line = 1.0,
    cex = 1.05
  )
  
  par(mar = c(0, 0, 0, 0))
  
  plot.new()
  plot.window(xlim = c(0, 1), ylim = c(0, 1))
  
  legend(
    x = 0.5,
    y = 0.5,
    xjust = 0.5,
    yjust = 0.5,
    legend = c(
      expression(tilde(S)[t]),
      "Equity avg volatility (SX*)",
      "Commodity avg volatility (4 commodities)",
      expression(paste("90th percentile of ", tilde(S)[t]))
    ),
    ncol = 2,
    bty = "n",
    cex = 1.00,
    lwd = c(2.8, 1.9, 1.9, 1.2),
    lty = c(1, 2, 1, 2),
    col = c("black", "dodgerblue3", "red2", "gray45"),
    seg.len = 2.6,
    x.intersp = 1.2,
    y.intersp = 1.2
  )
}

stress_data <- build_stress_environment_data(
  vol_models = vol_models,
  benchmark_model = benchmark_model,
  commodities4 = commodities4,
  equity_regex = equity_regex,
  smooth_window_stress = smooth_window_stress,
  smooth_window_vol = smooth_window_vol,
  stress_q = stress_q
)

save_png(
  out_stress_png,
  {
    plot_stress_environment(
      stress_data = stress_data,
      benchmark_model = benchmark_model,
      commodities4 = commodities4,
      episodes = stress_episodes,
      label_x_shift_days = label_x_shift_days,
      label_y_frac_from_top = label_y_frac_from_top
    )
  },
  width = 2600,
  height = 1550,
  res = 220
)

# ------------------------------------------------------------------------------
# 5) Rolling TCI comparison across volatility models
#    Uses transformed inputs: Δlog(h_t) or Δlog(σ_t)
# ------------------------------------------------------------------------------
default_cols <- function(models) {
  c(
    sBEKK_sym = "black",
    dBEKK_sym = "dodgerblue3",
    dBEKK_asym = "firebrick2",
    DCC_full = "darkgreen",
    DCC_scalar_stage = "purple3",
    cDCC_Aielli_stage = "goldenrod2"
  )[models]
}

default_ltys <- function(models) {
  c(
    sBEKK_sym = 1,
    dBEKK_sym = 2,
    dBEKK_asym = 4,
    DCC_full = 1,
    DCC_scalar_stage = 6,
    cDCC_Aielli_stage = 5
  )[models]
}

compute_rolling_tci_models <- function(vol_models,
                                       models_used,
                                       nlag,
                                       nfore,
                                       window_size,
                                       input_transform = c("dlog_var", "dlog_vol")) {
  input_transform <- match.arg(input_transform)
  
  tci_list <- list()
  
  common_cols <- Reduce(intersect, lapply(models_used, function(m) {
    colnames(vol_models[[m]])
  }))
  
  if (length(common_cols) < 2) {
    stop("Fewer than two common volatility columns across selected models.")
  }
  
  cat("\nCommon columns used for rolling TCI:\n")
  print(common_cols)
  
  cat("\nTCI input transform:", input_transform, "\n")
  
  for (m in models_used) {
    cat("Computing rolling TCI for:", m, "\n")
    
    sig <- vol_models[[m]][, common_cols, drop = FALSE]
    sig <- ensure_date_index(ensure_vol_colnames(sig, m))
    
    z <- build_v_from_sigma(sig, transform = input_transform)
    z <- z[stats::complete.cases(z), , drop = FALSE]
    
    if (NROW(z) <= window_size + nlag + 5) {
      stop(
        "Too few observations in ",
        m,
        " after transformation and cleaning for window_size=",
        window_size,
        ", nlag=",
        nlag
      )
    }
    
    ca <- ConnectednessApproach(
      x = z,
      nlag = nlag,
      nfore = nfore,
      window.size = window_size,
      model = "VAR",
      connectedness = "Time"
    )
    
    tci <- ca$TCI
    
    if (!inherits(tci, "zoo")) {
      idx <- tail(zoo::index(z), length(tci))
      tci <- zoo::zoo(tci, idx)
    }
    
    tci <- zoo::zoo(
      matrix(zoo::coredata(tci), ncol = 1),
      zoo::index(tci)
    )
    
    colnames(tci) <- m
    
    tci_list[[m]] <- tci
  }
  
  Z <- do.call(merge, c(tci_list, all = FALSE))
  Z <- ensure_date_index(Z)
  
  Z
}

plot_rolling_TCI_two_panel <- function(TCI_zoo,
                                       benchmark = "DCC_full",
                                       episodes = NULL,
                                       shade_col = gray(0.90, alpha = 0.95),
                                       label_x_shift_days = NULL,
                                       label_y_frac_from_top = NULL,
                                       title_suffix = NULL) {
  Z <- ensure_date_index(TCI_zoo)
  models <- colnames(Z)
  
  if (!(benchmark %in% models)) {
    stop("benchmark not found in TCI_zoo: ", benchmark)
  }
  
  D <- Z
  
  for (nm in models) {
    D[, nm] <- Z[, nm] - Z[, benchmark]
  }
  
  cols <- default_cols(models)
  ltys <- default_ltys(models)
  
  x <- zoo::index(Z)
  
  op <- par(no.readonly = TRUE)
  on.exit(par(op), add = TRUE)
  
  par(
    mfrow = c(2, 1),
    oma = c(9, 0, 3, 0),
    cex.axis = 1.25,
    cex.lab = 1.45
  )
  
  par(mar = c(0, 5.8, 0, 4.8))
  
  plot(
    x,
    Z[, 1],
    type = "n",
    ylim = c(0, 100),
    xaxt = "n",
    xlab = "",
    ylab = "TCI"
  )
  
  if (!is.null(episodes) && nrow(episodes) > 0) {
    add_episode_shading(episodes, col = shade_col)
  }
  
  box()
  
  main_title <- "Rolling TCI (window=200, nlag=1, nfore=20) — model comparison"
  if (!is.null(title_suffix)) {
    main_title <- paste0(main_title, " — ", title_suffix)
  }
  
  mtext(
    main_title,
    side = 3,
    outer = TRUE,
    line = 1,
    cex = 1.15
  )
  
  for (m in models) {
    if (m != benchmark) {
      lines(
        x,
        Z[, m],
        col = cols[m],
        lty = ltys[m],
        lwd = 1.35
      )
    }
  }
  
  lines(
    x,
    Z[, benchmark],
    col = cols[benchmark],
    lwd = 2.7
  )
  
  if (!is.null(episodes) && nrow(episodes) > 0) {
    add_episode_labels_top(
      episodes = episodes,
      x_shift_days = label_x_shift_days,
      y_frac_from_top = label_y_frac_from_top,
      cex = 0.92,
      col = "black"
    )
  }
  
  par(mar = c(4.5, 5.8, 0, 4.8))
  
  mdev <- max(abs(zoo::coredata(D)), na.rm = TRUE)
  
  plot(
    x,
    D[, 1],
    type = "n",
    ylim = c(-mdev, mdev),
    xaxt = "n",
    xlab = "",
    ylab = expression(Delta * "TCI")
  )
  
  if (!is.null(episodes) && nrow(episodes) > 0) {
    add_episode_shading(episodes, col = shade_col)
  }
  
  box()
  abline(h = 0, lty = 3)
  
  for (m in models) {
    if (m != benchmark) {
      lines(
        x,
        D[, m],
        col = cols[m],
        lty = ltys[m],
        lwd = 1.35
      )
    }
  }
  
  axis_years(x, cex_axis = 1.25)
  
  par(xpd = NA)
  
  legend(
    "bottom",
    inset = c(0, -0.55),
    legend = c(paste0(benchmark, " (benchmark)"), setdiff(models, benchmark)),
    col = c(cols[benchmark], cols[setdiff(models, benchmark)]),
    lty = c(ltys[benchmark], ltys[setdiff(models, benchmark)]),
    lwd = c(2.7, rep(1.35, length(models) - 1)),
    ncol = 3,
    bty = "n",
    cex = 1.15
  )
}

build_tci_summary_table <- function(TCI_zoo) {
  data.frame(
    Model = colnames(TCI_zoo),
    Mean = round(colMeans(TCI_zoo), 3),
    SD = round(apply(TCI_zoo, 2, sd), 3),
    Min = round(apply(TCI_zoo, 2, min), 3),
    Median = round(apply(TCI_zoo, 2, median), 3),
    Max = round(apply(TCI_zoo, 2, max), 3),
    stringsAsFactors = FALSE
  )
}

TCI_res <- compute_rolling_tci_models(
  vol_models = vol_models,
  models_used = models_used,
  nlag = nlag_tci,
  nfore = nfore_tci,
  window_size = window_tci,
  input_transform = input_transform
)

tci_table <- build_tci_summary_table(TCI_res)

write.csv(
  tci_table,
  out_tci_csv,
  row.names = FALSE
)

save_png(
  out_tci_png,
  {
    plot_rolling_TCI_two_panel(
      TCI_res,
      benchmark = benchmark_model,
      episodes = stress_episodes,
      shade_col = episode_shade_col,
      label_x_shift_days = label_x_shift_days,
      label_y_frac_from_top = label_y_frac_from_top,
      title_suffix = paste0("input: ", input_transform)
    )
  },
  width = 2600,
  height = 1700,
  res = 200
)

# ------------------------------------------------------------------------------
# 6) Robustness checks for H, p, and W
#    Uses transformed inputs: Δlog(h_t) or Δlog(σ_t)
# ------------------------------------------------------------------------------
compute_rolling_tci_single <- function(vol_zoo,
                                       model_name = "model",
                                       nlag = 1,
                                       nfore = 20,
                                       window_size = 200,
                                       input_transform = c("dlog_var", "dlog_vol")) {
  input_transform <- match.arg(input_transform)
  
  sig <- ensure_vol_colnames(vol_zoo, model_name)
  sig <- ensure_date_index(sig)
  
  z <- build_v_from_sigma(sig, transform = input_transform)
  z <- z[stats::complete.cases(z), , drop = FALSE]
  
  if (NROW(z) <= window_size + nlag + 5) {
    stop(
      "Too few observations in ",
      model_name,
      " after transformation and cleaning for window_size=",
      window_size,
      ", nlag=",
      nlag
    )
  }
  
  ca_obj <- ConnectednessApproach(
    x = z,
    nlag = nlag,
    nfore = nfore,
    window.size = window_size,
    model = "VAR",
    connectedness = "Time"
  )
  
  tci <- ca_obj$TCI
  
  if (!inherits(tci, "zoo")) {
    idx <- tail(zoo::index(z), length(tci))
    tci <- zoo::zoo(tci, idx)
  }
  
  tci <- zoo::zoo(
    matrix(zoo::coredata(tci), ncol = 1),
    zoo::index(tci)
  )
  
  colnames(tci) <- model_name
  
  tci
}

compute_spec_grid <- function(vol_zoo,
                              model_name,
                              family = c("H", "p", "W"),
                              grid_vals,
                              baseline_p,
                              baseline_H,
                              baseline_W,
                              input_transform = c("dlog_var", "dlog_vol")) {
  family <- match.arg(family)
  input_transform <- match.arg(input_transform)
  
  out <- list()
  
  for (g in grid_vals) {
    if (family == "H") {
      this_p <- baseline_p
      this_H <- g
      this_W <- baseline_W
      nm <- paste0("H=", g)
    } else if (family == "p") {
      this_p <- g
      this_H <- baseline_H
      this_W <- baseline_W
      nm <- paste0("p=", g)
    } else {
      this_p <- baseline_p
      this_H <- baseline_H
      this_W <- g
      nm <- paste0("W=", g)
    }
    
    message(
      "Computing ",
      family,
      " robustness spec: ",
      nm,
      " | input transform: ",
      input_transform
    )
    
    zz <- compute_rolling_tci_single(
      vol_zoo = vol_zoo,
      model_name = model_name,
      nlag = this_p,
      nfore = this_H,
      window_size = this_W,
      input_transform = input_transform
    )
    
    colnames(zz) <- nm
    out[[nm]] <- zz
  }
  
  align_common(out)
}

build_family_summary <- function(Z, baseline_name, family_name) {
  stopifnot(inherits(Z, "zoo"))
  stopifnot(baseline_name %in% colnames(Z))
  
  base <- as.numeric(zoo::coredata(Z[, baseline_name]))
  
  out <- data.frame(
    Family = family_name,
    Spec = colnames(Z),
    Mean = NA_real_,
    SD = NA_real_,
    Min = NA_real_,
    Max = NA_real_,
    Corr_with_baseline = NA_real_,
    MAD_vs_baseline = NA_real_,
    stringsAsFactors = FALSE
  )
  
  for (i in seq_along(colnames(Z))) {
    x <- as.numeric(zoo::coredata(Z[, i]))
    
    out$Mean[i] <- mean(x, na.rm = TRUE)
    out$SD[i]   <- sd(x, na.rm = TRUE)
    out$Min[i]  <- min(x, na.rm = TRUE)
    out$Max[i]  <- max(x, na.rm = TRUE)
    
    if (all(is.finite(base)) && all(is.finite(x))) {
      out$Corr_with_baseline[i] <- suppressWarnings(cor(x, base))
      out$MAD_vs_baseline[i]    <- mean(abs(x - base), na.rm = TRUE)
    }
  }
  
  num_cols <- setdiff(names(out), c("Family", "Spec"))
  out[num_cols] <- lapply(out[num_cols], function(x) round(x, 3))
  
  out
}

write_latex_table <- function(df, file) {
  con <- file(file, open = "wt")
  on.exit(close(con), add = TRUE)
  
  writeLines("\\begin{table}[H]", con)
  writeLines("\\centering", con)
  writeLines(
    paste0(
      "\\caption{Robustness of the rolling TCI with respect to $H$, $p$, and $W$",
      " using input transform ",
      input_transform,
      "}"
    ),
    con
  )
  writeLines("\\label{tab:tci_robustness}", con)
  writeLines("\\begin{tabular}{llrrrrrr}", con)
  writeLines("\\hline", con)
  writeLines("Family & Spec & Mean & SD & Min & Max & Corr. base & MAD base\\\\", con)
  writeLines("\\hline", con)
  
  for (i in seq_len(nrow(df))) {
    line <- paste(
      df$Family[i],
      "&",
      gsub("_", "\\\\_", df$Spec[i]),
      "&",
      sprintf("%.3f", df$Mean[i]),
      "&",
      sprintf("%.3f", df$SD[i]),
      "&",
      sprintf("%.3f", df$Min[i]),
      "&",
      sprintf("%.3f", df$Max[i]),
      "&",
      sprintf("%.3f", df$Corr_with_baseline[i]),
      "&",
      sprintf("%.3f", df$MAD_vs_baseline[i]),
      "\\\\"
    )
    
    writeLines(line, con)
  }
  
  writeLines("\\hline", con)
  writeLines("\\end{tabular}", con)
  writeLines("\\end{table}", con)
}

simple_palette <- function(n) {
  base <- c(
    "black",
    "dodgerblue3",
    "firebrick2",
    "darkgreen",
    "purple3",
    "goldenrod2"
  )
  
  rep(base, length.out = n)
}

plot_family_panel <- function(Z,
                              panel_title,
                              baseline_name,
                              ylab = "TCI",
                              bottom_axis = FALSE,
                              legend_below = FALSE) {
  cols <- simple_palette(NCOL(Z))
  ltys <- rep(1:6, length.out = NCOL(Z))
  
  names(cols) <- colnames(Z)
  names(ltys) <- colnames(Z)
  
  x <- zoo::index(Z)
  
  if (bottom_axis) {
    par(mar = c(4.4, 5.4, 0, 3.5) + 0.1)
  } else {
    par(mar = c(0, 5.4, 0, 3.5) + 0.1)
  }
  
  yl <- range(zoo::coredata(Z), finite = TRUE)
  
  plot(
    x,
    zoo::coredata(Z[, 1]),
    type = "n",
    xlab = "",
    ylab = ylab,
    xaxt = "n",
    ylim = yl,
    main = NULL
  )
  
  mtext(panel_title, side = 3, line = 0.2, cex = 0.95)
  
  for (j in seq_len(NCOL(Z))) {
    nm <- colnames(Z)[j]
    lw <- if (nm == baseline_name) 2.3 else 1.4
    
    lines(
      x,
      zoo::coredata(Z[, j]),
      col = cols[nm],
      lty = ltys[nm],
      lwd = lw
    )
  }
  
  if (bottom_axis) {
    axis_years(x, cex_axis = 1.15)
  }
  
  if (legend_below) {
    par(xpd = NA)
    
    legend(
      "bottom",
      inset = c(0, -0.30),
      legend = colnames(Z),
      col = cols[colnames(Z)],
      lty = ltys[colnames(Z)],
      lwd = ifelse(colnames(Z) == baseline_name, 2.3, 1.4),
      bty = "n",
      ncol = 2,
      cex = 1.00
    )
  }
}

plot_robustness_3panel <- function(H_Z,
                                   p_Z,
                                   W_Z,
                                   baseline_H_name,
                                   baseline_p_name,
                                   baseline_W_name,
                                   main) {
  op <- par(no.readonly = TRUE)
  on.exit(par(op), add = TRUE)
  
  par(
    mfrow = c(3, 1),
    oma = c(8.5, 0, 3.2, 0),
    cex.axis = 1.15,
    cex.lab = 1.25
  )
  
  plot_family_panel(
    Z = H_Z,
    panel_title = expression(paste("Forecast horizon robustness (", H, ")")),
    baseline_name = baseline_H_name,
    ylab = "TCI",
    bottom_axis = FALSE,
    legend_below = FALSE
  )
  
  plot_family_panel(
    Z = p_Z,
    panel_title = expression(paste("Lag-order robustness (", p, ")")),
    baseline_name = baseline_p_name,
    ylab = "TCI",
    bottom_axis = FALSE,
    legend_below = FALSE
  )
  
  plot_family_panel(
    Z = W_Z,
    panel_title = expression(paste("Window-size robustness (", W, ")")),
    baseline_name = baseline_W_name,
    ylab = "TCI",
    bottom_axis = TRUE,
    legend_below = TRUE
  )
  
  mtext(main, side = 3, outer = TRUE, line = 1, cex = 1.10)
}

vol_z <- vol_models[[robustness_model]]

H_Z <- compute_spec_grid(
  vol_zoo = vol_z,
  model_name = robustness_model,
  family = "H",
  grid_vals = H_grid,
  baseline_p = baseline_p,
  baseline_H = baseline_H,
  baseline_W = baseline_W,
  input_transform = input_transform
)

baseline_H_name <- paste0("H=", baseline_H)

p_Z <- compute_spec_grid(
  vol_zoo = vol_z,
  model_name = robustness_model,
  family = "p",
  grid_vals = p_grid,
  baseline_p = baseline_p,
  baseline_H = baseline_H,
  baseline_W = baseline_W,
  input_transform = input_transform
)

baseline_p_name <- paste0("p=", baseline_p)

W_Z <- compute_spec_grid(
  vol_zoo = vol_z,
  model_name = robustness_model,
  family = "W",
  grid_vals = W_grid,
  baseline_p = baseline_p,
  baseline_H = baseline_H,
  baseline_W = baseline_W,
  input_transform = input_transform
)

baseline_W_name <- paste0("W=", baseline_W)

tab_H <- build_family_summary(H_Z, baseline_H_name, "H")
tab_p <- build_family_summary(p_Z, baseline_p_name, "p")
tab_W <- build_family_summary(W_Z, baseline_W_name, "W")

robust_tab <- rbind(tab_H, tab_p, tab_W)

write.csv(
  robust_tab,
  out_robust_csv,
  row.names = FALSE
)

write_latex_table(
  robust_tab,
  out_robust_tex
)

save_png(
  out_robust_png,
  {
    plot_robustness_3panel(
      H_Z = H_Z,
      p_Z = p_Z,
      W_Z = W_Z,
      baseline_H_name = baseline_H_name,
      baseline_p_name = baseline_p_name,
      baseline_W_name = baseline_W_name,
      main = paste0(
        "Rolling TCI robustness — model: ",
        robustness_model,
        " | input: ",
        input_transform,
        " (baseline: p=",
        baseline_p,
        ", H=",
        baseline_H,
        ", W=",
        baseline_W,
        ")"
      )
    )
  },
  width = 2400,
  height = 2200,
  res = 200
)

# ------------------------------------------------------------------------------
# 7) Summary
# ------------------------------------------------------------------------------
cat("\n=============================\n")
cat("FIGURES AND ROBUSTNESS CHECKS COMPLETE\n")
cat("=============================\n")

cat("Input file:\n", vol_models_file, "\n\n", sep = "")

cat("Connectedness input transform for TCI:\n")
cat(" - ", input_transform, "\n\n", sep = "")

cat("Saved outputs:\n")
cat(" - ", out_stress_png, "\n", sep = "")
cat(" - ", out_tci_png, "\n", sep = "")
cat(" - ", out_tci_csv, "\n", sep = "")
cat(" - ", out_robust_png, "\n", sep = "")
cat(" - ", out_robust_csv, "\n", sep = "")
cat(" - ", out_robust_tex, "\n\n", sep = "")

cat("Stress-environment inputs:\n")
cat("Benchmark model:", benchmark_model, "\n")
cat("Equity columns:", length(stress_data$equity_cols), "\n")
cat("Commodity columns:", paste(stress_data$commodity_cols, collapse = ", "), "\n")
cat("Stress quantile:", round(stress_data$q_stress, 4), "\n\n")

cat("Rolling TCI summary:\n")
print(tci_table)

cat("\nRobustness summary:\n")
print(robust_tab)

################################################################################
# END
################################################################################