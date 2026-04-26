################################################################################
# Pre-modeling diagnostics and innovation extraction
# Daily log returns: energy, carbon, and European sector indices
################################################################################

# ------------------------------------------------------------------------------
# 0) Libraries
# ------------------------------------------------------------------------------
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(scales)
  library(tseries)
  library(vars)
  library(FinTS)
})

# ------------------------------------------------------------------------------
# 1) User settings
# ------------------------------------------------------------------------------
base_dir  <- "C:/Users/sezue/OneDrive/Desktop/BA"
file_path <- file.path(base_dir, "log_returns2.xlsx")

output_dir <- file.path(base_dir, "outputs")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

ret_cols <- c(
  "TTF_ret", "Brent_ret", "API21_ret", "MO1_ret",
  "SXEP_ret", "SXQP_ret", "SXNP_ret", "SX6P_ret", "SX7P_ret"
)

series_order <- c(
  "TTF (Gas)",
  "Brent (Oil)",
  "API2 (Coal)",
  "MO1 (Carbon)",
  "SXEP (Oil & Gas)",
  "SXQP (Consumer)",
  "SXNP (Industrials)",
  "SX6P (Utilities)",
  "SX7P (Financials)"
)

stopifnot(length(ret_cols) == length(series_order))

LB_LAGS     <- 10
ARCH_LAGS   <- 12
ACF_LAG_MAX <- 20
VAR_LAG_MAX <- 5
VAR_P       <- 1

# ------------------------------------------------------------------------------
# 2) Helper functions
# ------------------------------------------------------------------------------
make_acf_df <- function(x, lag_max, series_order) {
  x <- as.matrix(x)
  ci <- 1.96 / sqrt(nrow(x))
  
  lapply(seq_len(ncol(x)), function(j) {
    ac <- acf(x[, j], lag.max = lag_max, plot = FALSE)
    data.frame(
      Series = colnames(x)[j],
      Lag    = as.numeric(ac$lag),
      ACF    = as.numeric(ac$acf),
      CI     = ci,
      stringsAsFactors = FALSE
    )
  }) |>
    bind_rows() |>
    filter(Lag >= 1) |>
    mutate(Series = factor(Series, levels = series_order))
}

make_acf_plot <- function(acf_data, ncol_facets = 3) {
  ggplot(acf_data, aes(x = Lag, y = ACF)) +
    geom_ribbon(aes(ymin = -CI, ymax = CI), fill = "steelblue", alpha = 0.18) +
    geom_hline(yintercept = 0, linewidth = 0.3) +
    geom_segment(aes(xend = Lag, yend = 0), linewidth = 0.35) +
    facet_wrap(~ Series, ncol = ncol_facets) +
    scale_x_continuous(
      breaks = seq(1, max(acf_data$Lag), by = 5),
      limits = c(1, max(acf_data$Lag))
    ) +
    labs(title = NULL, x = "Lag", y = "ACF") +
    theme_bw(base_size = 12) +
    theme(
      strip.text = element_text(face = "bold"),
      panel.grid.minor = element_blank()
    )
}

ljung_box_table <- function(x, lags, squared = FALSE) {
  x <- as.matrix(x)
  
  out <- data.frame(
    Series = colnames(x),
    Q_stat = NA_real_,
    p_value = NA_real_,
    Reject_5pct = NA,
    stringsAsFactors = FALSE
  )
  
  for (j in seq_len(ncol(x))) {
    test_series <- if (squared) x[, j]^2 else x[, j]
    bt <- Box.test(test_series, lag = lags, type = "Ljung-Box")
    
    out$Q_stat[j] <- unname(bt$statistic)
    out$p_value[j] <- unname(bt$p.value)
    out$Reject_5pct[j] <- bt$p.value < 0.05
  }
  
  out
}

arch_lm_table <- function(x, lags) {
  x <- as.matrix(x)
  
  out <- data.frame(
    Series = colnames(x),
    Q_stat = NA_real_,
    p_value = NA_real_,
    Reject_5pct = NA,
    stringsAsFactors = FALSE
  )
  
  for (j in seq_len(ncol(x))) {
    at <- ArchTest(x[, j], lags = lags)
    
    out$Q_stat[j] <- unname(at$statistic)
    out$p_value[j] <- unname(at$p.value)
    out$Reject_5pct[j] <- at$p.value < 0.05
  }
  
  out
}

pick_min_lag <- function(x) {
  as.integer(names(which.min(x)))
}

save_table <- function(x, filename) {
  write.csv(x, file.path(output_dir, filename), row.names = FALSE)
}

# ------------------------------------------------------------------------------
# 3) Read and prepare data
# ------------------------------------------------------------------------------
df <- read_excel(file_path)
df$Date <- as.Date(df$Date)

dat <- df |>
  dplyr::select(Date, dplyr::all_of(ret_cols)) |>
  dplyr::filter(stats::complete.cases(dplyr::across(dplyr::all_of(ret_cols))))

returns <- as.matrix(dat[, ret_cols])
dates <- dat$Date
colnames(returns) <- series_order

Tn <- nrow(returns)
K  <- ncol(returns)

cat("Data loaded successfully.\n")
cat("Observations:", Tn, "\n")
cat("Series:", K, "\n\n")

# ------------------------------------------------------------------------------
# 4) Return and squared-return plots
# ------------------------------------------------------------------------------
ret_df <- data.frame(Date = dates, returns, check.names = FALSE) |>
  tidyr::pivot_longer(-Date, names_to = "Series", values_to = "Return") |>
  dplyr::mutate(
    Series = factor(Series, levels = series_order),
    SqReturn = Return^2
  )

p_returns <- ggplot(ret_df, aes(x = Date, y = Return)) +
  geom_hline(yintercept = 0, linewidth = 0.25) +
  geom_line(linewidth = 0.25) +
  facet_wrap(~ Series, ncol = 3, scales = "free_y") +
  scale_x_date(date_breaks = "2 years", date_labels = "%Y") +
  labs(title = NULL, x = NULL, y = "Log Returns (%)") +
  theme_bw(base_size = 12) +
  theme(
    strip.text = element_text(face = "bold"),
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank(),
    axis.text.x = element_text(size = 9),
    axis.text.y = element_text(size = 9)
  )

p_squared_returns <- ggplot(ret_df, aes(x = Date, y = SqReturn)) +
  geom_line(linewidth = 0.25) +
  facet_wrap(~ Series, ncol = 3, scales = "free_y") +
  scale_x_date(date_breaks = "2 years", date_labels = "%Y") +
  labs(title = NULL, x = NULL, y = expression(r[t]^2)) +
  theme_bw(base_size = 12) +
  theme(
    strip.text = element_text(face = "bold"),
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank(),
    axis.text.x = element_text(size = 9),
    axis.text.y = element_text(size = 9)
  )

print(p_returns)
print(p_squared_returns)

ggsave(file.path(output_dir, "Fig_Returns_10x3.png"), p_returns, width = 10, height = 7.5, dpi = 300)
ggsave(file.path(output_dir, "Fig_Returns_10x3.pdf"), p_returns, width = 10, height = 7.5)
ggsave(file.path(output_dir, "Fig_SquaredReturns_10x3.png"), p_squared_returns, width = 10, height = 7.5, dpi = 300)
ggsave(file.path(output_dir, "Fig_SquaredReturns_10x3.pdf"), p_squared_returns, width = 10, height = 7.5)

# ------------------------------------------------------------------------------
# 5) Stationarity and distribution diagnostics
# ------------------------------------------------------------------------------
adf_jb_table <- data.frame(
  Series = colnames(returns),
  ADF_stat = NA_real_,
  ADF_p = NA_real_,
  JB_stat = NA_real_,
  JB_p = NA_real_,
  stringsAsFactors = FALSE
)

for (j in seq_len(K)) {
  xj <- returns[, j]
  
  adf_res <- tryCatch(
    tseries::adf.test(xj, alternative = "stationary"),
    error = function(e) NULL
  )
  
  jb_res <- tryCatch(
    tseries::jarque.bera.test(xj),
    error = function(e) NULL
  )
  
  adf_jb_table$ADF_stat[j] <- if (!is.null(adf_res)) unname(adf_res$statistic) else NA_real_
  adf_jb_table$ADF_p[j]    <- if (!is.null(adf_res)) unname(adf_res$p.value) else NA_real_
  adf_jb_table$JB_stat[j]  <- if (!is.null(jb_res)) unname(jb_res$statistic) else NA_real_
  adf_jb_table$JB_p[j]     <- if (!is.null(jb_res)) unname(jb_res$p.value) else NA_real_
}

adf_jb_table <- adf_jb_table |>
  mutate(
    ADF_reject_5pct = ADF_p < 0.05,
    JB_reject_5pct = JB_p < 0.05
  )

print(adf_jb_table)
save_table(adf_jb_table, "Table_ADF_JB_raw_returns.csv")

# ------------------------------------------------------------------------------
# 6) Serial correlation and volatility clustering diagnostics on raw returns
# ------------------------------------------------------------------------------
lb_returns <- ljung_box_table(returns, lags = LB_LAGS, squared = FALSE)
lb_squared_returns <- ljung_box_table(returns, lags = LB_LAGS, squared = TRUE)

lb_compare <- lb_returns |>
  dplyr::select(Series, Q_stat, p_value, Reject_5pct) |>
  dplyr::rename(Q_ret = Q_stat, p_ret = p_value, Reject_ret_5pct = Reject_5pct) |>
  dplyr::left_join(
    lb_squared_returns |>
      dplyr::select(Series, Q_stat, p_value, Reject_5pct) |>
      dplyr::rename(Q_sq = Q_stat, p_sq = p_value, Reject_sq_5pct = Reject_5pct),
    by = "Series"
  )

print(lb_compare)
save_table(lb_compare, "Table_LjungBox_raw_and_squared_returns.csv")

acf_raw <- make_acf_df(returns, lag_max = ACF_LAG_MAX, series_order = series_order)
p_acf_raw <- make_acf_plot(acf_raw)

print(p_acf_raw)
ggsave(file.path(output_dir, "Fig_ACF_RawReturns_3x3.png"), p_acf_raw, width = 10, height = 7.5, dpi = 300)
ggsave(file.path(output_dir, "Fig_ACF_RawReturns_3x3.pdf"), p_acf_raw, width = 10, height = 7.5)

# ------------------------------------------------------------------------------
# 7) VAR lag selection
# ------------------------------------------------------------------------------
var_selection <- VARselect(returns, lag.max = VAR_LAG_MAX, type = "const")

var_criteria <- as.data.frame(var_selection$criteria)
var_criteria$Criterion <- rownames(var_criteria)
var_criteria <- var_criteria |>
  dplyr::relocate(Criterion)

print(var_selection$criteria)
print(var_criteria)
save_table(var_criteria, "Table_VAR_lag_selection.csv")

selected_var_lags <- c(
  AIC = pick_min_lag(var_selection$criteria["AIC(n)", ]),
  HQ  = pick_min_lag(var_selection$criteria["HQ(n)", ]),
  SC  = pick_min_lag(var_selection$criteria["SC(n)", ]),
  FPE = pick_min_lag(var_selection$criteria["FPE(n)", ])
)

cat("\nVAR lag selection:\n")
print(selected_var_lags)

# ------------------------------------------------------------------------------
# 8) VAR(1) residual diagnostics
# ------------------------------------------------------------------------------
var_fit <- VAR(returns, p = VAR_P, type = "const")
var_residuals <- resid(var_fit)
colnames(var_residuals) <- colnames(returns)

acf_var_residuals <- make_acf_df(var_residuals, lag_max = ACF_LAG_MAX, series_order = series_order)
p_acf_var_residuals <- make_acf_plot(acf_var_residuals)

print(p_acf_var_residuals)
ggsave(file.path(output_dir, "Fig_ACF_VAR1_Residuals_3x3.png"), p_acf_var_residuals, width = 10, height = 7.5, dpi = 300)
ggsave(file.path(output_dir, "Fig_ACF_VAR1_Residuals_3x3.pdf"), p_acf_var_residuals, width = 10, height = 7.5)

lb_var_residuals <- ljung_box_table(var_residuals, lags = LB_LAGS, squared = FALSE) |>
  dplyr::mutate(Series = factor(Series, levels = series_order)) |>
  dplyr::arrange(Series)

lb_squared_var_residuals <- ljung_box_table(var_residuals, lags = LB_LAGS, squared = TRUE) |>
  dplyr::mutate(Series = factor(Series, levels = series_order)) |>
  dplyr::arrange(Series)

print(lb_var_residuals)
print(lb_squared_var_residuals)
save_table(lb_var_residuals, "Table_LjungBox_VAR1_residuals.csv")
save_table(lb_squared_var_residuals, "Table_LjungBox_squared_VAR1_residuals.csv")

# ------------------------------------------------------------------------------
# 9) Innovation extraction for multivariate GARCH models
# ------------------------------------------------------------------------------
# Mean specification used here:
#   - API2 coal returns: AR(1)
#   - All other series: constant mean

coal_name <- "API2 (Coal)"
coal_idx <- which(colnames(returns) == coal_name)

if (length(coal_idx) != 1) {
  stop("Could not uniquely identify the API2 coal column.")
}

innovations <- matrix(NA_real_, nrow = nrow(returns), ncol = ncol(returns))
colnames(innovations) <- colnames(returns)

coal_fit <- arima(returns[, coal_idx], order = c(1, 0, 0), include.mean = TRUE)
innovations[, coal_idx] <- residuals(coal_fit)

for (j in seq_len(ncol(returns))) {
  if (j != coal_idx) {
    innovations[, j] <- returns[, j] - mean(returns[, j], na.rm = TRUE)
  }
}

# Align all series to the AR(1) innovation sample.
innovations <- innovations[-1, , drop = FALSE]
innovation_dates <- dates[-1]

arch_table <- arch_lm_table(innovations, lags = ARCH_LAGS) |>
  dplyr::mutate(Series = factor(Series, levels = series_order)) |>
  dplyr::arrange(Series)

print(arch_table)
save_table(arch_table, "Table_ARCH_LM_innovations.csv")

innovation_df <- data.frame(Date = innovation_dates, innovations, check.names = FALSE)

saveRDS(innovation_df, file.path(output_dir, "epsilon_innovations.rds"))
write.csv(innovation_df, file.path(output_dir, "epsilon_innovations.csv"), row.names = FALSE)

# ------------------------------------------------------------------------------
# 10) Summary
# ------------------------------------------------------------------------------
n_lb_reject <- sum(lb_returns$Reject_5pct, na.rm = TRUE)
n_arch_reject <- sum(arch_table$Reject_5pct, na.rm = TRUE)

cat("\n=============================\n")
cat("PRE-MODELING SUMMARY\n")
cat("=============================\n")
cat("Observations in raw returns:", Tn, "\n")
cat("Number of series:", K, "\n")
cat("Ljung-Box lag:", LB_LAGS, "\n")
cat("Raw-return Ljung-Box rejections at 5%:", n_lb_reject, "of", K, "\n")
cat("VARselect suggested lags:\n")
print(selected_var_lags)
cat("ARCH-LM lag:", ARCH_LAGS, "\n")
cat("ARCH-LM rejections at 5%:", n_arch_reject, "of", K, "\n")
cat("\nSaved outputs to:\n")
cat(output_dir, "\n")
cat("\nSaved innovation files:\n")
cat(" - epsilon_innovations.rds\n")
cat(" - epsilon_innovations.csv\n")
cat("Rows:", nrow(innovation_df), "Cols:", ncol(innovation_df) - 1, "\n")
