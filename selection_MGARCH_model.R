################################################################################
# Volatility estimation, diagnostics, connectedness inputs, and model comparison
# Models: sBEKK, dBEKK, asymmetric BEKK, DCC, two-step DCC, and cDCC
################################################################################

# ------------------------------------------------------------------------------
# 0) Libraries
# ------------------------------------------------------------------------------
suppressPackageStartupMessages({
  library(zoo)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(rugarch)
  library(rmgarch)
  library(BEKKs)
  library(xdcclarge)
  library(portes)
})

# ------------------------------------------------------------------------------
# 1) User settings
# ------------------------------------------------------------------------------
base_dir <- "C:/Users/sezue/OneDrive/Desktop/BA"
input_file <- file.path(base_dir, "outputs", "epsilon_innovations.csv")

if (!file.exists(input_file)) {
  input_file <- file.path(base_dir, "epsilon_innovations.csv")
}

output_dir <- file.path(base_dir, "outputs")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

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

commodities <- c(
  "TTF (Gas)",
  "Brent (Oil)",
  "API2 (Coal)",
  "MO1 (Carbon)"
)

sector_order <- c(
  "SX6P (Utilities)",
  "SXEP (Oil & Gas)",
  "SXQP (Consumer)",
  "SX7P (Financials)",
  "SXNP (Industrials)"
)

models_raw <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "sBEKK_asym",
  "dBEKK_asym",
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage"
)

models_for_connectedness <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "sBEKK_asym",
  "dBEKK_asym",
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage"
)

bekk_mean_models <- c("sBEKK_sym", "dBEKK_sym", "sBEKK_asym", "dBEKK_asym")

BEKK_MAX_ITER <- 200
BEKK_CRIT <- 1e-6
CDCC_METHOD <- "NLS"

PT_LAGS <- 20
LB_LAGS <- 20
EPS_PD <- 1e-12

mode_used <- "dlog"
q_high_top25 <- 0.75
alpha_grid_used <- 0:8
alpha_min_used <- 0
alpha_max_used <- 8
nlag_used <- 1
nfore_used <- 20
window_qw_used <- 60

# ------------------------------------------------------------------------------
# 2) General helpers
# ------------------------------------------------------------------------------
strip_vol_suffix <- function(x) {
  sub("_vol$", "", x)
}

ensure_vol_colnames <- function(vol_zoo, model_name = "model") {
  stopifnot(inherits(vol_zoo, "zoo"))
  cn <- colnames(vol_zoo)
  if (is.null(cn)) stop(model_name, ": volatility object has no column names.")
  colnames(vol_zoo) <- ifelse(grepl("_vol$", cn), cn, paste0(cn, "_vol"))
  vol_zoo
}

make_short_key <- function(long_name) {
  gsub("\\s.*$", "", long_name)
}

make_name_map <- function(long_names) {
  short_keys <- vapply(long_names, make_short_key, character(1))
  stats::setNames(long_names, short_keys)
}

rename_short_to_long <- function(x, name_map, fallback_names = NULL) {
  x <- as.matrix(x)
  
  if (is.null(colnames(x))) {
    if (is.null(fallback_names)) stop("Object has no column names and no fallback names were supplied.")
    colnames(x) <- fallback_names[seq_len(ncol(x))]
    return(x)
  }
  
  old <- colnames(x)
  new <- old
  hit <- old %in% names(name_map)
  new[hit] <- unname(name_map[old[hit]])
  colnames(x) <- new
  x
}

apply_canonical_names <- function(x, canonical_names) {
  x <- as.matrix(x)
  k <- ncol(x)
  
  if (k <= length(canonical_names)) {
    colnames(x) <- canonical_names[seq_len(k)]
  } else {
    colnames(x) <- c(canonical_names, paste0("Extra", seq_len(k - length(canonical_names))))
  }
  
  x
}

safe_get <- function(expr, msg) {
  tryCatch(expr, error = function(e) {
    warning(msg, " :: ", conditionMessage(e))
    NULL
  })
}

save_table <- function(x, filename) {
  write.csv(x, file.path(output_dir, filename), row.names = FALSE)
}

save_vol_csv <- function(vol_zoo, filename) {
  df_out <- data.frame(
    Date = as.Date(zoo::index(vol_zoo)),
    zoo::coredata(vol_zoo),
    check.names = FALSE
  )
  
  write.csv(
    df_out,
    file.path(output_dir, filename),
    row.names = FALSE
  )
}

# ------------------------------------------------------------------------------
# 3) Read innovations
# ------------------------------------------------------------------------------
eps_df <- read.csv(input_file, stringsAsFactors = FALSE, check.names = FALSE)
eps_df$Date <- as.Date(eps_df$Date)

dates <- eps_df$Date
E <- as.matrix(eps_df[, -1, drop = FALSE])
storage.mode(E) <- "numeric"

stopifnot(!anyNA(E))

if (is.null(colnames(E))) {
  colnames(E) <- series_order[seq_len(ncol(E))]
}

if (ncol(E) == length(series_order)) {
  colnames(E) <- series_order
}

T_obs <- nrow(E)
K <- ncol(E)

cat("Innovations loaded successfully.\n")
cat("Observations:", T_obs, "\n")
cat("Series:", K, "\n\n")

# ------------------------------------------------------------------------------
# 4) BEKK volatility extraction
# ------------------------------------------------------------------------------
extract_bekk_vol <- function(fit, dates, base_names) {
  Hvec <- fit$H_t
  if (is.null(Hvec)) stop("fit$H_t is NULL.")
  
  Tn <- nrow(Hvec)
  K_fit <- length(base_names)
  diag_idx <- seq(1, K_fit * K_fit, by = K_fit + 1)
  vol_mat <- sqrt(pmax(Hvec[, diag_idx, drop = FALSE], 0))
  
  colnames(vol_mat) <- paste0(base_names, "_vol")
  zoo::zoo(vol_mat, order.by = dates[seq_len(Tn)])
}

fit_bekk_model <- function(E, dates, type, asymmetric, max_iter = BEKK_MAX_ITER, crit = BEKK_CRIT) {
  spec <- BEKKs::bekk_spec(model = list(type = type, asymmetric = asymmetric))
  fit <- BEKKs::bekk_fit(
    spec,
    data = E,
    QML_t_ratios = FALSE,
    max_iter = max_iter,
    crit = crit
  )
  
  vol <- extract_bekk_vol(fit, dates, colnames(E))
  list(fit = fit, vol = vol)
}

# ------------------------------------------------------------------------------
# 5) DCC and cDCC volatility extraction
# ------------------------------------------------------------------------------
make_univariate_spec <- function() {
  rugarch::ugarchspec(
    variance.model = list(model = "sGARCH", garchOrder = c(1, 1)),
    mean.model = list(armaOrder = c(0, 0), include.mean = FALSE),
    distribution.model = "norm"
  )
}

fit_multifit_stage <- function(E) {
  K_fit <- ncol(E)
  uspec <- make_univariate_spec()
  mspec <- rugarch::multispec(replicate(K_fit, uspec))
  rugarch::multifit(mspec, data = E)
}

extract_dcc_vol_rmgarch <- function(E, dates) {
  K_fit <- ncol(E)
  uspec <- make_univariate_spec()
  mspec <- rugarch::multispec(replicate(K_fit, uspec))
  
  dcc_spec <- rmgarch::dccspec(
    uspec = mspec,
    dccOrder = c(1, 1),
    model = "DCC",
    distribution = "mvnorm"
  )
  
  dcc_fit <- rmgarch::dccfit(dcc_spec, data = E, fit.control = list(eval.se = FALSE))
  Ht <- rmgarch::rcov(dcc_fit)
  Tn <- dim(Ht)[3]
  
  vol <- matrix(NA_real_, Tn, K_fit)
  for (tt in seq_len(Tn)) {
    vol[tt, ] <- sqrt(pmax(diag(Ht[, , tt]), 0))
  }
  
  colnames(vol) <- paste0(colnames(E), "_vol")
  
  list(
    fit = dcc_fit,
    vol = zoo::zoo(vol, order.by = dates[seq_len(Tn)])
  )
}

extract_cdcc_vol_xdcclarge <- function(E, dates, method = CDCC_METHOD) {
  if (!exists("cdcc_estimation", mode = "function")) {
    stop("cdcc_estimation() was not found. Load or install xdcclarge before estimating cDCC.")
  }
  
  fit_g <- fit_multifit_stage(E)
  ht <- as.matrix(rugarch::sigma(fit_g)^2)
  resid <- as.matrix(rugarch::residuals(fit_g))
  Tn <- nrow(resid)
  
  cdcc_fit <- cdcc_estimation(
    ini.para = c(0.05, 0.93),
    ht = ht,
    residuals = resid,
    method = method,
    ts = Tn
  )
  
  vol <- sqrt(pmax(ht, 0))
  colnames(vol) <- paste0(colnames(E), "_vol")
  
  list(
    fit = cdcc_fit,
    fit_g = fit_g,
    vol = zoo::zoo(vol, order.by = dates[seq_len(Tn)])
  )
}

# ------------------------------------------------------------------------------
# 6) Two-step scalar DCC helpers
# ------------------------------------------------------------------------------
dcc_scalar_build <- function(z, a, b, eps = 1e-12) {
  z <- as.matrix(z)
  Tn <- nrow(z)
  K_fit <- ncol(z)
  
  Qbar <- stats::cor(z, use = "pairwise.complete.obs")
  Qbar <- (Qbar + t(Qbar)) / 2
  
  Q <- Qbar
  Rt <- array(NA_real_, dim = c(K_fit, K_fit, Tn))
  Rt[, , 1] <- stats::cov2cor(Qbar)
  
  for (tt in 2:Tn) {
    zlag <- as.numeric(z[tt - 1, ])
    Q <- (1 - a - b) * Qbar + a * tcrossprod(zlag) + b * Q
    Q <- (Q + t(Q)) / 2
    
    d <- sqrt(pmax(diag(Q), eps))
    Dinv <- diag(1 / d, K_fit, K_fit)
    R <- Dinv %*% Q %*% Dinv
    R <- (R + t(R)) / 2
    diag(R) <- 1
    Rt[, , tt] <- R
  }
  
  list(R = Rt, Qbar = Qbar)
}

dcc_scalar_negloglik <- function(par, z, penalty = 1e12, eps = 1e-12) {
  a <- max(par[1], 0)
  b <- max(par[2], 0)
  
  if (a + b >= 0.999) return(penalty)
  
  z <- as.matrix(z)
  Tn <- nrow(z)
  K_fit <- ncol(z)
  
  Qbar <- stats::cor(z, use = "pairwise.complete.obs")
  Qbar <- (Qbar + t(Qbar)) / 2
  Q <- Qbar
  nll <- 0
  
  for (tt in 2:Tn) {
    zlag <- as.numeric(z[tt - 1, ])
    Q <- (1 - a - b) * Qbar + a * tcrossprod(zlag) + b * Q
    Q <- (Q + t(Q)) / 2
    
    d <- sqrt(pmax(diag(Q), eps))
    Dinv <- diag(1 / d, K_fit, K_fit)
    R <- Dinv %*% Q %*% Dinv
    R <- (R + t(R)) / 2
    diag(R) <- 1
    
    ev <- eigen(R, symmetric = TRUE, only.values = TRUE)$values
    if (any(!is.finite(ev)) || any(ev <= eps)) return(penalty)
    
    logdet <- as.numeric(determinant(R, logarithm = TRUE)$modulus)
    quad <- drop(crossprod(z[tt, ], solve(R, z[tt, ])))
    nll <- nll + 0.5 * (logdet + quad)
  }
  
  nll / (Tn - 1)
}

standardize_R_array <- function(R_array, K_fit) {
  if (is.list(R_array)) R_array <- simplify2array(R_array)
  
  if (length(dim(R_array)) != 3) {
    stop("R_array is not 3D. dim = ", paste(dim(R_array), collapse = "x"))
  }
  
  dR <- dim(R_array)
  
  if (identical(dR[1:2], c(K_fit, K_fit))) return(R_array)
  if (identical(dR[2:3], c(K_fit, K_fit))) return(aperm(R_array, c(2, 3, 1)))
  if (identical(c(dR[1], dR[3]), c(K_fit, K_fit))) return(aperm(R_array, c(1, 3, 2)))
  
  stop(
    "Cannot reconcile R_array dimensions with K=", K_fit,
    ". dim(R_array)=", paste(dR, collapse = "x")
  )
}

Ht_from_Rt_and_ht <- function(Rt, ht, eps = 1e-12) {
  ht <- as.matrix(ht)
  Tuse <- dim(Rt)[3]
  N <- dim(Rt)[1]
  Ht <- array(NA_real_, dim = c(N, N, Tuse))
  
  for (tt in seq_len(Tuse)) {
    D <- diag(sqrt(pmax(ht[tt, ], eps)), N, N)
    H <- D %*% Rt[, , tt] %*% D
    Ht[, , tt] <- (H + t(H)) / 2
  }
  
  Ht
}

fit_dcc_scalar_stage <- function(fit_g, dates, base_names) {
  sigma_t <- as.matrix(rugarch::sigma(fit_g))
  resid_t <- as.matrix(rugarch::residuals(fit_g))
  z_t <- resid_t / sigma_t
  
  opt <- optim(
    par = c(0.05, 0.93),
    fn = dcc_scalar_negloglik,
    z = z_t,
    method = "L-BFGS-B",
    lower = c(0, 0),
    upper = c(0.999, 0.999)
  )
  
  a_hat <- opt$par[1]
  b_hat <- opt$par[2]
  
  dcc_stage_obj <- dcc_scalar_build(z_t, a_hat, b_hat)
  R_array <- standardize_R_array(dcc_stage_obj$R, ncol(z_t))
  
  Tuse <- min(dim(R_array)[3], nrow(sigma_t))
  ht <- sigma_t[seq_len(Tuse), , drop = FALSE]^2
  Ht_stage <- Ht_from_Rt_and_ht(R_array[, , seq_len(Tuse), drop = FALSE], ht, eps = EPS_PD)
  
  vol <- matrix(NA_real_, Tuse, ncol(z_t))
  for (tt in seq_len(Tuse)) {
    vol[tt, ] <- sqrt(pmax(diag(Ht_stage[, , tt]), 0))
  }
  
  colnames(vol) <- paste0(base_names, "_vol")
  
  list(
    fit = dcc_stage_obj,
    Ht = Ht_stage,
    vol = zoo::zoo(vol, order.by = dates[seq_len(Tuse)]),
    ab = c(a = a_hat, b = b_hat),
    z = z_t,
    ht = ht
  )
}

# ------------------------------------------------------------------------------
# 7) cDCC correlation unpacking
# ------------------------------------------------------------------------------
infer_N_from_p <- function(p) {
  r <- sqrt(p)
  if (is.finite(r) && abs(r - round(r)) < 1e-8) return(as.integer(round(r)))
  
  disc <- 1 + 8 * p
  r2 <- (-1 + sqrt(disc)) / 2
  if (is.finite(r2) && abs(r2 - round(r2)) < 1e-8) return(as.integer(round(r2)))
  
  NA_integer_
}

unpack_Rt_row <- function(v, N) {
  v <- as.numeric(v)
  
  if (length(v) == N * N) {
    R <- matrix(v, N, N, byrow = FALSE)
    R <- (R + t(R)) / 2
    diag(R) <- 1
    return(R)
  }
  
  if (length(v) == N * (N + 1) / 2) {
    R <- matrix(0, N, N)
    idx <- 1
    for (j in seq_len(N)) {
      for (i in seq_len(j)) {
        R[i, j] <- v[idx]
        R[j, i] <- v[idx]
        idx <- idx + 1
      }
    }
    R <- (R + t(R)) / 2
    diag(R) <- 1
    return(R)
  }
  
  stop("Cannot unpack Rt row of length ", length(v), " for N = ", N, ".")
}

cdcc_Rt_array_from_fit <- function(cdcc_fit, Tuse_target) {
  if (is.null(cdcc_fit$cdcc_Rt)) stop("cdcc_fit$cdcc_Rt is NULL.")
  
  M <- suppressWarnings(as.matrix(cdcc_fit$cdcc_Rt))
  
  N_by_cols <- infer_N_from_p(ncol(M))
  N_by_rows <- infer_N_from_p(nrow(M))
  
  if (is.na(N_by_cols) && !is.na(N_by_rows)) {
    M <- t(M)
    N_by_cols <- infer_N_from_p(ncol(M))
  }
  
  if (is.na(N_by_cols)) {
    stop("Unsupported cdcc_Rt shape: ", nrow(M), " x ", ncol(M), ".")
  }
  
  N_use <- N_by_cols
  Tuse <- min(Tuse_target, nrow(M))
  Rt <- array(NA_real_, dim = c(N_use, N_use, Tuse))
  
  for (tt in seq_len(Tuse)) {
    Rt[, , tt] <- unpack_Rt_row(M[tt, ], N_use)
  }
  
  Rt
}

# ------------------------------------------------------------------------------
# 8) Estimate volatility models
# ------------------------------------------------------------------------------
cat("\n--- Estimating sBEKK symmetric ---\n")
sbekk_sym_out <- fit_bekk_model(E, dates, type = "sbekk", asymmetric = FALSE)
fit_s_sym <- sbekk_sym_out$fit
vol_sbekk_sym <- sbekk_sym_out$vol

cat("\n--- Estimating dBEKK symmetric ---\n")
dbekk_sym_out <- fit_bekk_model(E, dates, type = "dbekk", asymmetric = FALSE)
fit_d_sym <- dbekk_sym_out$fit
vol_dbekk_sym <- dbekk_sym_out$vol

cat("\n--- Estimating sBEKK asymmetric ---\n")
sbekk_asym_out <- fit_bekk_model(E, dates, type = "sbekk", asymmetric = TRUE)
fit_s_asym <- sbekk_asym_out$fit
vol_sbekk_asym <- sbekk_asym_out$vol

cat("\n--- Estimating dBEKK asymmetric ---\n")
dbekk_asym_out <- fit_bekk_model(E, dates, type = "dbekk", asymmetric = TRUE)
fit_d_asym <- dbekk_asym_out$fit
vol_dbekk_asym <- dbekk_asym_out$vol

cat("\n--- Estimating DCC full ---\n")
dcc_out <- extract_dcc_vol_rmgarch(E, dates)
dcc_fit <- dcc_out$fit
vol_dcc_full <- dcc_out$vol

cat("\n--- Estimating shared univariate GARCH stage ---\n")
fit_g <- fit_multifit_stage(E)

cat("\n--- Estimating two-step scalar DCC stage ---\n")
dcc_scalar_out <- fit_dcc_scalar_stage(fit_g, dates, colnames(E))
dcc_stage_obj <- dcc_scalar_out$fit
dcc_stage_ab <- dcc_scalar_out$ab
Ht_stage <- dcc_scalar_out$Ht
z_g <- dcc_scalar_out$z
ht_g <- dcc_scalar_out$ht
vol_dcc_scalar_stage <- dcc_scalar_out$vol

cat("\n--- Estimating cDCC Aielli stage ---\n")
cdcc_out <- extract_cdcc_vol_xdcclarge(E, dates, method = CDCC_METHOD)
cdcc_fit <- cdcc_out$fit
vol_cdcc_aielli_stage <- cdcc_out$vol

vol_models <- list(
  sBEKK_sym = vol_sbekk_sym,
  dBEKK_sym = vol_dbekk_sym,
  sBEKK_asym = vol_sbekk_asym,
  dBEKK_asym = vol_dbekk_asym,
  DCC_full = vol_dcc_full,
  DCC_scalar_stage = vol_dcc_scalar_stage,
  cDCC_Aielli_stage = vol_cdcc_aielli_stage
)

vol_models <- stats::setNames(
  lapply(names(vol_models), function(nm) ensure_vol_colnames(vol_models[[nm]], nm)),
  names(vol_models)
)

saveRDS(vol_models, file.path(output_dir, "vol_models_all_7models.rds"))

save_vol_csv(vol_sbekk_sym, "vol_sBEKK_sym.csv")
save_vol_csv(vol_dbekk_sym, "vol_dBEKK_sym.csv")
save_vol_csv(vol_sbekk_asym, "vol_sBEKK_asym.csv")
save_vol_csv(vol_dbekk_asym, "vol_dBEKK_asym.csv")
save_vol_csv(vol_dcc_full, "vol_DCC_full.csv")
save_vol_csv(vol_dcc_scalar_stage, "vol_DCC_scalar_stage.csv")
save_vol_csv(vol_cdcc_aielli_stage, "vol_cDCC_Aielli_stage.csv")

cat("\nVolatility models saved.\n")
cat("Two-step DCC parameters: a =", round(dcc_stage_ab["a"], 4), ", b =", round(dcc_stage_ab["b"], 4), "\n")

# ------------------------------------------------------------------------------
# 9) Standardized residual diagnostics
# ------------------------------------------------------------------------------
get_Z_bekk <- function(fit, canonical_names) {
  e <- as.matrix(fit$e_t)
  Hvec <- as.matrix(fit$H_t)
  
  if (is.null(e) || is.null(Hvec)) stop("BEKK fit is missing e_t or H_t.")
  
  N <- ncol(e)
  Tn <- nrow(e)
  
  if (ncol(Hvec) != N * N) stop("BEKK H_t does not have N^2 columns.")
  
  cn <- canonical_names
  if (length(cn) < N) cn <- c(cn, paste0("Extra", seq_len(N - length(cn))))
  cn <- cn[seq_len(N)]
  
  Z <- matrix(NA_real_, Tn, N)
  colnames(Z) <- cn
  
  for (tt in seq_len(Tn)) {
    H <- matrix(Hvec[tt, ], nrow = N, byrow = TRUE)
    H <- (H + t(H)) / 2
    R <- chol(H + diag(1e-10, N))
    Z[tt, ] <- backsolve(R, e[tt, ], transpose = FALSE)
  }
  
  Z
}

get_Z_dcc_full <- function(dcc_fit, canonical_names) {
  Z <- as.matrix(rmgarch::residuals(dcc_fit)) / as.matrix(rmgarch::sigma(dcc_fit))
  apply_canonical_names(Z, canonical_names)
}

run_hosking <- function(Z, model_name, lags = PT_LAGS) {
  Z <- as.matrix(Z)
  Z <- Z[complete.cases(Z), , drop = FALSE]
  
  if (nrow(Z) <= lags + 5) stop("Not enough observations for Hosking test.")
  
  h1 <- portes::Hosking(Z, lags = lags, fitdf = 0, sqrd.res = FALSE)
  h2 <- portes::Hosking(Z, lags = lags, fitdf = 0, sqrd.res = TRUE)
  
  get_stat_p <- function(h) {
    if (is.list(h) && !is.null(h$statistic) && !is.null(h$p.value)) {
      return(c(stat = as.numeric(h$statistic), p = as.numeric(h$p.value)))
    }
    
    if (is.matrix(h) || is.data.frame(h)) {
      df <- as.data.frame(h)
      nms <- names(df)
      stat_col <- nms[grepl("stat|Q", nms, ignore.case = TRUE)][1]
      p_col <- nms[grepl("p", nms, ignore.case = TRUE)][1]
      
      if (is.na(stat_col) || is.na(p_col)) {
        stat_col <- nms[max(1, ncol(df) - 1)]
        p_col <- nms[ncol(df)]
      }
      
      return(c(stat = as.numeric(df[1, stat_col]), p = as.numeric(df[1, p_col])))
    }
    
    v <- as.numeric(h)
    if (length(v) >= 2) return(c(stat = v[length(v) - 1], p = v[length(v)]))
    
    stop("Cannot parse Hosking output.")
  }
  
  sp1 <- get_stat_p(h1)
  sp2 <- get_stat_p(h2)
  
  data.frame(
    Model = model_name,
    Type = c("StdResiduals", "StdResiduals^2"),
    lags = lags,
    Statistic = c(sp1["stat"], sp2["stat"]),
    p_value = c(sp1["p"], sp2["p"]),
    stringsAsFactors = FALSE
  )
}

ljungbox_sq_table <- function(Z, model_name, lags = LB_LAGS) {
  Z <- as.matrix(Z)
  Z <- Z[complete.cases(Z), , drop = FALSE]
  
  if (is.null(colnames(Z))) colnames(Z) <- paste0("S", seq_len(ncol(Z)))
  
  out <- data.frame(
    Model = model_name,
    Series = colnames(Z),
    Q_stat = NA_real_,
    p_value = NA_real_,
    Reject_5pct = NA,
    stringsAsFactors = FALSE
  )
  
  for (j in seq_len(ncol(Z))) {
    bt <- Box.test(Z[, j]^2, lag = lags, type = "Ljung-Box")
    out$Q_stat[j] <- unname(bt$statistic)
    out$p_value[j] <- unname(bt$p.value)
    out$Reject_5pct[j] <- out$p_value[j] < 0.05
  }
  
  out
}

Z_models <- list(
  sBEKK_sym = safe_get(get_Z_bekk(fit_s_sym, series_order), "sBEKK_sym standardized residuals failed"),
  dBEKK_sym = safe_get(get_Z_bekk(fit_d_sym, series_order), "dBEKK_sym standardized residuals failed"),
  sBEKK_asym = safe_get(get_Z_bekk(fit_s_asym, series_order), "sBEKK_asym standardized residuals failed"),
  dBEKK_asym = safe_get(get_Z_bekk(fit_d_asym, series_order), "dBEKK_asym standardized residuals failed"),
  DCC_full = safe_get(get_Z_dcc_full(dcc_fit, series_order), "DCC_full standardized residuals failed"),
  DCC_scalar_stage = apply_canonical_names(z_g, series_order),
  cDCC_Aielli_stage = apply_canonical_names(z_g, series_order)
)

Z_models <- Z_models[!vapply(Z_models, is.null, logical(1))]

hosking_table <- dplyr::bind_rows(lapply(names(Z_models), function(nm) {
  safe_get(run_hosking(Z_models[[nm]], nm, PT_LAGS), paste(nm, "Hosking test failed"))
}))

hosking_table <- hosking_table |>
  dplyr::mutate(Reject_5pct = p_value < 0.05)

ljungbox_sq_results <- dplyr::bind_rows(lapply(names(Z_models), function(nm) {
  ljungbox_sq_table(Z_models[[nm]], nm, LB_LAGS)
}))

print(hosking_table)
print(ljungbox_sq_results)

save_table(hosking_table, "Table_Hosking_standardized_residuals.csv")
save_table(ljungbox_sq_results, "Table_LjungBox_squared_standardized_residuals.csv")

# ------------------------------------------------------------------------------
# 10) H_t validity checks
# ------------------------------------------------------------------------------
bekk_H_to_array <- function(fit) {
  Hvec <- as.matrix(fit$H_t)
  if (is.null(Hvec)) stop("BEKK fit has no H_t.")
  if (is.null(fit$e_t)) stop("Cannot infer dimension because BEKK fit has no e_t.")
  
  N <- ncol(as.matrix(fit$e_t))
  if (ncol(Hvec) != N * N) stop("BEKK H_t does not have N^2 columns.")
  
  Tn <- nrow(Hvec)
  H_array <- array(NA_real_, dim = c(N, N, Tn))
  
  for (tt in seq_len(Tn)) {
    H_array[, , tt] <- matrix(Hvec[tt, ], nrow = N, byrow = TRUE)
  }
  
  H_array
}

ht_validity_checks <- function(H_array, model_name, tol_pd = 1e-10) {
  N <- dim(H_array)[1]
  Tn <- dim(H_array)[3]
  
  finite_ok <- logical(Tn)
  sym_err_max <- numeric(Tn)
  chol_ok <- logical(Tn)
  min_eig <- numeric(Tn)
  cond_num <- numeric(Tn)
  var_min <- numeric(Tn)
  corr_max_abs <- numeric(Tn)
  
  for (tt in seq_len(Tn)) {
    Ht <- H_array[, , tt]
    finite_ok[tt] <- all(is.finite(Ht))
    
    if (!finite_ok[tt]) {
      sym_err_max[tt] <- NA_real_
      chol_ok[tt] <- FALSE
      min_eig[tt] <- NA_real_
      cond_num[tt] <- NA_real_
      var_min[tt] <- NA_real_
      corr_max_abs[tt] <- NA_real_
      next
    }
    
    sym_err_max[tt] <- max(abs(Ht - t(Ht)))
    Hs <- (Ht + t(Ht)) / 2
    ev <- eigen(Hs, symmetric = TRUE, only.values = TRUE)$values
    
    min_eig[tt] <- min(ev)
    chol_ok[tt] <- !inherits(try(chol(Hs + diag(tol_pd, N)), silent = TRUE), "try-error")
    cond_num[tt] <- if (min(ev) > tol_pd) max(ev) / min(ev) else Inf
    var_min[tt] <- min(diag(Hs))
    
    Dinv <- diag(1 / sqrt(pmax(diag(Hs), tol_pd)))
    Rt <- Dinv %*% Hs %*% Dinv
    Rt[diag(N) == 1] <- 0
    corr_max_abs[tt] <- max(abs(Rt), na.rm = TRUE)
  }
  
  data.frame(
    Model = model_name,
    T = Tn,
    K = N,
    Finite_share = mean(finite_ok),
    SymErr_max = max(sym_err_max, na.rm = TRUE),
    SymErr_q99 = unname(quantile(sym_err_max, 0.99, na.rm = TRUE)),
    PD_share_chol = mean(chol_ok),
    MinEig_min = min(min_eig, na.rm = TRUE),
    MinEig_q01 = unname(quantile(min_eig, 0.01, na.rm = TRUE)),
    CondNum_med = unname(median(cond_num[is.finite(cond_num)], na.rm = TRUE)),
    CondNum_q99 = unname(quantile(cond_num[is.finite(cond_num)], 0.99, na.rm = TRUE)),
    VarMin_min = min(var_min, na.rm = TRUE),
    CorrMaxAbs_max = max(corr_max_abs, na.rm = TRUE),
    stringsAsFactors = FALSE
  )
}

Ht_cdcc <- NULL
Rt_cdcc <- safe_get(cdcc_Rt_array_from_fit(cdcc_fit, Tuse_target = nrow(ht_g)), "cDCC Rt unpack failed")

if (!is.null(Rt_cdcc)) {
  N_cdcc <- dim(Rt_cdcc)[1]
  if (ncol(ht_g) >= N_cdcc) {
    Tuse <- dim(Rt_cdcc)[3]
    Ht_cdcc <- Ht_from_Rt_and_ht(
      Rt_cdcc,
      ht_g[seq_len(Tuse), seq_len(N_cdcc), drop = FALSE],
      eps = EPS_PD
    )
  }
}

H_list <- list(
  sBEKK_sym = safe_get(bekk_H_to_array(fit_s_sym), "sBEKK_sym H array failed"),
  dBEKK_sym = safe_get(bekk_H_to_array(fit_d_sym), "dBEKK_sym H array failed"),
  sBEKK_asym = safe_get(bekk_H_to_array(fit_s_asym), "sBEKK_asym H array failed"),
  dBEKK_asym = safe_get(bekk_H_to_array(fit_d_asym), "dBEKK_asym H array failed"),
  DCC_full = safe_get(rmgarch::rcov(dcc_fit), "DCC_full H array failed"),
  DCC_scalar_stage = Ht_stage,
  cDCC_Aielli_stage = Ht_cdcc
)

H_list <- H_list[!vapply(H_list, is.null, logical(1))]

ht_summary_table <- dplyr::bind_rows(lapply(names(H_list), function(nm) {
  ht_validity_checks(H_list[[nm]], nm)
}))

print(ht_summary_table)
save_table(ht_summary_table, "Table_Ht_validity_summary.csv")

# ------------------------------------------------------------------------------
# 11) Aligned Gaussian log-likelihood and information criteria
# ------------------------------------------------------------------------------
gauss_loglik_from_Ht <- function(y, Ht, eps_pd = 1e-12) {
  y <- as.matrix(y)
  N <- ncol(y)
  Tn <- dim(Ht)[3]
  
  if (nrow(y) < Tn) stop("y has fewer rows than H_t time dimension.")
  
  y <- y[seq_len(Tn), , drop = FALSE]
  ll <- 0
  c0 <- N * log(2 * pi)
  
  for (tt in seq_len(Tn)) {
    H <- (Ht[, , tt] + t(Ht[, , tt])) / 2
    ev <- eigen(H, symmetric = TRUE, only.values = TRUE)$values
    
    if (any(!is.finite(ev)) || any(ev <= eps_pd)) return(NA_real_)
    
    logdet <- as.numeric(determinant(H, logarithm = TRUE)$modulus)
    quad <- drop(crossprod(y[tt, ], solve(H, y[tt, ])))
    ll <- ll + (-0.5) * (c0 + logdet + quad)
  }
  
  ll
}

aic_from_ll <- function(ll, k) {
  if (anyNA(c(ll, k))) NA_real_ else -2 * ll + 2 * k
}

bic_from_ll <- function(ll, k, n) {
  if (anyNA(c(ll, k, n))) NA_real_ else -2 * ll + k * log(n)
}

bekk_Ht_array_tryboth <- function(fit, N, y_sub) {
  Hvec <- as.matrix(fit$H_t)
  if (is.null(Hvec)) stop("fit$H_t is NULL.")
  if (ncol(Hvec) != N * N) stop("fit$H_t does not have N^2 columns.")
  
  Tn <- nrow(Hvec)
  HtA <- array(NA_real_, dim = c(N, N, Tn))
  HtB <- array(NA_real_, dim = c(N, N, Tn))
  
  for (tt in seq_len(Tn)) {
    Ha <- matrix(Hvec[tt, ], nrow = N, ncol = N, byrow = FALSE)
    Hb <- matrix(Hvec[tt, ], nrow = N, ncol = N, byrow = TRUE)
    HtA[, , tt] <- (Ha + t(Ha)) / 2
    HtB[, , tt] <- (Hb + t(Hb)) / 2
  }
  
  llA <- suppressWarnings(gauss_loglik_from_Ht(y_sub, HtA))
  llB <- suppressWarnings(gauss_loglik_from_Ht(y_sub, HtB))
  
  if (is.finite(llA) && !is.finite(llB)) return(HtA)
  if (is.finite(llB) && !is.finite(llA)) return(HtB)
  if (is.finite(llA) && is.finite(llB)) return(if (llA >= llB) HtA else HtB)
  
  HtA
}

npar_bekk <- function(fit) {
  if (!is.null(fit$theta)) length(fit$theta) else NA_integer_
}

npar_multifit_total <- function(fit_g) {
  k <- NA_integer_
  
  try({
    fits <- fit_g@fit
    if (is.list(fits) && length(fits) > 0) {
      k <- sum(vapply(fits, function(onefit) length(coef(onefit)), numeric(1)))
    }
  }, silent = TRUE)
  
  if (!is.na(k)) return(as.integer(k))
  
  cf <- coef(fit_g)
  if (is.matrix(cf)) return(as.integer(nrow(cf) * ncol(cf)))
  
  NA_integer_
}

name_map <- make_name_map(series_order)
y_full <- rename_short_to_long(E, name_map, fallback_names = series_order)

ht_g_full <- as.matrix(rugarch::sigma(fit_g)^2)
ht_g_full <- rename_short_to_long(ht_g_full, name_map, fallback_names = series_order)

series_keep <- series_order[series_order %in% colnames(y_full)]
y <- y_full[, series_keep, drop = FALSE]
ht_sub <- ht_g_full[, series_keep, drop = FALSE]
N <- ncol(y)

Ht_list_models <- list(
  sBEKK_sym = bekk_Ht_array_tryboth(fit_s_sym, N, y),
  dBEKK_sym = bekk_Ht_array_tryboth(fit_d_sym, N, y),
  sBEKK_asym = bekk_Ht_array_tryboth(fit_s_asym, N, y),
  dBEKK_asym = bekk_Ht_array_tryboth(fit_d_asym, N, y),
  DCC_full = rmgarch::rcov(dcc_fit),
  DCC_scalar_stage = Ht_stage
)

if (!is.null(Ht_cdcc)) {
  Ht_list_models$cDCC_Aielli_stage <- Ht_cdcc
}

if (dim(Ht_list_models$DCC_full)[1] != N) {
  Ht_list_models$DCC_full <- Ht_list_models$DCC_full[seq_len(N), seq_len(N), , drop = FALSE]
}

if (dim(Ht_list_models$DCC_scalar_stage)[1] != N) {
  Ht_list_models$DCC_scalar_stage <- Ht_list_models$DCC_scalar_stage[seq_len(N), seq_len(N), , drop = FALSE]
}

if ("cDCC_Aielli_stage" %in% names(Ht_list_models) && dim(Ht_list_models$cDCC_Aielli_stage)[1] != N) {
  Ht_list_models$cDCC_Aielli_stage <- Ht_list_models$cDCC_Aielli_stage[seq_len(N), seq_len(N), , drop = FALSE]
}

k_dcc_full <- NA_integer_
coef_dcc <- try(length(coef(dcc_fit)), silent = TRUE)
if (!inherits(coef_dcc, "try-error")) k_dcc_full <- as.integer(coef_dcc)

k_univ_total <- npar_multifit_total(fit_g)
k_stage_2par <- if (is.na(k_univ_total)) NA_integer_ else as.integer(k_univ_total + 2L)

k_list <- c(
  sBEKK_sym = npar_bekk(fit_s_sym),
  dBEKK_sym = npar_bekk(fit_d_sym),
  sBEKK_asym = npar_bekk(fit_s_asym),
  dBEKK_asym = npar_bekk(fit_d_asym),
  DCC_full = k_dcc_full,
  DCC_scalar_stage = k_stage_2par,
  cDCC_Aielli_stage = k_stage_2par
)

model_comparison_table <- data.frame(
  Model = names(Ht_list_models),
  Nobs = NA_integer_,
  Npar = as.integer(k_list[names(Ht_list_models)]),
  LogLik_Gauss = NA_real_,
  AvgLogLik_Gauss = NA_real_,
  AIC_Gauss = NA_real_,
  BIC_Gauss = NA_real_,
  stringsAsFactors = FALSE
)

for (i in seq_len(nrow(model_comparison_table))) {
  nm <- model_comparison_table$Model[i]
  Ht <- Ht_list_models[[nm]]
  Tn <- dim(Ht)[3]
  ll <- gauss_loglik_from_Ht(y, Ht)
  
  model_comparison_table$Nobs[i] <- Tn
  model_comparison_table$LogLik_Gauss[i] <- ll
  model_comparison_table$AvgLogLik_Gauss[i] <- ll / Tn
  model_comparison_table$AIC_Gauss[i] <- aic_from_ll(ll, model_comparison_table$Npar[i])
  model_comparison_table$BIC_Gauss[i] <- bic_from_ll(ll, model_comparison_table$Npar[i], Tn)
}

model_comparison_table <- model_comparison_table |>
  dplyr::arrange(BIC_Gauss)

print(model_comparison_table)
save_table(model_comparison_table, "Table_Model_Comparison_Gaussian_LL_AIC_BIC.csv")

ll_term_decomposition <- function(y, Ht, eps_pd = 1e-12) {
  y <- as.matrix(y)
  Tn <- dim(Ht)[3]
  y <- y[seq_len(Tn), , drop = FALSE]
  
  logdet_vec <- rep(NA_real_, Tn)
  quad_vec <- rep(NA_real_, Tn)
  
  for (tt in seq_len(Tn)) {
    H <- (Ht[, , tt] + t(Ht[, , tt])) / 2
    ev <- eigen(H, symmetric = TRUE, only.values = TRUE)$values
    
    if (any(!is.finite(ev)) || any(ev <= eps_pd)) next
    
    logdet_vec[tt] <- as.numeric(determinant(H, logarithm = TRUE)$modulus)
    quad_vec[tt] <- drop(crossprod(y[tt, ], solve(H, y[tt, ])))
  }
  
  c(
    mean_logdet = mean(logdet_vec, na.rm = TRUE),
    mean_quad = mean(quad_vec, na.rm = TRUE)
  )
}

ll_terms <- dplyr::bind_rows(lapply(names(Ht_list_models), function(nm) {
  out <- ll_term_decomposition(y, Ht_list_models[[nm]])
  data.frame(Model = nm, t(out), row.names = NULL)
}))

print(ll_terms)
save_table(ll_terms, "Table_LL_Term_Decomposition.csv")

################################################################################
# 12) Final summary
################################################################################

cat("\n=============================\n")
cat("ESTIMATION AND DIAGNOSTICS COMPLETE\n")
cat("=============================\n")

cat("Saved output directory:\n")
cat(output_dir, "\n")

cat("\nSaved volatility object:\n")
cat(" - vol_models_all_7models.rds\n")

cat("\nSaved diagnostic/model-comparison tables:\n")
cat(" - Table_Hosking_standardized_residuals.csv\n")
cat(" - Table_LjungBox_squared_standardized_residuals.csv\n")
cat(" - Table_Ht_validity_summary.csv\n")
cat(" - Table_Model_Comparison_Gaussian_LL_AIC_BIC.csv\n")
cat(" - Table_LL_Term_Decomposition.csv\n")

cat("\nTwo-step DCC parameters:\n")
print(dcc_stage_ab)



