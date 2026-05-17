################################################################################
# DROP-IN SCRIPT — BIDIR CECI engine + rolling plots + joint-stress heatmaps
#
# Robust version:
#   - Cleans duplicate Date indices before all model runs
#   - Performs dlog_var transform on numeric matrix, not directly on zoo
#   - Avoids fragile zoo merge() for CECI series alignment
#   - Computes BIDIR CECI:
#       BIDIR = 100*sum(theta[E,C]) + 100*sum(theta[C,E])
################################################################################

suppressPackageStartupMessages({
  library(zoo)
  library(ConnectednessApproach)
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  library(scales)
  library(grid)
  library(gridExtra)
  library(gtable)
})

# ==============================================================================
# 1) USER SETTINGS
# ==============================================================================

base_dir <- "C:/Users/sezue/OneDrive/Desktop/BA"

output_dir <- file.path(base_dir, "outputs")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

vol_models_file <- file.path(output_dir, "vol_models_all_7models.rds")
if (!file.exists(vol_models_file)) {
  vol_models_file <- file.path(base_dir, "vol_models_all_7models.rds")
}

DIAG <- FALSE

models_used <- c(
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage",
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym"
)

commodities4 <- c(
  "TTF (Gas)",
  "Brent (Oil)",
  "API2 (Coal)",
  "MO1 (Carbon)"
)

commodity_short_labels <- c(
  "TTF (Gas)"    = "Gas",
  "Brent (Oil)"  = "Oil",
  "API2 (Coal)"  = "Coal",
  "MO1 (Carbon)" = "Carbon"
)

mode_static  <- "dlog_var"
mode_rolling <- "dlog_var"

ceci_direction <- "BIDIR"
equity_regex   <- "^SX"

W     <- 200
nlag  <- 1
nfore <- 20

regime_model <- "DCC_full"
rolling_commodity_for_universe <- "TTF (Gas)"
benchmark_model <- "DCC_full"

include_panels <- c("joint_stress", "ceci", "dev")

vol_panel_transform <- "level"
equity_avg_top5     <- TRUE

joint_stress_method    <- "product"
joint_stress_transform <- "log"

downside_window <- 20
downside_as_vol <- TRUE
downside_stress_method <- "product"
downside_transform <- "log"

legend_ncol <- 3
legend_cex  <- 0.85
lwd         <- 1.35

cex_axis <- 1.25
cex_lab  <- 1.45
cex_main <- 1.15
cex_leg  <- 1.40

left_margin_lines  <- 5.8
right_margin_lines <- 6
legend_inset_y     <- -0.36

shade_stress_episodes <- TRUE
label_stress_episodes <- TRUE

episodes_named <- data.frame(
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

episode_label_cex <- 0.92
episode_label_col <- "black"

heatmap_models_used <- models_used
heatmap_equity_top5 <- TRUE

heatmap_joint_stress_method    <- "product"
heatmap_joint_stress_transform <- "log"

out_png_main <- file.path(
  output_dir,
  "RollingCECI_BIDIR_and_StaticCECI_modelsUsed_commodities4.png"
)

out_png_compare <- file.path(
  output_dir,
  "RollingCECI_BIDIR_ModelComparison_VOL_JOINT_DOWNSIDE_CECI_DEV_withStressEpisodes.png"
)

out_png_heatmap <- file.path(
  output_dir,
  "Heatmap_DeltaBucketMean_from_OverallMean_BIDIR_CECI_6Panel_equalPanels.png"
)

out_csv_heatmap <- file.path(
  output_dir,
  "Heatmap_DeltaBucketMean_from_OverallMean_BIDIR_CECI_6Panel_equalPanels.csv"
)

# ==============================================================================
# 2) SAFE HELPERS
# ==============================================================================

save_png <- function(filename, expr, width = 2600, height = 2300, res = 200) {
  if (is.null(filename)) {
    eval.parent(substitute(expr))
    return(invisible(NULL))
  }
  
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

as_date_index <- function(idx, object_name = "zoo_object") {
  if (inherits(idx, "Date")) {
    return(idx)
  }
  
  if (inherits(idx, c("POSIXct", "POSIXt"))) {
    return(as.Date(idx))
  }
  
  if (inherits(idx, c("yearmon", "yearqtr"))) {
    return(as.Date(idx))
  }
  
  if (is.numeric(idx)) {
    return(as.Date(idx, origin = "1970-01-01"))
  }
  
  idx2 <- suppressWarnings(as.Date(idx))
  
  if (all(is.na(idx2))) {
    stop(
      object_name,
      ": could not coerce index to Date. Index class: ",
      paste(class(idx), collapse = ", ")
    )
  }
  
  idx2
}

clean_zoo <- function(z, object_name = "zoo_object") {
  stopifnot(inherits(z, "zoo"))
  
  idx <- as_date_index(zoo::index(z), object_name)
  cd <- zoo::coredata(z)
  
  if (is.null(dim(cd))) {
    cd <- matrix(cd, ncol = 1)
    if (!is.null(colnames(z))) colnames(cd) <- colnames(z)
  }
  
  ok <- !is.na(idx)
  idx <- idx[ok]
  cd <- cd[ok, , drop = FALSE]
  
  ord <- order(idx)
  idx <- idx[ord]
  cd <- cd[ord, , drop = FALSE]
  
  if (anyDuplicated(idx)) {
    warning(
      object_name,
      ": duplicate Date index entries detected. Keeping last row per Date."
    )
    
    keep <- unlist(
      lapply(
        split(seq_along(idx), idx),
        function(ii) tail(ii, 1)
      ),
      use.names = FALSE
    )
    
    keep <- sort(keep)
    idx <- idx[keep]
    cd <- cd[keep, , drop = FALSE]
  }
  
  out <- zoo::zoo(cd, order.by = idx)
  colnames(out) <- colnames(cd)
  out
}

ensure_vol_colnames <- function(vol_zoo, model_name = "model") {
  stopifnot(inherits(vol_zoo, "zoo"))
  
  cn <- colnames(vol_zoo)
  if (is.null(cn)) stop(model_name, ": volatility object has no column names.")
  
  colnames(vol_zoo) <- ifelse(grepl("_vol$", cn), cn, paste0(cn, "_vol"))
  vol_zoo
}

ensure_ret_colnames <- function(ret_zoo, model_name = "model") {
  stopifnot(inherits(ret_zoo, "zoo"))
  
  cn <- colnames(ret_zoo)
  if (is.null(cn)) stop(model_name, ": return object has no column names.")
  
  colnames(ret_zoo) <- ifelse(grepl("_ret$", cn), cn, paste0(cn, "_ret"))
  ret_zoo
}

strip_vol <- function(x) sub("_vol$", "", x)
strip_ret <- function(x) sub("_ret$", "", x)

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

align_to_index <- function(z, idx, object_name = "zoo_object") {
  z <- clean_zoo(z, object_name)
  idx <- as.Date(idx)
  
  cd <- zoo::coredata(z)
  if (is.null(dim(cd))) cd <- matrix(cd, ncol = 1)
  
  pos <- match(as.character(idx), as.character(zoo::index(z)))
  
  out <- matrix(NA_real_, nrow = length(idx), ncol = NCOL(cd))
  ok <- !is.na(pos)
  
  if (any(ok)) {
    out[ok, ] <- cd[pos[ok], , drop = FALSE]
  }
  
  colnames(out) <- colnames(cd)
  zoo::zoo(out, order.by = idx)
}

merge_zoo_list_inner <- function(zlist, names_out = names(zlist), object_name = "merged_zoo") {
  stopifnot(is.list(zlist), length(zlist) >= 1)
  
  zlist <- lapply(seq_along(zlist), function(i) {
    nm <- if (!is.null(names_out) && length(names_out) >= i) names_out[i] else paste0("series_", i)
    z <- clean_zoo(zlist[[i]], paste0(object_name, "_", nm))
    cd <- zoo::coredata(z)
    if (is.null(dim(cd))) {
      cd <- matrix(cd, ncol = 1)
    }
    zoo::zoo(as.numeric(cd[, 1]), order.by = zoo::index(z))
  })
  
  common_idx <- Reduce(
    intersect,
    lapply(zlist, function(z) as.character(zoo::index(z)))
  )
  
  common_idx <- as.Date(common_idx)
  common_idx <- sort(common_idx)
  
  if (length(common_idx) == 0) {
    stop(object_name, ": no common dates across series.")
  }
  
  mat <- matrix(NA_real_, nrow = length(common_idx), ncol = length(zlist))
  
  for (j in seq_along(zlist)) {
    zj <- zlist[[j]]
    pos <- match(as.character(common_idx), as.character(zoo::index(zj)))
    mat[, j] <- as.numeric(zoo::coredata(zj))[pos]
  }
  
  colnames(mat) <- names_out
  zoo::zoo(mat, order.by = common_idx)
}

subset_by_dates_zoo <- function(Z, dates) {
  Z <- clean_zoo(Z, "subset_by_dates_zoo_input")
  dates <- as.Date(dates)
  
  if (length(dates) == 0) {
    return(Z[0, , drop = FALSE])
  }
  
  keep <- as.character(zoo::index(Z)) %in% as.character(dates)
  Z[keep, , drop = FALSE]
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

zscore <- function(x) {
  xv <- as.numeric(zoo::coredata(x))
  mu <- mean(xv, na.rm = TRUE)
  sdv <- stats::sd(xv, na.rm = TRUE)
  
  if (!is.finite(sdv) || sdv == 0) {
    return(rep(NA_real_, length(xv)))
  }
  
  (xv - mu) / sdv
}

joint_stress <- function(eq, com, method = c("product", "max", "zsum")) {
  method <- match.arg(method)
  
  eqv  <- as.numeric(zoo::coredata(eq))
  comv <- as.numeric(zoo::coredata(com))
  
  out <- switch(
    method,
    product = eqv * comv,
    max     = pmax(eqv, comv),
    zsum    = zscore(eq) + zscore(com)
  )
  
  zoo::zoo(out, order.by = zoo::index(eq))
}

roll_downside_semivar <- function(r, width = 20) {
  rr <- as.numeric(zoo::coredata(r))
  ds <- (pmin(rr, 0))^2
  
  z <- zoo::zoo(ds, order.by = zoo::index(r))
  
  zoo::zoo(
    zoo::rollapply(
      z,
      width = width,
      FUN = function(x) mean(x, na.rm = TRUE),
      align = "right",
      fill = NA
    ),
    order.by = zoo::index(r)
  )
}

default_cols <- function(models) {
  base <- c(
    sBEKK_sym         = "black",
    dBEKK_sym         = "dodgerblue3",
    dBEKK_asym        = "firebrick2",
    DCC_full          = "darkgreen",
    DCC_scalar_stage  = "purple3",
    cDCC_Aielli_stage = "goldenrod2"
  )
  
  if (!all(models %in% names(base))) {
    pal <- rep(
      c("black", "dodgerblue3", "firebrick2", "darkgreen", "purple3", "goldenrod2"),
      length.out = length(models)
    )
    names(pal) <- models
    return(pal)
  }
  
  base[models]
}

default_ltys <- function(models) {
  base <- c(
    sBEKK_sym         = 1,
    dBEKK_sym         = 2,
    dBEKK_asym        = 4,
    DCC_full          = 1,
    DCC_scalar_stage  = 6,
    cDCC_Aielli_stage = 5
  )
  
  if (!all(models %in% names(base))) {
    lt <- rep(1:6, length.out = length(models))
    names(lt) <- models
    return(lt)
  }
  
  base[models]
}

shade_episodes <- function(episodes,
                           shade_col = gray(0.90, alpha = 0.95),
                           x_range = NULL) {
  if (is.null(episodes) || nrow(episodes) == 0) return(invisible(NULL))
  
  usr <- par("usr")
  
  for (i in seq_len(nrow(episodes))) {
    xleft  <- episodes$start[i]
    xright <- episodes$end[i]
    
    if (!is.null(x_range)) {
      xleft  <- max(xleft,  min(x_range))
      xright <- min(xright, max(x_range))
      if (xleft > xright) next
    }
    
    rect(
      xleft   = xleft,
      ybottom = usr[3],
      xright  = xright,
      ytop    = usr[4],
      col     = shade_col,
      border  = NA
    )
  }
  
  invisible(NULL)
}

label_episodes_top <- function(episodes,
                               label_x_shift_days = NULL,
                               label_y_frac_from_top = NULL,
                               cex = 0.92,
                               col = "black",
                               x_range = NULL) {
  if (is.null(episodes) || nrow(episodes) == 0) return(invisible(NULL))
  
  n <- nrow(episodes)
  
  if (is.null(label_x_shift_days) || length(label_x_shift_days) != n) {
    label_x_shift_days <- rep(0, n)
  }
  
  if (is.null(label_y_frac_from_top) || length(label_y_frac_from_top) != n) {
    label_y_frac_from_top <- rep(0.025, n)
  }
  
  usr <- par("usr")
  y_span <- diff(usr[3:4])
  
  for (i in seq_len(n)) {
    x_mid <- episodes$start[i] +
      floor(as.numeric(episodes$end[i] - episodes$start[i]) / 2) +
      label_x_shift_days[i]
    
    if (!is.null(x_range)) {
      if (x_mid < min(x_range) || x_mid > max(x_range)) next
    }
    
    y_lab <- usr[4] - label_y_frac_from_top[i] * y_span
    
    text(
      x = x_mid,
      y = y_lab,
      labels = episodes$label[i],
      cex = cex,
      col = col
    )
  }
  
  invisible(NULL)
}

make_stress_bucket <- function(x) {
  out <- rep(NA_character_, length(x))
  ok <- is.finite(x)
  levels_out <- c("0-25", "25-50", "50-75", "75-100")
  
  if (sum(ok) == 0) return(factor(out, levels = levels_out))
  
  if (sum(ok) == 1) {
    out[ok] <- "75-100"
    return(factor(out, levels = levels_out))
  }
  
  r <- rank(x[ok], na.last = "keep", ties.method = "average")
  p <- (r - 1) / (length(r) - 1)
  
  out_ok <- rep(NA_character_, length(p))
  out_ok[p <= 0.25] <- "0-25"
  out_ok[p > 0.25 & p <= 0.50] <- "25-50"
  out_ok[p > 0.50 & p <= 0.75] <- "50-75"
  out_ok[p > 0.75] <- "75-100"
  
  out[ok] <- out_ok
  factor(out, levels = levels_out)
}

# ==============================================================================
# 3) BIDIR CECI ENGINE
# ==============================================================================

prep_input <- function(vol_zoo,
                       mode = c("dlog_var", "dlog_vol", "dlog", "level", "loglevel"),
                       object_name = "connectedness_input") {
  mode <- match.arg(mode)
  
  vol_zoo <- clean_zoo(vol_zoo, object_name)
  
  idx <- zoo::index(vol_zoo)
  X <- zoo::coredata(vol_zoo)
  if (is.null(dim(X))) X <- matrix(X, ncol = 1)
  
  X <- pmax(X, 1e-12)
  
  if (mode == "level") {
    Y <- X
    yidx <- idx
  } else if (mode == "loglevel") {
    Y <- log(X)
    yidx <- idx
  } else if (mode == "dlog" || mode == "dlog_vol") {
    Y <- diff(log(X))
    yidx <- idx[-1]
  } else if (mode == "dlog_var") {
    Y <- diff(log(X^2))
    yidx <- idx[-1]
  }
  
  colnames(Y) <- colnames(vol_zoo)
  
  ok <- stats::complete.cases(Y)
  Y <- Y[ok, , drop = FALSE]
  yidx <- yidx[ok]
  
  out <- zoo::zoo(Y, order.by = yidx)
  colnames(out) <- colnames(vol_zoo)
  clean_zoo(out, paste0(object_name, "_prepared"))
}

subset_by_base <- function(vol_zoo, base_keep) {
  cols <- paste0(base_keep, "_vol")
  missing <- setdiff(cols, colnames(vol_zoo))
  
  if (length(missing) > 0) {
    stop(
      "Missing columns: ",
      paste(missing, collapse = ", "),
      "\nAvailable: ",
      paste(colnames(vol_zoo), collapse = ", ")
    )
  }
  
  vol_zoo[, cols, drop = FALSE]
}

run_conn <- function(X, nlag = 1, nfore = 20, window.size = NULL) {
  ConnectednessApproach(
    x = X,
    nlag = nlag,
    nfore = nfore,
    model = "VAR",
    connectedness = "Time",
    window.size = window.size
  )
}

run_rolling_conn <- function(X, nlag = 1, nfore = 20, window.size = 200) {
  ConnectednessApproach(
    x = X,
    nlag = nlag,
    nfore = nfore,
    model = "VAR",
    connectedness = "Time",
    window.size = window.size
  )
}

get_theta_cube_from_dca <- function(dca_obj) {
  cand <- c("GFEVD", "FEVD", "Theta", "theta", "CT", "ct", "TABLE")
  
  for (nm in cand) {
    if (!is.null(dca_obj[[nm]])) {
      obj <- dca_obj[[nm]]
      
      if (is.array(obj) && length(dim(obj)) == 3) {
        return(obj)
      }
      
      if (is.list(obj)) {
        inner_cand <- c("GFEVD", "FEVD", "Theta", "theta")
        
        for (jj in inner_cand) {
          if (!is.null(obj[[jj]]) && is.array(obj[[jj]]) && length(dim(obj[[jj]])) == 3) {
            return(obj[[jj]])
          }
        }
      }
    }
  }
  
  stop(
    "Could not find an N x N x T FEVD/GFEVD array. Available names: ",
    paste(names(dca_obj), collapse = ", ")
  )
}

get_theta_matrix_from_dca <- function(dca_obj) {
  cand <- c("GFEVD", "FEVD", "Theta", "theta", "CT", "ct", "TABLE")
  
  for (nm in cand) {
    if (!is.null(dca_obj[[nm]])) {
      obj <- dca_obj[[nm]]
      
      if (is.matrix(obj) && nrow(obj) == ncol(obj)) {
        return(obj)
      }
      
      if (is.array(obj) && length(dim(obj)) == 3) {
        return(obj[, , dim(obj)[3]])
      }
      
      if (is.list(obj)) {
        inner_cand <- c("GFEVD", "FEVD", "Theta", "theta")
        
        for (jj in inner_cand) {
          if (!is.null(obj[[jj]])) {
            inner <- obj[[jj]]
            
            if (is.matrix(inner) && nrow(inner) == ncol(inner)) {
              return(inner)
            }
            
            if (is.array(inner) && length(dim(inner)) == 3) {
              return(inner[, , dim(inner)[3]])
            }
          }
        }
      }
    }
  }
  
  stop(
    "Could not find FEVD/GFEVD matrix. Available names: ",
    paste(names(dca_obj), collapse = ", ")
  )
}

build_directional_ceci_from_theta_matrix <- function(theta_mat,
                                                     X_input,
                                                     commodity_base,
                                                     equity_regex = "^SX",
                                                     scale_100 = TRUE,
                                                     diagnostics = FALSE) {
  theta <- as.matrix(theta_mat)
  
  asset_names <- rownames(theta)
  
  if (is.null(asset_names)) {
    asset_names <- colnames(theta)
  }
  
  if (is.null(asset_names)) {
    asset_names <- colnames(X_input)
  }
  
  if (is.null(asset_names)) {
    stop("Could not infer asset names for theta matrix.")
  }
  
  base_names <- strip_vol(asset_names)
  
  com_base_hit <- match_base_name(commodity_base, base_names)
  
  if (is.na(com_base_hit)) {
    stop(
      "Commodity '", commodity_base, "' not found in theta assets. Example assets: ",
      paste(head(asset_names, 8), collapse = ", ")
    )
  }
  
  com_asset <- asset_names[match(com_base_hit, base_names)][1]
  com_idx <- which(asset_names == com_asset)
  
  if (length(com_idx) != 1) {
    stop("Could not uniquely locate commodity asset: ", com_asset)
  }
  
  eq_idx <- grep(equity_regex, base_names)
  
  if (length(eq_idx) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "'.")
  }
  
  e_to_c <- sum(theta[eq_idx, com_idx], na.rm = TRUE)
  c_to_e <- sum(theta[com_idx, eq_idx], na.rm = TRUE)
  
  if (isTRUE(scale_100)) {
    e_to_c <- 100 * e_to_c
    c_to_e <- 100 * c_to_e
  }
  
  out <- c(
    E_to_C = e_to_c,
    C_to_E = c_to_e,
    BIDIR  = e_to_c + c_to_e
  )
  
  if (isTRUE(diagnostics)) {
    cat("\n[BIDIR CECI diagnostics]\n")
    cat("Commodity:", com_asset, "\n")
    cat("Equities :", paste(asset_names[eq_idx], collapse = ", "), "\n")
    print(out)
  }
  
  out
}

compute_ceci_from_dca <- function(dca_obj,
                                  X_input,
                                  commodity_base,
                                  direction = c("BIDIR", "E_to_C", "C_to_E"),
                                  equity_regex = "^SX",
                                  diagnostics = FALSE) {
  direction <- match.arg(direction)
  
  theta_mat <- get_theta_matrix_from_dca(dca_obj)
  
  suite <- build_directional_ceci_from_theta_matrix(
    theta_mat = theta_mat,
    X_input = X_input,
    commodity_base = commodity_base,
    equity_regex = equity_regex,
    scale_100 = TRUE,
    diagnostics = diagnostics
  )
  
  unname(suite[direction])
}

compute_rolling_ceci_series <- function(roll_obj,
                                        X_input,
                                        commodity_base,
                                        direction = c("BIDIR", "E_to_C", "C_to_E"),
                                        equity_regex = "^SX",
                                        diagnostics = FALSE) {
  direction <- match.arg(direction)
  
  theta <- get_theta_cube_from_dca(roll_obj)
  
  Tn <- dim(theta)[3]
  idx <- tail(zoo::index(X_input), Tn)
  
  ceci <- rep(NA_real_, Tn)
  
  for (tt in seq_len(Tn)) {
    theta_mat <- theta[, , tt]
    
    suite <- try(
      build_directional_ceci_from_theta_matrix(
        theta_mat = theta_mat,
        X_input = X_input,
        commodity_base = commodity_base,
        equity_regex = equity_regex,
        scale_100 = TRUE,
        diagnostics = FALSE
      ),
      silent = TRUE
    )
    
    if (inherits(suite, "try-error")) {
      ceci[tt] <- NA_real_
    } else {
      ceci[tt] <- unname(suite[direction])
    }
  }
  
  z <- zoo::zoo(ceci, order.by = as.Date(idx))
  clean_zoo(z, paste0("rolling_CECI_", commodity_base))
}

run_rolling_CECI_all_models <- function(vol_models,
                                        models,
                                        commodity_asset,
                                        regime_model = "DCC_full",
                                        mode = c("dlog_var", "dlog_vol", "dlog", "loglevel", "level"),
                                        nlag = 1,
                                        nfore = 20,
                                        W = 200,
                                        ceci_direction = c("BIDIR", "E_to_C", "C_to_E"),
                                        equity_regex = "^SX",
                                        diagnostics = FALSE) {
  mode <- match.arg(mode)
  ceci_direction <- match.arg(ceci_direction)
  
  vol_r <- clean_zoo(ensure_vol_colnames(vol_models[[regime_model]], regime_model), regime_model)
  base_cols <- strip_vol(colnames(vol_r))
  
  equities_base <- base_cols[grepl(equity_regex, base_cols)]
  
  if (length(equities_base) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "' in regime_model columns.")
  }
  
  assets_keep <- unique(c(equities_base, commodity_asset))
  
  rolling <- list()
  X_inputs <- list()
  CECI <- list()
  
  for (m in models) {
    cat("Starting rolling BIDIR CECI:", commodity_asset, "|", m, "\n")
    
    vol_m <- clean_zoo(ensure_vol_colnames(vol_models[[m]], m), m)
    
    X <- prep_input(
      subset_by_base(vol_m, assets_keep),
      mode = mode,
      object_name = paste0("input_", commodity_asset, "_", m)
    )
    
    if (NROW(X) <= W + nlag + 5) {
      stop(
        "Too few observations for ", m,
        " after cleaning. N=", NROW(X),
        ", W=", W,
        ", nlag=", nlag
      )
    }
    
    X_inputs[[m]] <- X
    
    roll_obj <- run_rolling_conn(
      X,
      nlag = nlag,
      nfore = nfore,
      window.size = W
    )
    
    rolling[[m]] <- roll_obj
    
    CECI[[m]] <- compute_rolling_ceci_series(
      roll_obj = roll_obj,
      X_input = X,
      commodity_base = commodity_asset,
      direction = ceci_direction,
      equity_regex = equity_regex,
      diagnostics = diagnostics
    )
    
    cat("Rolling", ceci_direction, "CECI done:", mode, "|", commodity_asset, "|", m, "\n")
  }
  
  list(
    rolling = rolling,
    X_inputs = X_inputs,
    CECI = CECI,
    assets_keep = assets_keep,
    ceci_direction = ceci_direction,
    input_mode = mode
  )
}

plot_rolling_CECI <- function(rolling_obj,
                              main = NULL,
                              lwd = 1.35,
                              legend_cex = 0.85,
                              legend_ncol = 4,
                              legend_mar_lines = 7,
                              legend_inset_y = -0.22) {
  CECI <- rolling_obj$CECI
  models <- names(CECI)
  
  Z <- merge_zoo_list_inner(CECI, models, "plot_rolling_CECI")
  
  cols <- default_cols(models)
  ltys <- default_ltys(models)
  
  if (is.null(main)) {
    main <- paste0("Rolling ", rolling_obj$ceci_direction, " CECI")
  }
  
  op <- par(no.readonly = TRUE)
  on.exit(par(op), add = TRUE)
  
  par(mar = c(legend_mar_lines, 4, 4, 2) + 0.1)
  
  plot(
    zoo::index(Z),
    zoo::coredata(Z[, 1]),
    type = "l",
    xlab = "",
    ylab = "CECI",
    main = main,
    col = cols[1],
    lty = ltys[1],
    lwd = lwd
  )
  
  if (NCOL(Z) > 1) {
    for (j in 2:NCOL(Z)) {
      nm <- colnames(Z)[j]
      
      lines(
        zoo::index(Z),
        zoo::coredata(Z[, j]),
        col = cols[nm],
        lty = ltys[nm],
        lwd = lwd
      )
    }
  }
  
  par(xpd = TRUE)
  
  legend(
    "bottom",
    legend = colnames(Z),
    col = cols[colnames(Z)],
    lty = ltys[colnames(Z)],
    lwd = lwd,
    bty = "n",
    cex = legend_cex,
    ncol = legend_ncol,
    inset = c(0, legend_inset_y)
  )
  
  invisible(Z)
}

get_regime_indices <- function(X,
                               base_col_vol,
                               q_low = 0.10,
                               q_high = 0.90,
                               qmid_lo = 0.45,
                               qmid_hi = 0.55) {
  x <- as.numeric(zoo::coredata(X[, base_col_vol]))
  
  ql <- quantile(x, q_low,   na.rm = TRUE)
  qh <- quantile(x, q_high,  na.rm = TRUE)
  q1 <- quantile(x, qmid_lo, na.rm = TRUE)
  q2 <- quantile(x, qmid_hi, na.rm = TRUE)
  
  list(
    low  = which(x <= ql),
    mid  = which(x > q1 & x < q2),
    high = which(x >= qh)
  )
}

adjust_idx <- function(idx, mode = c("dlog_var", "dlog_vol", "dlog", "level", "loglevel")) {
  mode <- match.arg(mode)
  
  if (mode %in% c("level", "loglevel")) return(idx)
  
  list(
    low  = idx$low[idx$low > 1] - 1,
    mid  = idx$mid[idx$mid > 1] - 1,
    high = idx$high[idx$high > 1] - 1
  )
}

get_regime_dates <- function(X,
                             base_col_vol,
                             regime = c("low", "mid", "high"),
                             q_low = 0.10,
                             q_high = 0.90,
                             qmid_lo = 0.45,
                             qmid_hi = 0.55,
                             mode = c("dlog_var", "dlog_vol", "dlog", "level", "loglevel")) {
  regime <- match.arg(regime)
  mode <- match.arg(mode)
  
  idx0 <- get_regime_indices(
    X,
    base_col_vol,
    q_low = q_low,
    q_high = q_high,
    qmid_lo = qmid_lo,
    qmid_hi = qmid_hi
  )
  
  idx0 <- adjust_idx(idx0, mode = mode)
  
  idx_vec <- idx0[[regime]]
  idx_vec <- idx_vec[idx_vec >= 1 & idx_vec <= nrow(X)]
  
  zoo::index(X)[idx_vec]
}

ceci_by_regime_one_commodity <- function(vol_models,
                                         commodity_asset,
                                         models,
                                         regime_model = "DCC_full",
                                         mode = "dlog_var",
                                         regimes = c("low", "mid", "high"),
                                         q_low = 0.10,
                                         q_high = 0.90,
                                         qmid_lo = 0.45,
                                         qmid_hi = 0.55,
                                         nlag = 1,
                                         nfore = 20,
                                         ceci_direction = c("BIDIR", "E_to_C", "C_to_E"),
                                         equity_regex = "^SX",
                                         diagnostics = FALSE) {
  ceci_direction <- match.arg(ceci_direction)
  
  miss <- setdiff(unique(c(models, regime_model)), names(vol_models))
  if (length(miss) > 0) stop("vol_models is missing: ", paste(miss, collapse = ", "))
  
  vol_r <- clean_zoo(ensure_vol_colnames(vol_models[[regime_model]], regime_model), regime_model)
  
  base_cols <- strip_vol(colnames(vol_r))
  equities_base <- base_cols[grepl(equity_regex, base_cols)]
  
  if (length(equities_base) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "' in regime_model columns.")
  }
  
  assets_keep <- unique(c(equities_base, commodity_asset))
  regime_col_vol <- paste0(commodity_asset, "_vol")
  
  X_reg <- prep_input(
    subset_by_base(vol_r, assets_keep),
    mode = mode,
    object_name = paste0("regime_input_", commodity_asset)
  )
  
  if (!(regime_col_vol %in% colnames(X_reg))) {
    stop("Regime asset column not found: ", regime_col_vol)
  }
  
  out <- list()
  
  for (rg in regimes) {
    reg_dates <- get_regime_dates(
      X_reg,
      base_col_vol = regime_col_vol,
      regime = rg,
      q_low = q_low,
      q_high = q_high,
      qmid_lo = qmid_lo,
      qmid_hi = qmid_hi,
      mode = mode
    )
    
    for (m in models) {
      vol_m <- clean_zoo(ensure_vol_colnames(vol_models[[m]], m), m)
      
      X_full <- prep_input(
        subset_by_base(vol_m, assets_keep),
        mode = mode,
        object_name = paste0("static_input_", commodity_asset, "_", m)
      )
      
      Xr <- subset_by_dates_zoo(X_full, reg_dates)
      
      if (nrow(Xr) < (nlag + 10)) {
        out[[length(out) + 1]] <- data.frame(
          Commodity = commodity_asset,
          Model = m,
          Regime = rg,
          CECI = NA_real_,
          RegimeDaysUsed = nrow(Xr),
          N = ncol(Xr),
          stringsAsFactors = FALSE
        )
        next
      }
      
      dca <- run_conn(
        Xr,
        nlag = nlag,
        nfore = nfore,
        window.size = NULL
      )
      
      ceci <- compute_ceci_from_dca(
        dca_obj = dca,
        X_input = Xr,
        commodity_base = commodity_asset,
        direction = ceci_direction,
        equity_regex = equity_regex,
        diagnostics = diagnostics
      )
      
      out[[length(out) + 1]] <- data.frame(
        Commodity = commodity_asset,
        Model = m,
        Regime = rg,
        CECI = ceci,
        RegimeDaysUsed = nrow(Xr),
        N = ncol(Xr),
        stringsAsFactors = FALSE
      )
    }
  }
  
  do.call(rbind, out)
}

ceci_by_regime_all_commodities <- function(vol_models,
                                           commodities,
                                           models,
                                           regime_model = "DCC_full",
                                           mode = "dlog_var",
                                           regimes = c("low", "mid", "high"),
                                           q_low = 0.10,
                                           q_high = 0.90,
                                           qmid_lo = 0.45,
                                           qmid_hi = 0.55,
                                           nlag = 1,
                                           nfore = 20,
                                           ceci_direction = c("BIDIR", "E_to_C", "C_to_E"),
                                           equity_regex = "^SX",
                                           diagnostics = FALSE) {
  ceci_direction <- match.arg(ceci_direction)
  
  res <- lapply(commodities, function(com) {
    ceci_by_regime_one_commodity(
      vol_models = vol_models,
      commodity_asset = com,
      models = models,
      regime_model = regime_model,
      mode = mode,
      regimes = regimes,
      q_low = q_low,
      q_high = q_high,
      qmid_lo = qmid_lo,
      qmid_hi = qmid_hi,
      nlag = nlag,
      nfore = nfore,
      ceci_direction = ceci_direction,
      equity_regex = equity_regex,
      diagnostics = diagnostics
    )
  })
  
  df <- do.call(rbind, res)
  
  df$Regime <- factor(
    df$Regime,
    levels = c("low", "mid", "high"),
    labels = c("Low", "Normal", "High")
  )
  
  df$Model <- factor(df$Model, levels = models)
  df$Commodity <- factor(df$Commodity, levels = commodities)
  
  df[order(df$Commodity, df$Model, df$Regime), , drop = FALSE]
}

run_all_outputs <- function(vol_models,
                            models,
                            commodities,
                            regime_model = "DCC_full",
                            mode_static = "dlog_var",
                            mode_rolling = "dlog_var",
                            regimes = c("low", "mid", "high"),
                            q_low = 0.10,
                            q_high = 0.90,
                            qmid_lo = 0.45,
                            qmid_hi = 0.55,
                            nlag = 1,
                            nfore = 20,
                            W = 200,
                            rolling_commodity_for_universe = "Brent (Oil)",
                            ceci_direction = c("BIDIR", "E_to_C", "C_to_E"),
                            equity_regex = "^SX",
                            out_png = NULL,
                            diagnostics = FALSE) {
  ceci_direction <- match.arg(ceci_direction)
  
  miss <- setdiff(unique(c(models, regime_model)), names(vol_models))
  if (length(miss) > 0) stop("vol_models is missing: ", paste(miss, collapse = ", "))
  
  if (!is.null(out_png)) {
    grDevices::png(
      filename = out_png,
      width = 2600,
      height = 1700,
      res = 200,
      type = "cairo"
    )
    on.exit(try(grDevices::dev.off(), silent = TRUE), add = TRUE)
  }
  
  roll_obj <- run_rolling_CECI_all_models(
    vol_models = vol_models,
    models = models,
    commodity_asset = rolling_commodity_for_universe,
    regime_model = regime_model,
    mode = mode_rolling,
    nlag = nlag,
    nfore = nfore,
    W = W,
    ceci_direction = ceci_direction,
    equity_regex = equity_regex,
    diagnostics = diagnostics
  )
  
  plot_rolling_CECI(
    roll_obj,
    main = paste0(
      "Rolling ",
      ceci_direction,
      " CECI (",
      mode_rolling,
      ", window=",
      W,
      ") — Universe: SX* + ",
      rolling_commodity_for_universe
    ),
    legend_cex = 0.80
  )
  
  ceci_df <- ceci_by_regime_all_commodities(
    vol_models = vol_models,
    commodities = commodities,
    models = models,
    regime_model = regime_model,
    mode = mode_static,
    regimes = regimes,
    q_low = q_low,
    q_high = q_high,
    qmid_lo = qmid_lo,
    qmid_hi = qmid_hi,
    nlag = nlag,
    nfore = nfore,
    ceci_direction = ceci_direction,
    equity_regex = equity_regex,
    diagnostics = diagnostics
  )
  
  cat("\n=================================================\n")
  cat("STATIC", ceci_direction, "CECI BY COMMODITY × MODEL × REGIME\n")
  cat(
    "mode:", mode_static,
    "| regime_model:", regime_model,
    "| CECI direction:", ceci_direction,
    "\n"
  )
  cat("=================================================\n")
  
  print(ceci_df)
  
  invisible(list(
    rolling = roll_obj,
    tci = ceci_df,
    out_png = out_png,
    ceci_direction = ceci_direction,
    input_mode = mode_rolling
  ))
}

# ==============================================================================
# 4) LOAD AND CLEAN vol_models
# ==============================================================================

if (!exists("vol_models")) {
  if (!file.exists(vol_models_file)) {
    stop("vol_models not found and file does not exist: ", vol_models_file)
  }
  
  vol_models <- readRDS(vol_models_file)
}

vol_models <- stats::setNames(
  lapply(names(vol_models), function(nm) {
    clean_zoo(ensure_vol_colnames(vol_models[[nm]], nm), nm)
  }),
  names(vol_models)
)

cat("\nDate/index check after cleaning vol_models:\n")
print(
  data.frame(
    model = names(vol_models),
    start = sapply(vol_models, function(z) as.character(min(zoo::index(z)))),
    end   = sapply(vol_models, function(z) as.character(max(zoo::index(z)))),
    nobs  = sapply(vol_models, NROW),
    duplicated_dates = sapply(vol_models, function(z) anyDuplicated(zoo::index(z))),
    row.names = NULL
  )
)

missing_models <- setdiff(unique(c(models_used, regime_model)), names(vol_models))
if (length(missing_models) > 0) {
  stop("vol_models is missing: ", paste(missing_models, collapse = ", "))
}

# ==============================================================================
# 5) PLOTTING HELPERS
# ==============================================================================

extract_ceci_series <- function(run_res, model_name) {
  if (
    is.null(run_res$rolling) ||
    is.null(run_res$rolling$CECI) ||
    is.null(run_res$rolling$CECI[[model_name]])
  ) {
    stop("Could not find run_res$rolling$CECI[['", model_name, "']].")
  }
  
  z <- clean_zoo(run_res$rolling$CECI[[model_name]], paste0("CECI_", model_name))
  cd <- zoo::coredata(z)
  if (is.null(dim(cd))) {
    cd <- matrix(as.numeric(cd), ncol = 1)
  } else {
    cd <- matrix(as.numeric(cd[, 1]), ncol = 1)
  }
  
  out <- zoo::zoo(cd, order.by = zoo::index(z))
  colnames(out) <- "CECI"
  out
}

build_joint_stress_series <- function(vol_models,
                                      model_name,
                                      commodity_requested,
                                      idx_ceci,
                                      equity_top5 = TRUE,
                                      joint_stress_method = "product",
                                      joint_stress_transform = "log",
                                      equity_regex = "^SX") {
  if (!(model_name %in% names(vol_models))) {
    stop("Model not found in vol_models: ", model_name)
  }
  
  vb <- clean_zoo(ensure_vol_colnames(vol_models[[model_name]], model_name), model_name)
  base_b <- strip_vol(colnames(vb))
  
  com_base <- match_base_name(commodity_requested, base_b)
  
  if (is.na(com_base)) {
    stop(
      "Commodity not found in vol_models[['", model_name, "']] for requested: ",
      commodity_requested
    )
  }
  
  com_sig <- align_to_index(
    vb[, paste0(com_base, "_vol"), drop = FALSE],
    idx_ceci,
    object_name = model_name
  )
  
  eq_cols <- grep(equity_regex, base_b)
  if (length(eq_cols) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "' in vol_models[['", model_name, "']].")
  }
  
  eq_names <- base_b[eq_cols]
  
  eq_use <- if (isTRUE(equity_top5)) {
    eq_names[seq_len(min(5, length(eq_names)))]
  } else {
    eq_names
  }
  
  eq_mat <- align_to_index(
    vb[, paste0(eq_use, "_vol"), drop = FALSE],
    idx_ceci,
    object_name = model_name
  )
  
  eq_avg <- zoo::zoo(
    rowMeans(zoo::coredata(eq_mat), na.rm = TRUE),
    order.by = idx_ceci
  )
  
  joint <- joint_stress(eq_avg, com_sig[, 1], method = joint_stress_method)
  
  if (joint_stress_transform == "log") {
    if (joint_stress_method == "zsum") {
      warning("joint_stress_method='zsum' can be negative; log transform skipped.")
    } else {
      joint <- log1p(pmax(joint, 0))
    }
  }
  
  clean_zoo(joint, paste0("joint_stress_", model_name, "_", commodity_requested))
}

plot_rolling_CECI_with_stress <- function(rolling_obj,
                                          vol_models,
                                          commodity_requested,
                                          benchmark = "DCC_full",
                                          transform = "dlog_var",
                                          window = 200,
                                          main = NULL,
                                          lwd = 1.35,
                                          legend_ncol = 3,
                                          legend_inset_y = -0.34,
                                          outer_bottom_lines = 8,
                                          outer_top_lines = 3.5,
                                          vol_transform = c("level", "log"),
                                          equity_top5 = TRUE,
                                          right_margin_lines = 6,
                                          include_panels = c("vol", "joint_stress", "downside", "ceci", "dev"),
                                          joint_stress_method = c("product", "max", "zsum"),
                                          joint_stress_transform = c("level", "log"),
                                          downside_window = 20,
                                          downside_as_vol = TRUE,
                                          downside_stress_method = c("product", "max", "zsum"),
                                          downside_transform = c("level", "log"),
                                          ret_models = NULL,
                                          cex_axis = 1.25,
                                          cex_lab = 1.35,
                                          cex_main = 1.15,
                                          cex_leg = 1.10,
                                          left_margin_lines = 5.8,
                                          xaxis_style = c("years", "default"),
                                          shade_episodes_flag = TRUE,
                                          episodes_named = NULL,
                                          episode_shade_col = gray(0.90, alpha = 0.95),
                                          label_episodes = TRUE,
                                          label_x_shift_days = NULL,
                                          label_y_frac_from_top = NULL,
                                          episode_label_cex = 0.92,
                                          episode_label_col = "black",
                                          equity_regex = "^SX") {
  
  xaxis_style <- match.arg(xaxis_style)
  vol_transform <- match.arg(vol_transform)
  joint_stress_method <- match.arg(joint_stress_method)
  joint_stress_transform <- match.arg(joint_stress_transform)
  downside_stress_method <- match.arg(downside_stress_method)
  downside_transform <- match.arg(downside_transform)
  
  include_panels <- unique(include_panels)
  ok_panels <- c("vol", "joint_stress", "downside", "ceci", "dev")
  
  if (!all(include_panels %in% ok_panels)) {
    stop("include_panels must be subset of: ", paste(ok_panels, collapse = ", "))
  }
  
  CECI_list <- rolling_obj$CECI
  models <- names(CECI_list)
  
  Z <- merge_zoo_list_inner(CECI_list, models, "comparison_CECI")
  models <- colnames(Z)
  
  if (!(benchmark %in% models)) {
    stop("benchmark must be one of: ", paste(models, collapse = ", "))
  }
  
  if (is.null(main)) {
    main <- paste0(
      "Rolling BIDIR CECI (",
      transform,
      ", window=",
      window,
      ") — SX* + ",
      commodity_requested,
      " — benchmark: ",
      benchmark
    )
  }
  
  D <- Z
  for (nm in models) {
    D[, nm] <- Z[, nm] - Z[, benchmark]
  }
  
  vb <- clean_zoo(ensure_vol_colnames(vol_models[[benchmark]], benchmark), benchmark)
  base_b <- strip_vol(colnames(vb))
  
  com_base <- match_base_name(commodity_requested, base_b)
  if (is.na(com_base)) stop("Commodity not found in benchmark volatility model: ", commodity_requested)
  
  idx_ceci <- zoo::index(Z)
  x <- idx_ceci
  
  com_sig <- align_to_index(
    vb[, paste0(com_base, "_vol"), drop = FALSE],
    idx_ceci,
    object_name = benchmark
  )
  
  eq_cols <- grep(equity_regex, base_b)
  if (length(eq_cols) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "' in vol_models[[benchmark]].")
  }
  
  eq_names <- base_b[eq_cols]
  
  eq_use <- if (isTRUE(equity_top5)) {
    eq_names[seq_len(min(5, length(eq_names)))]
  } else {
    eq_names
  }
  
  eq_mat <- align_to_index(
    vb[, paste0(eq_use, "_vol"), drop = FALSE],
    idx_ceci,
    object_name = benchmark
  )
  
  eq_avg <- zoo::zoo(
    rowMeans(zoo::coredata(eq_mat), na.rm = TRUE),
    order.by = idx_ceci
  )
  
  com_sig_volpanel <- com_sig
  eq_avg_volpanel <- eq_avg
  
  if (vol_transform == "log") {
    com_sig_volpanel <- log(pmax(com_sig_volpanel, 1e-12))
    eq_avg_volpanel <- log(pmax(eq_avg_volpanel, 1e-12))
  }
  
  joint <- joint_stress(eq_avg, com_sig[, 1], method = joint_stress_method)
  
  if (joint_stress_transform == "log") {
    if (joint_stress_method == "zsum") {
      warning("joint_stress_method='zsum' can be negative; log transform skipped.")
    } else {
      joint <- log1p(pmax(joint, 0))
    }
  }
  
  downside <- NULL
  
  if ("downside" %in% include_panels) {
    stop("Downside panel is disabled in this robust drop-in. Remove 'downside' from include_panels.")
  }
  
  cols <- default_cols(models)
  ltys <- default_ltys(models)
  
  draw_order <- unique(c(
    setdiff(models, c(benchmark, "DCC_scalar_stage", "cDCC_Aielli_stage")),
    benchmark,
    "DCC_scalar_stage",
    "cDCC_Aielli_stage"
  ))
  
  draw_order <- draw_order[draw_order %in% models]
  
  k <- length(include_panels)
  
  op <- par(no.readonly = TRUE)
  on.exit(par(op), add = TRUE)
  
  par(
    mfrow = c(k, 1),
    oma = c(outer_bottom_lines, 0, outer_top_lines, 0),
    cex.axis = cex_axis,
    cex.lab = cex_lab
  )
  
  global_title_drawn <- FALSE
  top_panel_labeled <- FALSE
  
  add_global_title_once <- function() {
    if (!global_title_drawn) {
      mtext(main, side = 3, outer = TRUE, line = 1, cex = cex_main)
      global_title_drawn <<- TRUE
    }
  }
  
  maybe_add_episode_labels <- function() {
    if (
      isTRUE(label_episodes) &&
      !top_panel_labeled &&
      !is.null(episodes_named) &&
      nrow(episodes_named) > 0
    ) {
      label_episodes_top(
        episodes = episodes_named,
        label_x_shift_days = label_x_shift_days,
        label_y_frac_from_top = label_y_frac_from_top,
        cex = episode_label_cex,
        col = episode_label_col,
        x_range = x
      )
      top_panel_labeled <<- TRUE
    }
  }
  
  set_mar <- function(is_bottom, is_top = FALSE) {
    if (is_bottom) {
      par(mar = c(4, left_margin_lines, 0, right_margin_lines) + 0.1)
    } else if (is_top) {
      par(mar = c(0, left_margin_lines, 2, right_margin_lines) + 0.1)
    } else {
      par(mar = c(0, left_margin_lines, 0, right_margin_lines) + 0.1)
    }
  }
  
  add_bottom_axis <- function() {
    if (xaxis_style == "years") {
      axis_years(x, cex_axis = cex_axis)
    } else {
      axis(1, cex.axis = cex_axis)
    }
  }
  
  if ("vol" %in% include_panels) {
    set_mar(is_bottom = FALSE, is_top = TRUE)
    y1 <- as.numeric(zoo::coredata(com_sig_volpanel))
    
    plot(x, y1, type = "n", xlab = "", ylab = ifelse(vol_transform == "log", "log(sigma)", "sigma"), main = NULL, xaxt = "n")
    add_global_title_once()
    if (isTRUE(shade_episodes_flag) && !is.null(episodes_named)) shade_episodes(episodes_named, episode_shade_col, x)
    box()
    lines(x, y1, lwd = lwd)
    maybe_add_episode_labels()
    
    par(new = TRUE)
    y2 <- as.numeric(zoo::coredata(eq_avg_volpanel))
    plot(x, y2, type = "n", axes = FALSE, xlab = "", ylab = "", xaxt = "n")
    lines(x, y2, lty = 2, col = "gray40", lwd = lwd)
    axis(4, cex.axis = cex_axis)
    mtext(ifelse(isTRUE(equity_top5), "Equity avg (top 5)", "Equity avg (all)"), side = 4, line = 3, cex = cex_lab)
    
    legend(
      "topleft",
      legend = c(paste0(com_base, " sigma"), ifelse(isTRUE(equity_top5), "Equity avg (5) sigma", "Equity avg (all) sigma")),
      lty = c(1, 2),
      col = c("black", "gray40"),
      bty = "n",
      cex = cex_leg
    )
  }
  
  if ("joint_stress" %in% include_panels) {
    set_mar(is_bottom = FALSE, is_top = !("vol" %in% include_panels))
    
    plot(x, as.numeric(zoo::coredata(joint)), type = "n", xlab = "", xaxt = "n", ylab = "Joint Stress", main = NULL)
    add_global_title_once()
    if (isTRUE(shade_episodes_flag) && !is.null(episodes_named)) shade_episodes(episodes_named, episode_shade_col, x)
    box()
    lines(x, as.numeric(zoo::coredata(joint)), lwd = lwd)
    maybe_add_episode_labels()
  }
  
  if ("ceci" %in% include_panels) {
    is_top_panel <- !("vol" %in% include_panels) && !("joint_stress" %in% include_panels)
    set_mar(is_bottom = FALSE, is_top = is_top_panel)
    
    yl <- range(zoo::coredata(Z), finite = TRUE)
    
    plot(x, zoo::coredata(Z[, 1]), type = "n", xlab = "", ylab = "BIDIR CECI", xaxt = "n", ylim = yl, main = NULL)
    add_global_title_once()
    if (isTRUE(shade_episodes_flag) && !is.null(episodes_named)) shade_episodes(episodes_named, episode_shade_col, x)
    box()
    
    for (nm in draw_order) {
      if (nm == benchmark) next
      lines(x, zoo::coredata(Z[, nm]), col = cols[nm], lty = ltys[nm], lwd = lwd)
    }
    
    lines(x, zoo::coredata(Z[, benchmark]), col = cols[benchmark], lty = ltys[benchmark], lwd = lwd * 2)
    maybe_add_episode_labels()
  }
  
  if ("dev" %in% include_panels) {
    set_mar(is_bottom = TRUE, is_top = length(include_panels) == 1)
    
    yd <- range(zoo::coredata(D), finite = TRUE)
    mdev <- max(abs(yd))
    
    plot(x, zoo::coredata(D[, 1]), type = "n", xlab = "", ylab = "Delta BIDIR CECI", ylim = c(-mdev, mdev), main = NULL, xaxt = "n")
    add_global_title_once()
    if (isTRUE(shade_episodes_flag) && !is.null(episodes_named)) shade_episodes(episodes_named, episode_shade_col, x)
    box()
    abline(h = 0, lty = 3)
    
    for (nm in setdiff(draw_order, benchmark)) {
      lines(x, zoo::coredata(D[, nm]), col = cols[nm], lty = ltys[nm], lwd = lwd)
    }
    
    maybe_add_episode_labels()
    add_bottom_axis()
    
    par(xpd = NA)
    legend(
      "bottom",
      inset = c(0, legend_inset_y),
      legend = c(paste0(benchmark, " (benchmark)"), setdiff(models, benchmark)),
      col = c(cols[benchmark], cols[setdiff(models, benchmark)]),
      lty = c(ltys[benchmark], ltys[setdiff(models, benchmark)]),
      lwd = c(lwd * 2, rep(lwd, length(setdiff(models, benchmark)))),
      bty = "n",
      cex = cex_leg,
      ncol = legend_ncol
    )
  }
  
  invisible(list(Z = Z, D = D, joint = joint))
}

# ==============================================================================
# 6) RUN MAIN ROLLING / STATIC OUTPUTS
# ==============================================================================

cat("\nEstimating main rolling/static BIDIR CECI outputs...\n")

res <- NULL

save_png(out_png_main, {
  res <- run_all_outputs(
    vol_models = vol_models,
    models = models_used,
    commodities = commodities4,
    regime_model = regime_model,
    mode_static = mode_static,
    mode_rolling = mode_rolling,
    regimes = c("low", "mid", "high"),
    q_low = 0.10,
    q_high = 0.90,
    qmid_lo = 0.45,
    qmid_hi = 0.55,
    nlag = nlag,
    nfore = nfore,
    W = W,
    rolling_commodity_for_universe = rolling_commodity_for_universe,
    ceci_direction = ceci_direction,
    equity_regex = equity_regex,
    diagnostics = DIAG
  )
})

save_png(out_png_compare, {
  plot_rolling_CECI_with_stress(
    rolling_obj = res$rolling,
    vol_models = vol_models,
    commodity_requested = rolling_commodity_for_universe,
    benchmark = benchmark_model,
    transform = mode_rolling,
    window = W,
    legend_ncol = legend_ncol,
    lwd = lwd,
    outer_bottom_lines = 9,
    legend_inset_y = legend_inset_y,
    vol_transform = vol_panel_transform,
    equity_top5 = equity_avg_top5,
    right_margin_lines = right_margin_lines,
    include_panels = include_panels,
    joint_stress_method = joint_stress_method,
    joint_stress_transform = joint_stress_transform,
    cex_axis = cex_axis,
    cex_lab = cex_lab,
    cex_main = cex_main,
    cex_leg = cex_leg,
    left_margin_lines = left_margin_lines,
    xaxis_style = "years",
    shade_episodes_flag = shade_stress_episodes,
    episodes_named = episodes_named,
    episode_shade_col = episode_shade_col,
    label_episodes = label_stress_episodes,
    label_x_shift_days = label_x_shift_days,
    label_y_frac_from_top = label_y_frac_from_top,
    episode_label_cex = episode_label_cex,
    episode_label_col = episode_label_col,
    equity_regex = equity_regex
  )
})

# ==============================================================================
# 7) HEATMAP
# ==============================================================================

get_legend_grob <- function(a_gplot) {
  gt <- ggplotGrob(a_gplot)
  guide_index <- which(sapply(gt$grobs, function(x) x$name) == "guide-box")
  if (length(guide_index) == 0) return(NULL)
  gt$grobs[[guide_index[1]]]
}

equalize_panel_widths <- function(grob_list) {
  panel_cols <- unique(unlist(lapply(grob_list, function(g) {
    g$layout$l[g$layout$name == "panel"]
  })))
  
  max_widths <- do.call(
    grid::unit.pmax,
    lapply(grob_list, function(g) g$widths[panel_cols])
  )
  
  for (i in seq_along(grob_list)) {
    grob_list[[i]]$widths[panel_cols] <- max_widths
  }
  
  grob_list
}

equalize_panel_heights <- function(grob_list) {
  panel_rows <- unique(unlist(lapply(grob_list, function(g) {
    g$layout$t[g$layout$name == "panel"]
  })))
  
  max_heights <- do.call(
    grid::unit.pmax,
    lapply(grob_list, function(g) g$heights[panel_rows])
  )
  
  for (i in seq_along(grob_list)) {
    grob_list[[i]]$heights[panel_rows] <- max_heights
  }
  
  grob_list
}

cat("\n==================================================================\n")
cat("HEATMAP OF BUCKET MEAN BIDIR CECI MINUS OVERALL MEAN BIDIR CECI\n")
cat("mode:", mode_rolling, "| regime_model:", regime_model, "\n")
cat("models:", paste(heatmap_models_used, collapse = ", "), "\n")
cat(
  "joint stress method:",
  heatmap_joint_stress_method,
  "| transform:",
  heatmap_joint_stress_transform,
  "\n"
)
cat("==================================================================\n\n")

heatmap_cells <- list()
cell_counter <- 1

for (commodity_i in commodities4) {
  message("Estimating rolling BIDIR outputs for commodity: ", commodity_i)
  
  res_i <- run_all_outputs(
    vol_models = vol_models,
    models = heatmap_models_used,
    commodities = commodities4,
    regime_model = regime_model,
    mode_static = mode_static,
    mode_rolling = mode_rolling,
    regimes = c("low", "mid", "high"),
    q_low = 0.10,
    q_high = 0.90,
    qmid_lo = 0.45,
    qmid_hi = 0.55,
    nlag = nlag,
    nfore = nfore,
    W = W,
    rolling_commodity_for_universe = commodity_i,
    ceci_direction = ceci_direction,
    equity_regex = equity_regex,
    diagnostics = DIAG
  )
  
  for (model_i in heatmap_models_used) {
    message("  -> Computing joint-stress bucket deltas for model: ", model_i)
    
    ceci_z <- extract_ceci_series(res_i, model_i)
    idx_i <- zoo::index(ceci_z)
    
    joint_z <- build_joint_stress_series(
      vol_models = vol_models,
      model_name = model_i,
      commodity_requested = commodity_i,
      idx_ceci = idx_i,
      equity_top5 = heatmap_equity_top5,
      joint_stress_method = heatmap_joint_stress_method,
      joint_stress_transform = heatmap_joint_stress_transform,
      equity_regex = equity_regex
    )
    
    df_i <- data.frame(
      date = as.Date(idx_i),
      commodity = commodity_i,
      model = model_i,
      joint_stress = as.numeric(zoo::coredata(joint_z)),
      ceci = as.numeric(ceci_z[, "CECI"]),
      stringsAsFactors = FALSE
    ) |>
      dplyr::filter(is.finite(joint_stress), is.finite(ceci))
    
    if (nrow(df_i) == 0) {
      warning("No valid aligned observations for commodity=", commodity_i, ", model=", model_i)
      next
    }
    
    overall_mean_ceci <- mean(df_i$ceci, na.rm = TRUE)
    overall_median_ceci <- median(df_i$ceci, na.rm = TRUE)
    
    summ_i <- df_i |>
      dplyr::mutate(stress_bucket = make_stress_bucket(joint_stress)) |>
      dplyr::group_by(model, commodity, stress_bucket) |>
      dplyr::summarise(
        bucket_mean_ceci = mean(ceci, na.rm = TRUE),
        bucket_median_ceci = median(ceci, na.rm = TRUE),
        overall_mean_ceci = overall_mean_ceci,
        overall_median_ceci = overall_median_ceci,
        delta_bucket_mean_from_overall_mean =
          mean(ceci, na.rm = TRUE) - overall_mean_ceci,
        delta_bucket_median_from_overall_median =
          median(ceci, na.rm = TRUE) - overall_median_ceci,
        obs = dplyr::n(),
        joint_min = min(joint_stress, na.rm = TRUE),
        joint_max = max(joint_stress, na.rm = TRUE),
        .groups = "drop"
      )
    
    heatmap_cells[[cell_counter]] <- summ_i
    cell_counter <- cell_counter + 1
  }
}

if (length(heatmap_cells) == 0) {
  stop("No heatmap cells were produced.")
}

heatmap_df <- dplyr::bind_rows(heatmap_cells) |>
  dplyr::mutate(
    commodity = factor(commodity, levels = commodities4),
    stress_bucket = factor(stress_bucket, levels = c("0-25", "25-50", "50-75", "75-100")),
    model = factor(model, levels = heatmap_models_used)
  ) |>
  tidyr::complete(
    model,
    commodity,
    stress_bucket,
    fill = list(
      bucket_mean_ceci = NA_real_,
      bucket_median_ceci = NA_real_,
      overall_mean_ceci = NA_real_,
      overall_median_ceci = NA_real_,
      delta_bucket_mean_from_overall_mean = NA_real_,
      delta_bucket_median_from_overall_median = NA_real_,
      obs = 0L,
      joint_min = NA_real_,
      joint_max = NA_real_
    )
  ) |>
  dplyr::mutate(
    commodity_label = factor(
      as.character(commodity_short_labels[as.character(commodity)]),
      levels = unname(commodity_short_labels[commodities4])
    ),
    heat_value = delta_bucket_mean_from_overall_mean
  )

write.csv(heatmap_df, out_csv_heatmap, row.names = FALSE)

cat("Heatmap cell summary:\n")
print(
  heatmap_df |>
    dplyr::arrange(model, commodity, stress_bucket) |>
    dplyr::select(
      model,
      commodity,
      stress_bucket,
      bucket_mean_ceci,
      overall_mean_ceci,
      delta_bucket_mean_from_overall_mean,
      obs
    ),
  row.names = FALSE
)

fill_lim <- max(abs(heatmap_df$heat_value), na.rm = TRUE)
if (!is.finite(fill_lim) || fill_lim <= 0) fill_lim <- 1

pretty_title <- c(
  DCC_full          = "DCC_full",
  DCC_scalar_stage  = "DCC_scalar_stage",
  cDCC_Aielli_stage = "cDCC_Aielli_stage",
  sBEKK_sym         = "sBEKK_sym",
  dBEKK_sym         = "dBEKK_sym",
  dBEKK_asym        = "dBEKK_asym"
)

base_heatmap <- function(df_plot,
                         panel_title = "",
                         show_y_title = TRUE,
                         hide_y_text_visually = FALSE,
                         show_legend = TRUE,
                         show_x_title = TRUE,
                         show_x_text = TRUE) {
  p <- ggplot(
    df_plot,
    aes(x = stress_bucket, y = commodity_label, fill = heat_value)
  ) +
    geom_tile(color = "white", linewidth = 0.8) +
    geom_text(
      aes(label = ifelse(is.na(heat_value), "", sprintf("%.2f", heat_value))),
      size = 4.6
    ) +
    scale_fill_gradient2(
      low = "#2166AC",
      mid = "white",
      high = "#B2182B",
      midpoint = 0,
      limits = c(-fill_lim, fill_lim),
      oob = scales::squish,
      name = expression(Delta * " CECI")
    ) +
    labs(
      title = panel_title,
      x = if (show_x_title) "Joint Stress Quantile" else NULL,
      y = if (show_y_title) "Commodity" else NULL
    ) +
    theme_minimal(base_size = 16) +
    theme(
      panel.grid = element_blank(),
      axis.text.x = if (show_x_text) element_text(size = 13) else element_blank(),
      axis.title.x = if (show_x_title) element_text(size = 15) else element_blank(),
      axis.ticks.x = if (show_x_text) element_line() else element_blank(),
      plot.title = element_text(size = 17, face = "bold", hjust = 0.5),
      legend.position = if (show_legend) "right" else "none",
      legend.title = element_text(size = 13),
      legend.text = element_text(size = 11)
    )
  
  if (!hide_y_text_visually) {
    p <- p +
      theme(
        axis.text.y = element_text(size = 13),
        axis.title.y = if (show_y_title) element_text(size = 15) else element_blank(),
        axis.ticks.y = element_line()
      )
  } else {
    p <- p +
      theme(
        axis.text.y = element_text(size = 13, colour = scales::alpha("black", 0)),
        axis.title.y = element_blank(),
        axis.ticks.y = element_line(colour = scales::alpha("black", 0))
      )
  }
  
  p
}

make_panel_grob <- function(model_name,
                            show_y_title,
                            hide_y_text_visually,
                            show_legend = FALSE,
                            show_x_title = TRUE,
                            show_x_text = TRUE) {
  df_plot <- heatmap_df |>
    dplyr::filter(model == model_name)
  
  p <- base_heatmap(
    df_plot = df_plot,
    panel_title = pretty_title[[model_name]],
    show_y_title = show_y_title,
    hide_y_text_visually = hide_y_text_visually,
    show_legend = show_legend,
    show_x_title = show_x_title,
    show_x_text = show_x_text
  )
  
  ggplotGrob(p)
}

g_dcc1 <- make_panel_grob("DCC_full", TRUE, FALSE, FALSE, FALSE, FALSE)
g_dcc2 <- make_panel_grob("DCC_scalar_stage", TRUE, FALSE, FALSE, FALSE, FALSE)
g_dcc3 <- make_panel_grob("cDCC_Aielli_stage", TRUE, FALSE, FALSE, TRUE, TRUE)

p_bekk_for_legend <- base_heatmap(
  df_plot = heatmap_df |> dplyr::filter(model == "dBEKK_asym"),
  panel_title = pretty_title[["dBEKK_asym"]],
  show_y_title = FALSE,
  hide_y_text_visually = TRUE,
  show_legend = TRUE,
  show_x_title = TRUE,
  show_x_text = TRUE
)

legend_grob <- get_legend_grob(p_bekk_for_legend)

g_bekk1 <- make_panel_grob("sBEKK_sym", FALSE, TRUE, FALSE, FALSE, FALSE)
g_bekk2 <- make_panel_grob("dBEKK_sym", FALSE, TRUE, FALSE, FALSE, FALSE)

g_bekk3 <- ggplotGrob(
  base_heatmap(
    df_plot = heatmap_df |> dplyr::filter(model == "dBEKK_asym"),
    panel_title = pretty_title[["dBEKK_asym"]],
    show_y_title = FALSE,
    hide_y_text_visually = TRUE,
    show_legend = FALSE,
    show_x_title = TRUE,
    show_x_text = TRUE
  )
)

all_panels <- list(g_dcc1, g_bekk1, g_dcc2, g_bekk2, g_dcc3, g_bekk3)

all_panels <- equalize_panel_widths(all_panels)
all_panels <- equalize_panel_heights(all_panels)

panel_grid <- gridExtra::arrangeGrob(
  grobs = all_panels,
  ncol = 2,
  widths = c(1, 1)
)

combined_plots <- gridExtra::arrangeGrob(
  grobs = list(panel_grid, legend_grob),
  ncol = 2,
  widths = c(1, 0.10)
)

title_grob <- grid::textGrob(
  "BIDIR CECI: Bucket Mean minus Overall Mean",
  gp = grid::gpar(fontsize = 22, fontface = "bold")
)

subtitle_grob <- grid::textGrob(
  "For each commodity × model, heatmap cells show: mean(BIDIR CECI | joint-stress bucket) - mean(BIDIR CECI over full sample)",
  gp = grid::gpar(fontsize = 14)
)

column_header <- grid::textGrob(
  "Left column: DCC specifications    |    Right column: BEKK specifications",
  gp = grid::gpar(fontsize = 14)
)

final_plot <- gridExtra::arrangeGrob(
  title_grob,
  subtitle_grob,
  column_header,
  combined_plots,
  ncol = 1,
  heights = c(0.05, 0.05, 0.04, 0.86)
)

save_png(
  out_png_heatmap,
  {
    grid::grid.newpage()
    grid::grid.draw(final_plot)
  },
  width = 3200,
  height = 2600,
  res = 220
)

# ==============================================================================
# 8) SUMMARY
# ==============================================================================

cat("\n=============================\n")
cat("BIDIR CECI FIGURES COMPLETE\n")
cat("=============================\n")

cat("Saved outputs:\n")
cat(" - ", out_png_main, "\n", sep = "")
cat(" - ", out_png_compare, "\n", sep = "")
cat(" - ", out_png_heatmap, "\n", sep = "")
cat(" - ", out_csv_heatmap, "\n\n", sep = "")

cat("Core settings:\n")
cat(" - CECI direction :", ceci_direction, "\n")
cat(" - Input mode     :", mode_rolling, "\n")
cat(" - Formula        : BIDIR = 100*sum(theta[E,C]) + 100*sum(theta[C,E])\n")
cat(" - Equity regex   :", equity_regex, "\n\n")

if (shade_stress_episodes) {
  cat("Stress episodes shaded:\n")
  print(episodes_named, row.names = FALSE)
}

















































################################################################################
# ADD-ON — Directional CECI plots using cached rolling objects
#
# This reuses the rolling ConnectednessApproach objects already estimated above.
# It does NOT re-estimate VAR/GFEVD if res_by_commodity[[commodity_requested]]
# exists.
################################################################################

cat("\n==================================================================\n")
cat("DIRECTIONAL CECI ADD-ON USING CACHED ROLLING OBJECTS\n")
cat("==================================================================\n\n")

# ------------------------------ USER SETTINGS --------------------------------

models_directional <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym",
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage"
)

# Choose the commodity for the directional E->C / C->E / BIDIR plots.
commodity_requested <- "API2 (Coal)"

benchmark_model_directional <- "DCC_full"

# Output files
commodity_file_tag <- gsub("[^A-Za-z0-9]+", "_", commodity_requested)
commodity_file_tag <- gsub("_+$", "", commodity_file_tag)

out_png_levels_directional <- file.path(
  output_dir,
  paste0("CECI_directional_levels_models_", commodity_file_tag, ".png")
)

out_png_devs_directional <- file.path(
  output_dir,
  paste0("CECI_directional_devs_vs_benchmark_", commodity_file_tag, ".png")
)

out_csv_stats_directional <- file.path(
  output_dir,
  paste0("CECI_directional_stats_vs_benchmark_", commodity_file_tag, ".csv")
)

# Colors for shaded episode bands
episode_fill_directional <- c("grey90", "grey85", "grey90", "grey85")
episode_border_directional <- NA
episode_label_cex_directional <- 0.95

# Typography / sizing
cex_axis_directional   <- 1.25
cex_lab_directional    <- 1.45
cex_title_directional  <- 1.15
legend_cex_directional <- 1.40
lwd_main_directional   <- 1.35
lwd_bench_mult_directional <- 2.0

legend_ncol_directional    <- 3
legend_inset_y_directional <- -0.36

left_margin_lines_directional  <- 5.2
right_margin_lines_directional <- 1.6
outer_bottom_lines_directional <- 6.5

# -------------------------- CACHE RETRIEVAL -----------------------------------

if (!exists("res_by_commodity")) {
  res_by_commodity <- list()
}

if (
  !is.null(res_by_commodity[[commodity_requested]]) &&
  !is.null(res_by_commodity[[commodity_requested]]$rolling)
) {
  cat("Using cached rolling result for:", commodity_requested, "\n")
  res_directional <- res_by_commodity[[commodity_requested]]
} else if (
  exists("res") &&
  !is.null(res) &&
  !is.null(res$rolling) &&
  identical(commodity_requested, rolling_commodity_for_universe)
) {
  cat("Using main res object for:", commodity_requested, "\n")
  res_directional <- res
  res_by_commodity[[commodity_requested]] <- res
} else {
  cat(
    "No cached rolling result found for ",
    commodity_requested,
    ". Re-estimating only this commodity.\n",
    sep = ""
  )
  
  res_directional <- run_all_outputs(
    vol_models = vol_models,
    models = models_directional,
    commodities = commodities4,
    regime_model = regime_model,
    mode_static = mode_static,
    mode_rolling = mode_rolling,
    regimes = c("low", "mid", "high"),
    q_low = 0.10,
    q_high = 0.90,
    qmid_lo = 0.45,
    qmid_hi = 0.55,
    nlag = nlag,
    nfore = nfore,
    W = W,
    rolling_commodity_for_universe = commodity_requested,
    ceci_direction = "BIDIR",
    equity_regex = equity_regex,
    diagnostics = DIAG
  )
  
  res_by_commodity[[commodity_requested]] <- res_directional
}

# ------------------------------ HELPERS ---------------------------------------

.get_directional_suite_from_cached_roll <- function(res_obj,
                                                    model_name,
                                                    commodity_base,
                                                    equity_regex = "^SX",
                                                    scale_100 = TRUE) {
  if (
    is.null(res_obj$rolling) ||
    is.null(res_obj$rolling$rolling) ||
    is.null(res_obj$rolling$rolling[[model_name]])
  ) {
    stop("Cached rolling ConnectednessApproach object not found for model: ", model_name)
  }
  
  if (
    is.null(res_obj$rolling$X_inputs) ||
    is.null(res_obj$rolling$X_inputs[[model_name]])
  ) {
    stop("Cached X input not found for model: ", model_name)
  }
  
  roll_obj <- res_obj$rolling$rolling[[model_name]]
  X_input  <- res_obj$rolling$X_inputs[[model_name]]
  
  theta <- get_theta_cube_from_dca(roll_obj)
  
  asset_names <- dimnames(theta)[[1]]
  if (is.null(asset_names)) {
    stop("Theta cube has no asset dimnames for model: ", model_name)
  }
  
  base_names <- strip_vol(asset_names)
  
  com_base_hit <- match_base_name(commodity_base, base_names)
  if (is.na(com_base_hit)) {
    stop(
      "Commodity '", commodity_base, "' not found in theta assets for model ",
      model_name, ". Example assets: ",
      paste(head(asset_names, 8), collapse = ", ")
    )
  }
  
  com_asset <- asset_names[match(com_base_hit, base_names)][1]
  com_idx <- which(asset_names == com_asset)
  
  if (length(com_idx) != 1) {
    stop("Could not uniquely locate commodity asset '", com_asset, "' for model ", model_name)
  }
  
  eq_idx <- grep(equity_regex, base_names)
  if (length(eq_idx) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "' for model ", model_name)
  }
  
  TT <- dim(theta)[3]
  e2c <- numeric(TT)
  c2e <- numeric(TT)
  
  for (tt in seq_len(TT)) {
    M <- theta[, , tt]
    e2c[tt] <- sum(M[eq_idx, com_idx], na.rm = TRUE)
    c2e[tt] <- sum(M[com_idx, eq_idx], na.rm = TRUE)
  }
  
  if (isTRUE(scale_100)) {
    e2c <- 100 * e2c
    c2e <- 100 * c2e
  }
  
  idx <- tail(zoo::index(X_input), TT)
  
  list(
    E_to_C = clean_zoo(zoo::zoo(e2c, order.by = idx), paste0("E_to_C_", model_name)),
    C_to_E = clean_zoo(zoo::zoo(c2e, order.by = idx), paste0("C_to_E_", model_name)),
    BIDIR  = clean_zoo(zoo::zoo(e2c + c2e, order.by = idx), paste0("BIDIR_", model_name))
  )
}

.stats_vs_benchmark_cached <- function(x, xb) {
  Z <- merge_zoo_list_inner(
    list(x = x, xb = xb),
    names_out = c("x", "xb"),
    object_name = "directional_stats"
  )
  
  x1 <- as.numeric(Z[, "x"])
  x2 <- as.numeric(Z[, "xb"])
  
  ok <- is.finite(x1) & is.finite(x2)
  x1 <- x1[ok]
  x2 <- x2[ok]
  
  if (length(x1) < 5) {
    return(c(MAD = NA_real_, MaxAD = NA_real_, Corr = NA_real_))
  }
  
  d <- x1 - x2
  
  c(
    MAD = mean(abs(d)),
    MaxAD = max(abs(d)),
    Corr = suppressWarnings(stats::cor(x1, x2))
  )
}

.add_episode_shading_directional <- function(x_dates,
                                             episodes_df,
                                             fills,
                                             border = NA,
                                             label_top = FALSE,
                                             label_cex = 0.95) {
  usr <- par("usr")
  x_min <- as.Date(usr[1], origin = "1970-01-01")
  x_max <- as.Date(usr[2], origin = "1970-01-01")
  
  for (i in seq_len(nrow(episodes_df))) {
    xs <- max(episodes_df$start[i], x_min)
    xe <- min(episodes_df$end[i], x_max)
    
    if (xs > xe) next
    
    rect(
      xs,
      usr[3],
      xe,
      usr[4],
      col = fills[(i - 1) %% length(fills) + 1],
      border = border
    )
    
    if (isTRUE(label_top)) {
      xmid <- xs + floor(as.numeric(xe - xs) / 2)
      ytxt <- usr[4] - 0.05 * (usr[4] - usr[3])
      text(
        xmid,
        ytxt,
        labels = episodes_df$label[i],
        cex = label_cex,
        xpd = NA
      )
    }
  }
}

# ------------------------------ BUILD SERIES ----------------------------------

missing_directional_models <- setdiff(models_directional, names(res_directional$rolling$rolling))
if (length(missing_directional_models) > 0) {
  stop(
    "Cached object does not contain these directional models: ",
    paste(missing_directional_models, collapse = ", ")
  )
}

ceci_dir <- list()

for (m in models_directional) {
  cat("Extracting directional CECI from cached rolling object:", commodity_requested, "|", m, "\n")
  
  ceci_dir[[m]] <- .get_directional_suite_from_cached_roll(
    res_obj = res_directional,
    model_name = m,
    commodity_base = commodity_requested,
    equity_regex = equity_regex,
    scale_100 = TRUE
  )
}

if (!(benchmark_model_directional %in% names(ceci_dir))) {
  stop("benchmark_model_directional not found in computed directional results.")
}

merge_models_directional <- function(direction) {
  lst <- lapply(models_directional, function(m) ceci_dir[[m]][[direction]])
  names(lst) <- models_directional
  
  merge_zoo_list_inner(
    zlist = lst,
    names_out = models_directional,
    object_name = paste0("directional_", direction)
  )
}

Z_e2c <- merge_models_directional("E_to_C")
Z_c2e <- merge_models_directional("C_to_E")
Z_bi  <- merge_models_directional("BIDIR")

# ------------------------------ STATS -----------------------------------------

stats_tbl_directional <- do.call(
  rbind,
  lapply(models_directional, function(m) {
    cbind(
      Model = m,
      Direction = c("E_to_C", "C_to_E", "BIDIR"),
      rbind(
        .stats_vs_benchmark_cached(Z_e2c[, m], Z_e2c[, benchmark_model_directional]),
        .stats_vs_benchmark_cached(Z_c2e[, m], Z_c2e[, benchmark_model_directional]),
        .stats_vs_benchmark_cached(Z_bi[,  m], Z_bi[,  benchmark_model_directional])
      )
    )
  })
)

stats_tbl_directional <- as.data.frame(stats_tbl_directional, stringsAsFactors = FALSE)
stats_tbl_directional$MAD   <- as.numeric(stats_tbl_directional$MAD)
stats_tbl_directional$MaxAD <- as.numeric(stats_tbl_directional$MaxAD)
stats_tbl_directional$Corr  <- as.numeric(stats_tbl_directional$Corr)

utils::write.csv(stats_tbl_directional, out_csv_stats_directional, row.names = FALSE)

# ------------------------------ PLOT STYLE ------------------------------------

cols_directional <- default_cols(models_directional)
ltys_directional <- default_ltys(models_directional)

legend_order_directional <- c(
  benchmark_model_directional,
  setdiff(models_directional, benchmark_model_directional)
)

.set_mar_tight_directional <- function(where = c("top", "mid", "bottom")) {
  where <- match.arg(where)
  
  if (where == "bottom") {
    par(
      mar = c(
        4.0,
        left_margin_lines_directional,
        0.6,
        right_margin_lines_directional
      ) + 0.1
    )
  } else {
    par(
      mar = c(
        0.0,
        left_margin_lines_directional,
        0.6,
        right_margin_lines_directional
      ) + 0.1
    )
  }
}

.plot_levels_panel_directional <- function(Z,
                                           ylab,
                                           xaxt = "n",
                                           shade_labels = FALSE) {
  x <- zoo::index(Z)
  yl <- range(zoo::coredata(Z), finite = TRUE)
  
  plot(
    x,
    zoo::coredata(Z[, 1]),
    type = "n",
    ylim = yl,
    xlab = "",
    ylab = ylab,
    main = NULL,
    xaxt = "n",
    cex.axis = cex_axis_directional,
    cex.lab = cex_lab_directional
  )
  
  .add_episode_shading_directional(
    x_dates = x,
    episodes_df = episodes_named,
    fills = episode_fill_directional,
    border = episode_border_directional,
    label_top = shade_labels,
    label_cex = episode_label_cex_directional
  )
  
  for (m in legend_order_directional) {
    lines(
      x,
      zoo::coredata(Z[, m]),
      col = cols_directional[m],
      lty = ltys_directional[m],
      lwd = if (m == benchmark_model_directional) {
        lwd_main_directional * lwd_bench_mult_directional
      } else {
        lwd_main_directional
      }
    )
  }
  
  box()
  
  if (xaxt == "s") {
    axis_years(x, cex_axis = cex_axis_directional)
  }
}

.plot_dev_panel_directional <- function(Z,
                                        ylab,
                                        xaxt = "n",
                                        shade_labels = FALSE) {
  x <- zoo::index(Z)
  
  D <- Z
  for (m in models_directional) {
    D[, m] <- Z[, m] - Z[, benchmark_model_directional]
  }
  
  yd <- range(zoo::coredata(D), finite = TRUE)
  mdev <- max(abs(yd), na.rm = TRUE)
  
  plot(
    x,
    zoo::coredata(D[, 1]),
    type = "n",
    ylim = c(-mdev, mdev),
    xlab = "",
    ylab = ylab,
    main = NULL,
    xaxt = "n",
    cex.axis = cex_axis_directional,
    cex.lab = cex_lab_directional
  )
  
  .add_episode_shading_directional(
    x_dates = x,
    episodes_df = episodes_named,
    fills = episode_fill_directional,
    border = episode_border_directional,
    label_top = shade_labels,
    label_cex = episode_label_cex_directional
  )
  
  abline(h = 0, lty = 3)
  
  for (m in setdiff(legend_order_directional, benchmark_model_directional)) {
    lines(
      x,
      zoo::coredata(D[, m]),
      col = cols_directional[m],
      lty = ltys_directional[m],
      lwd = lwd_main_directional
    )
  }
  
  box()
  
  if (xaxt == "s") {
    axis_years(x, cex_axis = cex_axis_directional)
  }
}

# ------------------------------ PLOTS -----------------------------------------

save_png(
  out_png_levels_directional,
  {
    op <- par(no.readonly = TRUE)
    on.exit(par(op), add = TRUE)
    
    par(
      mfrow = c(3, 1),
      oma = c(outer_bottom_lines_directional, 0, 3.0, 0)
    )
    
    .set_mar_tight_directional("top")
    .plot_levels_panel_directional(
      Z_e2c,
      "CECI (E\u2192C)",
      xaxt = "n",
      shade_labels = TRUE
    )
    
    .set_mar_tight_directional("mid")
    .plot_levels_panel_directional(
      Z_c2e,
      "CECI (C\u2192E)",
      xaxt = "n",
      shade_labels = FALSE
    )
    
    .set_mar_tight_directional("bottom")
    .plot_levels_panel_directional(
      Z_bi,
      "CECI (BIDIR)",
      xaxt = "s",
      shade_labels = FALSE
    )
    
    mtext(
      paste0(
        "Directional CECI (",
        commodity_requested,
        ") — input: ",
        mode_rolling,
        ", rolling W=",
        W,
        ", nlag=",
        nlag,
        ", H=",
        nfore,
        " — benchmark: ",
        benchmark_model_directional
      ),
      outer = TRUE,
      side = 3,
      line = 1,
      cex = cex_title_directional
    )
    
    par(xpd = NA)
    
    legend(
      "bottom",
      inset = c(0, legend_inset_y_directional),
      legend = c(
        paste0(benchmark_model_directional, " (benchmark)"),
        setdiff(legend_order_directional, benchmark_model_directional)
      ),
      col = c(
        cols_directional[benchmark_model_directional],
        cols_directional[setdiff(legend_order_directional, benchmark_model_directional)]
      ),
      lty = c(
        ltys_directional[benchmark_model_directional],
        ltys_directional[setdiff(legend_order_directional, benchmark_model_directional)]
      ),
      lwd = c(
        lwd_main_directional * lwd_bench_mult_directional,
        rep(
          lwd_main_directional,
          length(setdiff(legend_order_directional, benchmark_model_directional))
        )
      ),
      bty = "n",
      cex = legend_cex_directional,
      ncol = legend_ncol_directional
    )
  },
  width = 2800,
  height = 2200,
  res = 220
)

save_png(
  out_png_devs_directional,
  {
    op <- par(no.readonly = TRUE)
    on.exit(par(op), add = TRUE)
    
    par(
      mfrow = c(3, 1),
      oma = c(outer_bottom_lines_directional, 0, 3.0, 0)
    )
    
    .set_mar_tight_directional("top")
    .plot_dev_panel_directional(
      Z_e2c,
      "\u0394CECI (E\u2192C)",
      xaxt = "n",
      shade_labels = TRUE
    )
    
    .set_mar_tight_directional("mid")
    .plot_dev_panel_directional(
      Z_c2e,
      "\u0394CECI (C\u2192E)",
      xaxt = "n",
      shade_labels = FALSE
    )
    
    .set_mar_tight_directional("bottom")
    .plot_dev_panel_directional(
      Z_bi,
      "\u0394CECI (BIDIR)",
      xaxt = "s",
      shade_labels = FALSE
    )
    
    mtext(
      paste0(
        "Directional CECI deviations vs benchmark (",
        benchmark_model_directional,
        ") — ",
        commodity_requested
      ),
      outer = TRUE,
      side = 3,
      line = 1,
      cex = cex_title_directional
    )
    
    par(xpd = NA)
    
    legend(
      "bottom",
      inset = c(0, legend_inset_y_directional),
      legend = setdiff(legend_order_directional, benchmark_model_directional),
      col = cols_directional[setdiff(legend_order_directional, benchmark_model_directional)],
      lty = ltys_directional[setdiff(legend_order_directional, benchmark_model_directional)],
      lwd = lwd_main_directional,
      bty = "n",
      cex = legend_cex_directional,
      ncol = legend_ncol_directional
    )
  },
  width = 2800,
  height = 2200,
  res = 220
)

cat("\nDirectional CECI add-on complete.\n")
cat("Levels PNG   :", out_png_levels_directional, "\n")
cat("Devs PNG     :", out_png_devs_directional, "\n")
cat("Stats CSV    :", out_csv_stats_directional, "\n")











################################################################################
# ADD-ON — 4-panel directional CECI imbalance using cached rolling objects
#
# Reuses:
#   res_by_commodity[[commodity]]$rolling$rolling[[model]]
#   res_by_commodity[[commodity]]$rolling$X_inputs[[model]]
#
# Plots:
#   Imbalance_t = CECI(E->C)_t - CECI(C->E)_t
#
# No re-estimation if res_by_commodity contains all commodities.
################################################################################

cat("\n==================================================================\n")
cat("DIRECTIONAL CECI IMBALANCE ADD-ON USING CACHED ROLLING OBJECTS\n")
cat("==================================================================\n\n")

# ------------------------------ USER SETTINGS --------------------------------

models_imbalance <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym",
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage"
)

commodities_imbalance <- c(
  "TTF (Gas)",
  "Brent (Oil)",
  "API2 (Coal)",
  "MO1 (Carbon)"
)

benchmark_model_imbalance <- "DCC_full"

# Uses cached rolling settings. Do not set this to 150 unless the cache was built with W=150.
W_imbalance_label <- W

normalize_imbalance <- TRUE
normalization_method <- "zscore"
normalization_eps <- 1e-12

series_label_imbalance <- if (normalize_imbalance) "zscore" else "raw"

out_png_imbalance_cached <- file.path(
  output_dir,
  paste0(
    "CECI_directional_imbalance_EtoC_minus_CtoE_4commodities_",
    series_label_imbalance,
    "_cached.png"
  )
)

out_csv_imbalance_stats_cached <- file.path(
  output_dir,
  paste0(
    "CECI_directional_imbalance_stats_vs_benchmark_",
    series_label_imbalance,
    "_cached.csv"
  )
)

episode_fill_imbalance <- c("grey90", "grey85", "grey90", "grey85")
episode_border_imbalance <- NA
episode_label_cex_imbalance <- 0.92

cex_axis_imbalance   <- 1.20
cex_lab_imbalance    <- 1.40
cex_title_imbalance  <- 1.15
legend_cex_imbalance <- 1.35
lwd_main_imbalance   <- 1.35
lwd_bench_mult_imbalance <- 2.0

legend_ncol_imbalance    <- 3
legend_inset_y_imbalance <- -0.34

left_margin_lines_imbalance  <- 5.5
right_margin_lines_imbalance <- 1.6
outer_bottom_lines_imbalance <- 6.8

# ------------------------------ CACHE CHECK -----------------------------------

if (!exists("res_by_commodity")) {
  res_by_commodity <- list()
}

# Cache the main result if available and not already stored.
if (
  exists("res") &&
  !is.null(res) &&
  !is.null(res$rolling) &&
  exists("rolling_commodity_for_universe") &&
  is.null(res_by_commodity[[rolling_commodity_for_universe]])
) {
  res_by_commodity[[rolling_commodity_for_universe]] <- res
}

missing_cached_commodities <- commodities_imbalance[
  !vapply(
    commodities_imbalance,
    function(com) {
      !is.null(res_by_commodity[[com]]) &&
        !is.null(res_by_commodity[[com]]$rolling) &&
        !is.null(res_by_commodity[[com]]$rolling$rolling) &&
        !is.null(res_by_commodity[[com]]$rolling$X_inputs)
    },
    logical(1)
  )
]

if (length(missing_cached_commodities) > 0) {
  cat(
    "These commodities are missing from cache and will be estimated once: ",
    paste(missing_cached_commodities, collapse = ", "),
    "\n",
    sep = ""
  )
  
  for (commodity_i in missing_cached_commodities) {
    res_i <- run_all_outputs(
      vol_models = vol_models,
      models = models_imbalance,
      commodities = commodities4,
      regime_model = regime_model,
      mode_static = mode_static,
      mode_rolling = mode_rolling,
      regimes = c("low", "mid", "high"),
      q_low = 0.10,
      q_high = 0.90,
      qmid_lo = 0.45,
      qmid_hi = 0.55,
      nlag = nlag,
      nfore = nfore,
      W = W,
      rolling_commodity_for_universe = commodity_i,
      ceci_direction = "BIDIR",
      equity_regex = equity_regex,
      diagnostics = DIAG
    )
    
    res_by_commodity[[commodity_i]] <- res_i
  }
} else {
  cat("All imbalance commodities found in cache. No rolling VAR/GFEVD re-estimation needed.\n")
}

# ------------------------------ HELPERS ---------------------------------------

.extract_directional_suite_cached <- function(res_obj,
                                              model_name,
                                              commodity_base,
                                              equity_regex = "^SX",
                                              scale_100 = TRUE) {
  if (
    is.null(res_obj$rolling) ||
    is.null(res_obj$rolling$rolling) ||
    is.null(res_obj$rolling$rolling[[model_name]])
  ) {
    stop("Cached rolling ConnectednessApproach object not found for model: ", model_name)
  }
  
  if (
    is.null(res_obj$rolling$X_inputs) ||
    is.null(res_obj$rolling$X_inputs[[model_name]])
  ) {
    stop("Cached X input not found for model: ", model_name)
  }
  
  roll_obj <- res_obj$rolling$rolling[[model_name]]
  X_input  <- res_obj$rolling$X_inputs[[model_name]]
  
  theta <- get_theta_cube_from_dca(roll_obj)
  
  asset_names <- dimnames(theta)[[1]]
  if (is.null(asset_names)) {
    stop("Theta cube has no asset dimnames for model: ", model_name)
  }
  
  base_names <- strip_vol(asset_names)
  
  com_base_hit <- match_base_name(commodity_base, base_names)
  if (is.na(com_base_hit)) {
    stop(
      "Commodity '", commodity_base, "' not found in theta assets for model ",
      model_name, ". Example assets: ",
      paste(head(asset_names, 8), collapse = ", ")
    )
  }
  
  com_asset <- asset_names[match(com_base_hit, base_names)][1]
  com_idx <- which(asset_names == com_asset)
  
  if (length(com_idx) != 1) {
    stop("Could not uniquely locate commodity asset '", com_asset, "' for model ", model_name)
  }
  
  eq_idx <- grep(equity_regex, base_names)
  if (length(eq_idx) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "' for model ", model_name)
  }
  
  TT <- dim(theta)[3]
  e2c <- numeric(TT)
  c2e <- numeric(TT)
  
  for (tt in seq_len(TT)) {
    M <- theta[, , tt]
    e2c[tt] <- sum(M[eq_idx, com_idx], na.rm = TRUE)
    c2e[tt] <- sum(M[com_idx, eq_idx], na.rm = TRUE)
  }
  
  if (isTRUE(scale_100)) {
    e2c <- 100 * e2c
    c2e <- 100 * c2e
  }
  
  idx <- tail(zoo::index(X_input), TT)
  
  list(
    E_to_C = clean_zoo(zoo::zoo(e2c, order.by = idx), paste0("E_to_C_", model_name)),
    C_to_E = clean_zoo(zoo::zoo(c2e, order.by = idx), paste0("C_to_E_", model_name)),
    BIDIR  = clean_zoo(zoo::zoo(e2c + c2e, order.by = idx), paste0("BIDIR_", model_name)),
    IMBAL  = clean_zoo(zoo::zoo(e2c - c2e, order.by = idx), paste0("IMBAL_", model_name))
  )
}

.normalize_zoo_matrix <- function(Z,
                                  method = c("zscore"),
                                  eps = 1e-12) {
  method <- match.arg(method)
  
  Z <- clean_zoo(Z, "imbalance_normalization_input")
  
  X <- zoo::coredata(Z)
  if (is.null(dim(X))) {
    X <- matrix(X, ncol = 1)
  }
  
  out <- matrix(NA_real_, nrow = nrow(X), ncol = ncol(X))
  colnames(out) <- colnames(X)
  
  for (j in seq_len(ncol(X))) {
    xj <- X[, j]
    ok <- is.finite(xj)
    
    if (sum(ok) < 2) {
      out[, j] <- NA_real_
      next
    }
    
    mu <- mean(xj[ok], na.rm = TRUE)
    sdv <- stats::sd(xj[ok], na.rm = TRUE)
    
    if (!is.finite(sdv) || sdv < eps) {
      out[, j] <- 0
    } else {
      out[, j] <- (xj - mu) / sdv
    }
  }
  
  zoo::zoo(out, order.by = zoo::index(Z))
}

.stats_vs_benchmark_imbalance <- function(x, xb) {
  Z <- merge_zoo_list_inner(
    list(x = x, xb = xb),
    names_out = c("x", "xb"),
    object_name = "imbalance_stats"
  )
  
  x1 <- as.numeric(Z[, "x"])
  x2 <- as.numeric(Z[, "xb"])
  
  ok <- is.finite(x1) & is.finite(x2)
  x1 <- x1[ok]
  x2 <- x2[ok]
  
  if (length(x1) < 5) {
    return(c(MAD = NA_real_, MaxAD = NA_real_, Corr = NA_real_))
  }
  
  d <- x1 - x2
  
  c(
    MAD = mean(abs(d)),
    MaxAD = max(abs(d)),
    Corr = suppressWarnings(stats::cor(x1, x2))
  )
}

.pretty_commodity_name_imbalance <- function(x) {
  map <- c(
    "TTF (Gas)"    = "Gas",
    "Brent (Oil)"  = "Oil",
    "API2 (Coal)"  = "Coal",
    "MO1 (Carbon)" = "Carbon"
  )
  
  if (x %in% names(map)) {
    map[[x]]
  } else {
    x
  }
}

.plot_ylabel_expr_imbalance <- function(normalized = FALSE) {
  if (normalized) {
    expression(paste("Normalized ", CECI[E %->% C] - CECI[C %->% E], " (z-score)"))
  } else {
    expression(CECI[E %->% C] - CECI[C %->% E])
  }
}

.add_episode_shading_imbalance <- function(x_dates,
                                           episodes_df,
                                           fills,
                                           border = NA,
                                           label_top = FALSE,
                                           label_cex = 0.95) {
  usr <- par("usr")
  x_min <- as.Date(usr[1], origin = "1970-01-01")
  x_max <- as.Date(usr[2], origin = "1970-01-01")
  
  for (i in seq_len(nrow(episodes_df))) {
    xs <- max(episodes_df$start[i], x_min)
    xe <- min(episodes_df$end[i], x_max)
    
    if (xs > xe) next
    
    rect(
      xs,
      usr[3],
      xe,
      usr[4],
      col = fills[(i - 1) %% length(fills) + 1],
      border = border
    )
    
    if (isTRUE(label_top)) {
      xmid <- xs + floor(as.numeric(xe - xs) / 2)
      ytxt <- usr[4] - 0.05 * (usr[4] - usr[3])
      
      text(
        xmid,
        ytxt,
        labels = episodes_df$label[i],
        cex = label_cex,
        xpd = NA
      )
    }
  }
}

# ------------------------------ BUILD IMBALANCE SERIES -------------------------

ceci_imbal_cached <- list()

for (commodity_i in commodities_imbalance) {
  cat("Extracting cached imbalance series for commodity:", commodity_i, "\n")
  
  res_i <- res_by_commodity[[commodity_i]]
  ceci_imbal_cached[[commodity_i]] <- list()
  
  missing_models_i <- setdiff(models_imbalance, names(res_i$rolling$rolling))
  
  if (length(missing_models_i) > 0) {
    stop(
      "Cached object for ", commodity_i,
      " is missing models: ",
      paste(missing_models_i, collapse = ", ")
    )
  }
  
  for (model_i in models_imbalance) {
    suite_i <- .extract_directional_suite_cached(
      res_obj = res_i,
      model_name = model_i,
      commodity_base = commodity_i,
      equity_regex = equity_regex,
      scale_100 = TRUE
    )
    
    ceci_imbal_cached[[commodity_i]][[model_i]] <- suite_i$IMBAL
  }
}

merge_models_for_commodity_imbalance <- function(commodity_i) {
  lst <- lapply(
    models_imbalance,
    function(model_i) ceci_imbal_cached[[commodity_i]][[model_i]]
  )
  
  names(lst) <- models_imbalance
  
  merge_zoo_list_inner(
    zlist = lst,
    names_out = models_imbalance,
    object_name = paste0("imbalance_", commodity_i)
  )
}

Z_imbalance_raw_list <- lapply(
  commodities_imbalance,
  merge_models_for_commodity_imbalance
)

names(Z_imbalance_raw_list) <- commodities_imbalance

if (isTRUE(normalize_imbalance)) {
  if (normalization_method != "zscore") {
    stop("Unsupported normalization_method: ", normalization_method)
  }
  
  Z_imbalance_list <- lapply(
    Z_imbalance_raw_list,
    .normalize_zoo_matrix,
    method = normalization_method,
    eps = normalization_eps
  )
} else {
  Z_imbalance_list <- Z_imbalance_raw_list
}

# ------------------------------ STATS -----------------------------------------

stats_tbl_imbalance <- do.call(
  rbind,
  lapply(commodities_imbalance, function(com_i) {
    Z <- Z_imbalance_list[[com_i]]
    
    do.call(
      rbind,
      lapply(models_imbalance, function(model_i) {
        st <- .stats_vs_benchmark_imbalance(
          Z[, model_i],
          Z[, benchmark_model_imbalance]
        )
        
        data.frame(
          Commodity = com_i,
          Model = model_i,
          SeriesType = if (normalize_imbalance) "zscore" else "raw",
          MAD = unname(st["MAD"]),
          MaxAD = unname(st["MaxAD"]),
          Corr = unname(st["Corr"]),
          stringsAsFactors = FALSE
        )
      })
    )
  })
)

utils::write.csv(
  stats_tbl_imbalance,
  out_csv_imbalance_stats_cached,
  row.names = FALSE
)

# ------------------------------ PLOT STYLE ------------------------------------

cols_imbalance <- default_cols(models_imbalance)
ltys_imbalance <- default_ltys(models_imbalance)

legend_order_imbalance <- c(
  benchmark_model_imbalance,
  setdiff(models_imbalance, benchmark_model_imbalance)
)

.set_mar_tight_imbalance <- function(where = c("top", "mid", "bottom")) {
  where <- match.arg(where)
  
  if (where == "bottom") {
    par(
      mar = c(
        4.0,
        left_margin_lines_imbalance,
        0.7,
        right_margin_lines_imbalance
      ) + 0.1
    )
  } else {
    par(
      mar = c(
        0.1,
        left_margin_lines_imbalance,
        0.7,
        right_margin_lines_imbalance
      ) + 0.1
    )
  }
}

.plot_imbalance_panel_cached <- function(Z,
                                         ylab="z-score",
                                         main = NULL,
                                         xaxt = "n",
                                         shade_labels = FALSE) {
  x <- zoo::index(Z)
  yl <- range(zoo::coredata(Z), finite = TRUE)
  
  if (all(!is.finite(yl))) {
    yl <- c(-1, 1)
  } else {
    mabs <- max(abs(yl), na.rm = TRUE)
    
    if (!is.finite(mabs) || mabs <= 0) {
      mabs <- 1
    }
    
    yl <- c(-mabs, mabs)
  }
  
  plot(
    x,
    zoo::coredata(Z[, 1]),
    type = "n",
    ylim = yl,
    xlab = "",
    ylab = ylab,
    main = main,
    xaxt = "n",
    cex.axis = cex_axis_imbalance,
    cex.lab = cex_lab_imbalance
  )
  
  .add_episode_shading_imbalance(
    x_dates = x,
    episodes_df = episodes_named,
    fills = episode_fill_imbalance,
    border = episode_border_imbalance,
    label_top = shade_labels,
    label_cex = episode_label_cex_imbalance
  )
  
  abline(h = 0, lty = 3)
  
  for (model_i in legend_order_imbalance) {
    lines(
      x,
      zoo::coredata(Z[, model_i]),
      col = cols_imbalance[model_i],
      lty = ltys_imbalance[model_i],
      lwd = if (model_i == benchmark_model_imbalance) {
        lwd_main_imbalance * lwd_bench_mult_imbalance
      } else {
        lwd_main_imbalance
      }
    )
  }
  
  box()
  
  if (xaxt == "s") {
    axis_years(x, cex_axis = cex_axis_imbalance)
  }
}

# ------------------------------ PLOT ------------------------------------------

save_png(
  out_png_imbalance_cached,
  {
    op <- par(no.readonly = TRUE)
    on.exit(par(op), add = TRUE)
    
    par(
      mfrow = c(4, 1),
      oma = c(outer_bottom_lines_imbalance, 0, 3.0, 0)
    )
    
    for (i in seq_along(commodities_imbalance)) {
      com_i <- commodities_imbalance[i]
      Z_i <- Z_imbalance_list[[com_i]]
      
      if (i == 1) {
        .set_mar_tight_imbalance("top")
        .plot_imbalance_panel_cached(
          Z_i,
          ylab = "z-score",
          main = .pretty_commodity_name_imbalance(com_i),
          xaxt = "n",
          shade_labels = TRUE
        )
      } else if (i == length(commodities_imbalance)) {
        .set_mar_tight_imbalance("bottom")
        .plot_imbalance_panel_cached(
          Z_i,
          ylab = "z-score",
          main = .pretty_commodity_name_imbalance(com_i),
          xaxt = "s",
          shade_labels = FALSE
        )
      } else {
        .set_mar_tight_imbalance("mid")
        .plot_imbalance_panel_cached(
          Z_i,
          ylab = "z-score",
          main = .pretty_commodity_name_imbalance(com_i),
          xaxt = "n",
          shade_labels = FALSE
        )
      }
    }
    
    mtext(
      paste0(
        "Directional CECI imbalance by commodity",
        if (normalize_imbalance) " (z-score normalized)" else "",
        " — input: ",
        mode_rolling,
        ", rolling W=",
        W_imbalance_label,
        ", nlag=",
        nlag,
        ", H=",
        nfore,
        " — benchmark: ",
        benchmark_model_imbalance
      ),
      outer = TRUE,
      side = 3,
      line = 1,
      cex = cex_title_imbalance
    )
    
    par(xpd = NA)
    
    legend(
      "bottom",
      inset = c(0, legend_inset_y_imbalance),
      legend = c(
        paste0(benchmark_model_imbalance, " (benchmark)"),
        setdiff(legend_order_imbalance, benchmark_model_imbalance)
      ),
      col = c(
        cols_imbalance[benchmark_model_imbalance],
        cols_imbalance[setdiff(legend_order_imbalance, benchmark_model_imbalance)]
      ),
      lty = c(
        ltys_imbalance[benchmark_model_imbalance],
        ltys_imbalance[setdiff(legend_order_imbalance, benchmark_model_imbalance)]
      ),
      lwd = c(
        lwd_main_imbalance * lwd_bench_mult_imbalance,
        rep(
          lwd_main_imbalance,
          length(setdiff(legend_order_imbalance, benchmark_model_imbalance))
        )
      ),
      bty = "n",
      cex = legend_cex_imbalance,
      ncol = legend_ncol_imbalance
    )
  },
  width = 2800,
  height = 2600,
  res = 220
)

cat("\nDirectional imbalance add-on complete.\n")
cat("Series type   :", if (normalize_imbalance) "zscore" else "raw", "\n")
cat("Imbalance PNG :", out_png_imbalance_cached, "\n")
cat("Stats CSV     :", out_csv_imbalance_stats_cached, "\n")


























################################################################################
# ADD-ON — 6-panel heatmaps of DIRECTIONAL Delta CECI by joint stress
#
# Reuses cached rolling ConnectednessApproach objects from:
#   res_by_commodity[[commodity]]$rolling$rolling[[model]]
#   res_by_commodity[[commodity]]$rolling$X_inputs[[model]]
#
# Heatmap cell:
#   mean(Delta CECI | joint-stress bucket)
#
# where:
#   Delta CECI_t = CECI(E->C)_t - CECI(C->E)_t
#
# Joint stress is computed from the already cached transformed connectedness
# input X_inputs[[model]], so no rolling VAR/GFEVD is re-estimated if cache exists.
################################################################################

cat("\n==================================================================\n")
cat("DIRECTIONAL DELTA CECI HEATMAP ADD-ON USING CACHED ROLLING OBJECTS\n")
cat("==================================================================\n\n")

# ------------------------------ USER SETTINGS --------------------------------

models_delta_heatmap <- c(
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage",
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym"
)

commodities_delta_heatmap <- c(
  "TTF (Gas)",
  "Brent (Oil)",
  "API2 (Coal)",
  "MO1 (Carbon)"
)

commodity_short_labels_delta <- c(
  "TTF (Gas)"    = "Gas",
  "Brent (Oil)"  = "Oil",
  "API2 (Coal)"  = "Coal",
  "MO1 (Carbon)" = "Carbon"
)

stress_measure_delta <- "mean_z"      # "mean_z" or "sum_z"
normalize_delta_ceci_heatmap <- FALSE
normalization_eps_delta <- 1e-12

quantile_labels_delta <- c("0-25", "25-50", "50-75", "75-100")

series_suffix_delta <- paste0(
  if (normalize_delta_ceci_heatmap) "zscore_" else "raw_",
  "w", W
)

out_png_delta_heatmap_cached <- file.path(
  output_dir,
  paste0(
    "Heatmap_Directional_DeltaCECI_by_JointStressQuantile_6Panel_",
    series_suffix_delta,
    "_cached.png"
  )
)

out_csv_delta_heatmap_cached <- file.path(
  output_dir,
  paste0(
    "Heatmap_Directional_DeltaCECI_by_JointStressQuantile_6Panel_",
    series_suffix_delta,
    "_cached.csv"
  )
)

png_width_delta  <- 3200
png_height_delta <- 2600
png_res_delta    <- 220

# ------------------------------ CACHE CHECK -----------------------------------

if (!exists("res_by_commodity")) {
  res_by_commodity <- list()
}

if (
  exists("res") &&
  !is.null(res) &&
  !is.null(res$rolling) &&
  exists("rolling_commodity_for_universe") &&
  is.null(res_by_commodity[[rolling_commodity_for_universe]])
) {
  res_by_commodity[[rolling_commodity_for_universe]] <- res
}

missing_cached_delta_commodities <- commodities_delta_heatmap[
  !vapply(
    commodities_delta_heatmap,
    function(com) {
      !is.null(res_by_commodity[[com]]) &&
        !is.null(res_by_commodity[[com]]$rolling) &&
        !is.null(res_by_commodity[[com]]$rolling$rolling) &&
        !is.null(res_by_commodity[[com]]$rolling$X_inputs)
    },
    logical(1)
  )
]

if (length(missing_cached_delta_commodities) > 0) {
  cat(
    "These commodities are missing from cache and will be estimated once: ",
    paste(missing_cached_delta_commodities, collapse = ", "),
    "\n",
    sep = ""
  )
  
  for (commodity_i in missing_cached_delta_commodities) {
    res_i <- run_all_outputs(
      vol_models = vol_models,
      models = models_delta_heatmap,
      commodities = commodities4,
      regime_model = regime_model,
      mode_static = mode_static,
      mode_rolling = mode_rolling,
      regimes = c("low", "mid", "high"),
      q_low = 0.10,
      q_high = 0.90,
      qmid_lo = 0.45,
      qmid_hi = 0.55,
      nlag = nlag,
      nfore = nfore,
      W = W,
      rolling_commodity_for_universe = commodity_i,
      ceci_direction = "BIDIR",
      equity_regex = equity_regex,
      diagnostics = DIAG
    )
    
    res_by_commodity[[commodity_i]] <- res_i
  }
} else {
  cat("All directional heatmap commodities found in cache. No rolling VAR/GFEVD re-estimation needed.\n")
}

# ------------------------------ HELPERS ---------------------------------------

.extract_delta_ceci_from_cached_roll <- function(res_obj,
                                                 model_name,
                                                 commodity_base,
                                                 equity_regex = "^SX",
                                                 scale_100 = TRUE) {
  if (
    is.null(res_obj$rolling) ||
    is.null(res_obj$rolling$rolling) ||
    is.null(res_obj$rolling$rolling[[model_name]])
  ) {
    stop("Cached rolling ConnectednessApproach object not found for model: ", model_name)
  }
  
  if (
    is.null(res_obj$rolling$X_inputs) ||
    is.null(res_obj$rolling$X_inputs[[model_name]])
  ) {
    stop("Cached X input not found for model: ", model_name)
  }
  
  roll_obj <- res_obj$rolling$rolling[[model_name]]
  X_input  <- res_obj$rolling$X_inputs[[model_name]]
  
  theta <- get_theta_cube_from_dca(roll_obj)
  
  asset_names <- dimnames(theta)[[1]]
  if (is.null(asset_names)) {
    stop("Theta cube has no asset dimnames for model: ", model_name)
  }
  
  base_names <- strip_vol(asset_names)
  
  com_base_hit <- match_base_name(commodity_base, base_names)
  if (is.na(com_base_hit)) {
    stop(
      "Commodity '", commodity_base, "' not found in theta assets for model ",
      model_name, ". Example assets: ",
      paste(head(asset_names, 8), collapse = ", ")
    )
  }
  
  com_asset <- asset_names[match(com_base_hit, base_names)][1]
  com_idx <- which(asset_names == com_asset)
  
  if (length(com_idx) != 1) {
    stop("Could not uniquely locate commodity asset '", com_asset, "' for model ", model_name)
  }
  
  eq_idx <- grep(equity_regex, base_names)
  if (length(eq_idx) == 0) {
    stop("No equities matched equity_regex='", equity_regex, "' for model ", model_name)
  }
  
  TT <- dim(theta)[3]
  e2c <- numeric(TT)
  c2e <- numeric(TT)
  
  for (tt in seq_len(TT)) {
    M <- theta[, , tt]
    e2c[tt] <- sum(M[eq_idx, com_idx], na.rm = TRUE)
    c2e[tt] <- sum(M[com_idx, eq_idx], na.rm = TRUE)
  }
  
  if (isTRUE(scale_100)) {
    e2c <- 100 * e2c
    c2e <- 100 * c2e
  }
  
  idx <- tail(zoo::index(X_input), TT)
  
  clean_zoo(
    zoo::zoo(e2c - c2e, order.by = idx),
    paste0("DeltaCECI_", model_name, "_", commodity_base)
  )
}

.compute_joint_stress_cached <- function(X_input,
                                         T_out,
                                         method = c("mean_z", "sum_z")) {
  method <- match.arg(method)
  
  X_input <- clean_zoo(X_input, "joint_stress_cached_X_input")
  
  X <- zoo::coredata(X_input)
  if (!is.matrix(X)) {
    X <- matrix(X, ncol = 1)
  }
  
  Xz <- apply(X, 2, function(col) {
    ok <- is.finite(col)
    out <- rep(NA_real_, length(col))
    
    if (sum(ok) < 2) {
      return(out)
    }
    
    mu <- mean(col[ok], na.rm = TRUE)
    sdv <- stats::sd(col[ok], na.rm = TRUE)
    
    if (!is.finite(sdv) || sdv < 1e-12) {
      out[ok] <- 0
    } else {
      out[ok] <- (col[ok] - mu) / sdv
    }
    
    out
  })
  
  if (!is.matrix(Xz)) {
    Xz <- matrix(Xz, ncol = 1)
  }
  
  score <- switch(
    method,
    mean_z = rowMeans(Xz, na.rm = TRUE),
    sum_z  = rowSums(Xz, na.rm = TRUE)
  )
  
  z <- zoo::zoo(score, order.by = zoo::index(X_input))
  z <- clean_zoo(z, "joint_stress_cached")
  
  tail(z, T_out)
}

.normalize_delta_zoo <- function(z, eps = 1e-12) {
  z <- clean_zoo(z, "delta_ceci_normalization_input")
  
  x <- as.numeric(zoo::coredata(z))
  ok <- is.finite(x)
  
  out <- rep(NA_real_, length(x))
  
  if (sum(ok) < 2) {
    return(zoo::zoo(out, order.by = zoo::index(z)))
  }
  
  mu <- mean(x[ok], na.rm = TRUE)
  sdv <- stats::sd(x[ok], na.rm = TRUE)
  
  if (!is.finite(sdv) || sdv < eps) {
    out[ok] <- 0
  } else {
    out <- (x - mu) / sdv
  }
  
  clean_zoo(
    zoo::zoo(out, order.by = zoo::index(z)),
    "delta_ceci_normalized"
  )
}

.make_stress_bucket_delta <- function(x) {
  out <- rep(NA_character_, length(x))
  ok <- is.finite(x)
  
  if (sum(ok) == 0) {
    return(factor(out, levels = quantile_labels_delta))
  }
  
  if (sum(ok) == 1) {
    out[ok] <- "75-100"
    return(factor(out, levels = quantile_labels_delta))
  }
  
  r <- rank(x[ok], na.last = "keep", ties.method = "average")
  p <- (r - 1) / (length(r) - 1)
  
  out_ok <- rep(NA_character_, length(p))
  out_ok[p <= 0.25] <- "0-25"
  out_ok[p > 0.25 & p <= 0.50] <- "25-50"
  out_ok[p > 0.50 & p <= 0.75] <- "50-75"
  out_ok[p > 0.75] <- "75-100"
  
  out[ok] <- out_ok
  
  factor(out, levels = quantile_labels_delta)
}

# ------------------------------ MAIN SUMMARY ----------------------------------

cat("\n==================================================================\n")
cat("HEATMAP OF CACHED DIRECTIONAL DELTA CECI BY JOINT-STRESS QUANTILE\n")
cat("input:", mode_rolling, "| W:", W, "| nlag:", nlag, "| H:", nfore, "\n")
cat("models:", paste(models_delta_heatmap, collapse = ", "), "\n")
cat("joint stress:", stress_measure_delta, "\n")
cat("normalize delta CECI:", normalize_delta_ceci_heatmap, "\n")
cat("==================================================================\n\n")

delta_heatmap_cells <- list()
delta_cell_counter <- 1

for (commodity_i in commodities_delta_heatmap) {
  message("Building cached directional heatmap cells for commodity: ", commodity_i)
  
  res_i <- res_by_commodity[[commodity_i]]
  
  missing_models_i <- setdiff(models_delta_heatmap, names(res_i$rolling$rolling))
  if (length(missing_models_i) > 0) {
    stop(
      "Cached object for ", commodity_i,
      " is missing models: ",
      paste(missing_models_i, collapse = ", ")
    )
  }
  
  for (model_i in models_delta_heatmap) {
    message("  -> model: ", model_i)
    
    delta_ceci_z <- .extract_delta_ceci_from_cached_roll(
      res_obj = res_i,
      model_name = model_i,
      commodity_base = commodity_i,
      equity_regex = equity_regex,
      scale_100 = TRUE
    )
    
    if (isTRUE(normalize_delta_ceci_heatmap)) {
      delta_ceci_z <- .normalize_delta_zoo(
        delta_ceci_z,
        eps = normalization_eps_delta
      )
    }
    
    X_input_i <- res_i$rolling$X_inputs[[model_i]]
    
    joint_stress_z <- .compute_joint_stress_cached(
      X_input = X_input_i,
      T_out = NROW(delta_ceci_z),
      method = stress_measure_delta
    )
    
    Z_align <- merge_zoo_list_inner(
      list(joint_stress = joint_stress_z, delta_ceci = delta_ceci_z),
      names_out = c("joint_stress", "delta_ceci"),
      object_name = paste0("delta_heatmap_align_", commodity_i, "_", model_i)
    )
    
    df_i <- data.frame(
      date = as.Date(zoo::index(Z_align)),
      commodity = commodity_i,
      model = model_i,
      joint_stress = as.numeric(Z_align[, "joint_stress"]),
      delta_ceci = as.numeric(Z_align[, "delta_ceci"]),
      stringsAsFactors = FALSE
    ) |>
      dplyr::filter(is.finite(joint_stress), is.finite(delta_ceci))
    
    if (nrow(df_i) == 0) {
      warning(
        "No valid aligned observations for commodity=",
        commodity_i,
        ", model=",
        model_i
      )
      next
    }
    
    df_i <- df_i |>
      dplyr::mutate(
        stress_bucket = .make_stress_bucket_delta(joint_stress)
      )
    
    summ_i <- df_i |>
      dplyr::group_by(model, commodity, stress_bucket) |>
      dplyr::summarise(
        bucket_mean_delta_ceci = mean(delta_ceci, na.rm = TRUE),
        bucket_median_delta_ceci = median(delta_ceci, na.rm = TRUE),
        overall_mean_delta_ceci = mean(delta_ceci, na.rm = TRUE),
        overall_median_delta_ceci = median(delta_ceci, na.rm = TRUE),
        obs = dplyr::n(),
        joint_min = min(joint_stress, na.rm = TRUE),
        joint_max = max(joint_stress, na.rm = TRUE),
        .groups = "drop"
      )
    
    delta_heatmap_cells[[delta_cell_counter]] <- summ_i
    delta_cell_counter <- delta_cell_counter + 1
  }
}

if (length(delta_heatmap_cells) == 0) {
  stop("No directional delta heatmap cells were produced.")
}

delta_heatmap_df <- dplyr::bind_rows(delta_heatmap_cells) |>
  dplyr::mutate(
    commodity = factor(commodity, levels = commodities_delta_heatmap),
    stress_bucket = factor(stress_bucket, levels = quantile_labels_delta),
    model = factor(model, levels = models_delta_heatmap)
  ) |>
  tidyr::complete(
    model,
    commodity,
    stress_bucket,
    fill = list(
      bucket_mean_delta_ceci = NA_real_,
      bucket_median_delta_ceci = NA_real_,
      overall_mean_delta_ceci = NA_real_,
      overall_median_delta_ceci = NA_real_,
      obs = 0L,
      joint_min = NA_real_,
      joint_max = NA_real_
    )
  ) |>
  dplyr::mutate(
    commodity_label = factor(
      as.character(commodity_short_labels_delta[as.character(commodity)]),
      levels = unname(commodity_short_labels_delta[commodities_delta_heatmap])
    ),
    heat_value = bucket_mean_delta_ceci
  )

write.csv(
  delta_heatmap_df,
  out_csv_delta_heatmap_cached,
  row.names = FALSE
)

cat("Directional delta heatmap cell summary:\n")
print(
  delta_heatmap_df |>
    dplyr::arrange(model, commodity, stress_bucket) |>
    dplyr::select(
      model,
      commodity,
      stress_bucket,
      bucket_mean_delta_ceci,
      obs
    ),
  row.names = FALSE
)

# ------------------------------ HEATMAP PLOT ----------------------------------

fill_lim_delta <- max(abs(delta_heatmap_df$heat_value), na.rm = TRUE)
if (!is.finite(fill_lim_delta) || fill_lim_delta <= 0) {
  fill_lim_delta <- 1
}

pretty_title_delta <- c(
  DCC_full          = "DCC_full",
  DCC_scalar_stage  = "DCC_scalar_stage",
  cDCC_Aielli_stage = "cDCC_Aielli_stage",
  sBEKK_sym         = "sBEKK_sym",
  dBEKK_sym         = "dBEKK_sym",
  dBEKK_asym        = "dBEKK_asym"
)

base_delta_heatmap <- function(df_plot,
                               panel_title = "",
                               show_y_title = TRUE,
                               hide_y_text_visually = FALSE,
                               show_legend = TRUE,
                               show_x_title = TRUE,
                               show_x_text = TRUE) {
  p <- ggplot(
    df_plot,
    aes(x = stress_bucket, y = commodity_label, fill = heat_value)
  ) +
    geom_tile(color = "white", linewidth = 0.8) +
    geom_text(
      aes(label = ifelse(is.na(heat_value), "", sprintf("%.2f", heat_value))),
      size = 4.6
    ) +
    scale_fill_gradient2(
      low = "#2166AC",
      mid = "white",
      high = "#B2182B",
      midpoint = 0,
      limits = c(-fill_lim_delta, fill_lim_delta),
      oob = scales::squish,
      name = expression(Delta * " CECI")
    ) +
    labs(
      title = panel_title,
      x = if (show_x_title) "Joint Stress Quantile" else NULL,
      y = if (show_y_title) "Commodity" else NULL
    ) +
    theme_minimal(base_size = 16) +
    theme(
      panel.grid = element_blank(),
      axis.text.x = if (show_x_text) element_text(size = 13) else element_blank(),
      axis.title.x = if (show_x_title) element_text(size = 15) else element_blank(),
      axis.ticks.x = if (show_x_text) element_line() else element_blank(),
      plot.title = element_text(size = 17, face = "bold", hjust = 0.5),
      legend.position = if (show_legend) "right" else "none",
      legend.title = element_text(size = 13),
      legend.text = element_text(size = 11)
    )
  
  if (!hide_y_text_visually) {
    p <- p +
      theme(
        axis.text.y = element_text(size = 13),
        axis.title.y = if (show_y_title) element_text(size = 15) else element_blank(),
        axis.ticks.y = element_line()
      )
  } else {
    p <- p +
      theme(
        axis.text.y = element_text(size = 13, colour = scales::alpha("black", 0)),
        axis.title.y = element_blank(),
        axis.ticks.y = element_line(colour = scales::alpha("black", 0))
      )
  }
  
  p
}

make_delta_panel_grob <- function(model_name,
                                  show_y_title,
                                  hide_y_text_visually,
                                  show_legend = FALSE,
                                  show_x_title = TRUE,
                                  show_x_text = TRUE) {
  df_plot <- delta_heatmap_df |>
    dplyr::filter(model == model_name)
  
  p <- base_delta_heatmap(
    df_plot = df_plot,
    panel_title = pretty_title_delta[[model_name]],
    show_y_title = show_y_title,
    hide_y_text_visually = hide_y_text_visually,
    show_legend = show_legend,
    show_x_title = show_x_title,
    show_x_text = show_x_text
  )
  
  ggplotGrob(p)
}

g_delta_dcc1 <- make_delta_panel_grob(
  "DCC_full",
  TRUE,
  FALSE,
  FALSE,
  FALSE,
  FALSE
)

g_delta_dcc2 <- make_delta_panel_grob(
  "DCC_scalar_stage",
  TRUE,
  FALSE,
  FALSE,
  FALSE,
  FALSE
)

g_delta_dcc3 <- make_delta_panel_grob(
  "cDCC_Aielli_stage",
  TRUE,
  FALSE,
  FALSE,
  TRUE,
  TRUE
)

p_delta_bekk_for_legend <- base_delta_heatmap(
  df_plot = delta_heatmap_df |>
    dplyr::filter(model == "dBEKK_asym"),
  panel_title = pretty_title_delta[["dBEKK_asym"]],
  show_y_title = FALSE,
  hide_y_text_visually = TRUE,
  show_legend = TRUE,
  show_x_title = TRUE,
  show_x_text = TRUE
)

legend_grob_delta <- get_legend_grob(p_delta_bekk_for_legend)

g_delta_bekk1 <- make_delta_panel_grob(
  "sBEKK_sym",
  FALSE,
  TRUE,
  FALSE,
  FALSE,
  FALSE
)

g_delta_bekk2 <- make_delta_panel_grob(
  "dBEKK_sym",
  FALSE,
  TRUE,
  FALSE,
  FALSE,
  FALSE
)

g_delta_bekk3 <- ggplotGrob(
  base_delta_heatmap(
    df_plot = delta_heatmap_df |>
      dplyr::filter(model == "dBEKK_asym"),
    panel_title = pretty_title_delta[["dBEKK_asym"]],
    show_y_title = FALSE,
    hide_y_text_visually = TRUE,
    show_legend = FALSE,
    show_x_title = TRUE,
    show_x_text = TRUE
  )
)

all_delta_panels <- list(
  g_delta_dcc1,
  g_delta_bekk1,
  g_delta_dcc2,
  g_delta_bekk2,
  g_delta_dcc3,
  g_delta_bekk3
)

all_delta_panels <- equalize_panel_widths(all_delta_panels)
all_delta_panels <- equalize_panel_heights(all_delta_panels)

panel_grid_delta <- gridExtra::arrangeGrob(
  grobs = all_delta_panels,
  ncol = 2,
  widths = c(1, 1)
)

combined_plots_delta <- gridExtra::arrangeGrob(
  grobs = list(panel_grid_delta, legend_grob_delta),
  ncol = 2,
  widths = c(1, 0.10)
)

title_grob_delta <- grid::textGrob(
  "Directional commodity-equity connectedness by joint-stress quantile",
  gp = grid::gpar(fontsize = 22, fontface = "bold")
)

subtitle_grob_delta <- grid::textGrob(
  paste0(
    "Heatmap cells show mean(",
    if (normalize_delta_ceci_heatmap) "z-score " else "",
    "\u0394CECI | joint-stress bucket); input: ",
    mode_rolling,
    ", W=",
    W,
    ", nlag=",
    nlag,
    ", H=",
    nfore
  ),
  gp = grid::gpar(fontsize = 14)
)

column_header_delta <- grid::textGrob(
  "Left column: DCC specifications    |    Right column: BEKK specifications",
  gp = grid::gpar(fontsize = 14)
)

note_grob_delta <- grid::textGrob(
  paste0(
    "Joint stress is computed as the ",
    stress_measure_delta,
    " of column-standardized transformed volatility inputs used in each cached rolling connectedness system."
  ),
  gp = grid::gpar(fontsize = 12)
)

final_plot_delta <- gridExtra::arrangeGrob(
  title_grob_delta,
  subtitle_grob_delta,
  column_header_delta,
  combined_plots_delta,
  note_grob_delta,
  ncol = 1,
  heights = c(0.05, 0.05, 0.04, 0.80, 0.06)
)

save_png(
  out_png_delta_heatmap_cached,
  {
    grid::grid.newpage()
    grid::grid.draw(final_plot_delta)
  },
  width = 3200,
  height = 2600,
  res = 220
)

cat("\nDirectional Delta CECI heatmap add-on complete.\n")
cat("Heatmap PNG :", out_png_delta_heatmap_cached, "\n")
cat("Summary CSV :", out_csv_delta_heatmap_cached, "\n\n")
