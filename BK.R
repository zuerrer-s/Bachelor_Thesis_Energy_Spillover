################################################################################
# STANDALONE DROP-IN SCRIPT — BK FREQUENCY CECI + SHARE PLOTS
#
# PURPOSE
#   Barunik-Krehlik frequency-domain CECI using the same universe logic:
#     one commodity + all SX* equities
#
# OUTPUTS
#   1) bk_ceci_res object in memory
#   2) RDS cache of bk_ceci_res
#   3) BK model-comparison PNG for one commodity
#   4) BK stacked benchmark PNG for one commodity
#   5) BK 6-panel stacked-share PNG for one commodity
#
# INPUT REQUIRED
#   - vol_models in memory, or vol_models_all_7models.rds in output_dir/base_dir
#
# PACKAGE REQUIRED
#   - frequencyConnectedness
################################################################################

suppressPackageStartupMessages({
  library(zoo)
  library(vars)
  library(frequencyConnectedness)
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

FORCE_RECOMPUTE_BK <- FALSE

models_used <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym",
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage"
)

models_used_plot <- c(
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

# BK settings
W_bk     <- 200
nlag_bk  <- 1
nfore_bk <- 100

# Use raw volatility levels by default, matching your pasted BK script.
# Set to "dlog_var" if you want BK inputs to match the DY transformed-volatility input.
BK_INPUT_TRANSFORM <- "level"     # "level", "loglevel", "dlog_vol", "dlog_var"

USE_SQRT_ON_INPUT_BK <- FALSE
equity_regex <- "^SX"

make_partition_daily <- function() {
  c(pi + 1e-5, pi / 5, pi / 20, 0)
}

FREQ_LABELS <- c("Short (1-5d)", "Medium (5-20d)", "Long (20+d)")

# "share_total", "raw", or "avg_pair"
BK_CECI_SCALE <- "share_total"

bk_plot_commodity <- "TTF (Gas)"
benchmark_model <- "DCC_full"
bk_direction <- "bidir"   # "bidir", "c2e", "e2c"

bk_cache_file <- file.path(
  output_dir,
  paste0(
    "bk_ceci_res_",
    "input_", BK_INPUT_TRANSFORM,
    "_w", W_bk,
    "_p", nlag_bk,
    "_H", nfore_bk,
    "_scale_", BK_CECI_SCALE,
    ".rds"
  )
)

out_bk_ceci_compare_png <- file.path(
  output_dir,
  paste0("BK_CECI_ModelComparison_", gsub("[^A-Za-z0-9]+", "_", bk_plot_commodity), "_", bk_direction, ".png")
)

out_bk_ceci_stack_png <- file.path(
  output_dir,
  paste0("BK_CECI_Stacked_", benchmark_model, "_", gsub("[^A-Za-z0-9]+", "_", bk_plot_commodity), "_", bk_direction, ".png")
)

out_bk_ceci_share_6panel_png <- file.path(
  output_dir,
  paste0("BK_CECI_StackedShare_6Panel_", gsub("[^A-Za-z0-9]+", "_", bk_plot_commodity), "_", bk_direction, ".png")
)

# Stress episodes
shade_episodes <- FALSE
label_episodes <- FALSE

episodes_named <- data.frame(
  label = c(
    "2015-16 Equity Stress",
    "COVID-19",
    "Russia-Ukraine War",
    "Liberation Day"
  ),
  start = as.Date(c(
    "2015-08-01",
    "2020-01-01",
    "2021-10-01",
    "2025-02-01"
  )),
  end = as.Date(c(
    "2017-04-30",
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

# Graphics
cex_axis <- 1.15
cex_lab  <- 1.30
cex_main <- 1.20
cex_leg  <- 1.10
cex_sub  <- 1.00

legend_ncol <- 3
lwd <- 1.35

left_margin_lines  <- 5.6
right_margin_lines <- 2.0

png_width  <- 3200
png_height <- 2600
png_res    <- 220

fill_cols <- c(
  "Short (1-5d)"   = adjustcolor("gold",        alpha.f = 0.85),
  "Medium (5-20d)" = adjustcolor("chartreuse3", alpha.f = 0.85),
  "Long (20+d)"    = adjustcolor("red2",        alpha.f = 0.85)
)

pretty_title <- c(
  DCC_full          = "DCC_full",
  DCC_scalar_stage  = "DCC_scalar_stage",
  cDCC_Aielli_stage = "cDCC_Aielli_stage",
  sBEKK_sym         = "sBEKK_sym",
  dBEKK_sym         = "dBEKK_sym",
  dBEKK_asym        = "dBEKK_asym"
)

# ==============================================================================
# 2) GENERAL HELPERS
# ==============================================================================

save_png <- function(filename, expr, width = 2600, height = 1800, res = 220) {
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
  if (inherits(idx, "Date")) return(idx)
  if (inherits(idx, c("POSIXct", "POSIXt"))) return(as.Date(idx))
  if (inherits(idx, c("yearmon", "yearqtr"))) return(as.Date(idx))
  if (is.numeric(idx)) return(as.Date(idx, origin = "1970-01-01"))
  
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
      lapply(split(seq_along(idx), idx), function(ii) tail(ii, 1)),
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

.ensure_Date_index <- function(z) clean_zoo(z, "ensure_Date_index")

.ensure_vol_colnames <- function(vol_zoo, model_name = "model") {
  stopifnot(inherits(vol_zoo, "zoo"))
  cn <- colnames(vol_zoo)
  if (is.null(cn)) stop(model_name, ": vol_zoo has no colnames().")
  colnames(vol_zoo) <- ifelse(grepl("_vol$", cn), cn, paste0(cn, "_vol"))
  vol_zoo
}

.strip_vol <- function(x) sub("_vol$", "", x)

.norm_name <- function(x) {
  x <- tolower(x)
  gsub("[^a-z0-9]+", "", x)
}

.match_base_name <- function(requested_base, available_base) {
  req_n <- .norm_name(requested_base)
  av_n  <- .norm_name(available_base)
  
  hit <- which(av_n == req_n)
  if (length(hit) >= 1) return(available_base[hit[1]])
  
  hit <- which(grepl(req_n, av_n, fixed = TRUE) | grepl(av_n, req_n, fixed = TRUE))
  if (length(hit) >= 1) return(available_base[hit[1]])
  
  NA_character_
}

.default_cols <- function(models) {
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

.default_ltys <- function(models) {
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

.axis_years <- function(x_dates, cex_axis = 1.0) {
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

.shade_episodes <- function(episodes, shade_col = gray(0.90, alpha = 0.95), x_range = NULL) {
  if (is.null(episodes) || nrow(episodes) == 0) return(invisible(NULL))
  
  usr <- par("usr")
  
  for (i in seq_len(nrow(episodes))) {
    xleft  <- episodes$start[i]
    xright <- episodes$end[i]
    
    if (!is.null(x_range)) {
      xleft  <- max(xleft, min(x_range))
      xright <- min(xright, max(x_range))
      if (xleft > xright) next
    }
    
    rect(
      xleft = xleft,
      ybottom = usr[3],
      xright = xright,
      ytop = usr[4],
      col = shade_col,
      border = NA
    )
  }
  
  invisible(NULL)
}

.label_episodes_top <- function(episodes,
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

# ==============================================================================
# 3) BK INPUT AND ESTIMATION HELPERS
# ==============================================================================

transform_bk_input <- function(z, transform = c("level", "loglevel", "dlog_vol", "dlog_var")) {
  transform <- match.arg(transform)
  
  z <- clean_zoo(z, "BK input before transform")
  idx <- zoo::index(z)
  
  X <- zoo::coredata(z)
  if (!is.matrix(X)) X <- as.matrix(X)
  
  X <- pmax(X, 1e-12)
  
  if (transform == "level") {
    Y <- X
    yidx <- idx
  } else if (transform == "loglevel") {
    Y <- log(X)
    yidx <- idx
  } else if (transform == "dlog_vol") {
    Y <- diff(log(X))
    yidx <- idx[-1]
  } else if (transform == "dlog_var") {
    Y <- diff(log(X^2))
    yidx <- idx[-1]
  }
  
  colnames(Y) <- colnames(z)
  
  ok <- stats::complete.cases(Y)
  Y <- Y[ok, , drop = FALSE]
  yidx <- yidx[ok]
  
  out <- zoo::zoo(Y, order.by = yidx)
  colnames(out) <- colnames(z)
  clean_zoo(out, "BK transformed input")
}

get_vol_input_bk <- function(vol_models,
                             model_name,
                             use_sqrt = FALSE,
                             input_transform = "level") {
  stopifnot(model_name %in% names(vol_models))
  
  z <- vol_models[[model_name]]
  stopifnot(inherits(z, "zoo"))
  
  z <- .ensure_vol_colnames(z, model_name)
  z <- clean_zoo(z, model_name)
  
  keep <- vapply(as.data.frame(z), is.numeric, logical(1))
  z <- z[, keep, drop = FALSE]
  z <- z[, colSums(!is.na(z)) > 0, drop = FALSE]
  z <- z[complete.cases(z), , drop = FALSE]
  
  if (use_sqrt) {
    z <- sqrt(pmax(z, 0))
  }
  
  transform_bk_input(z, transform = input_transform)
}

get_bk_ceci_universe <- function(vol_models,
                                 model_name,
                                 commodity_requested,
                                 use_sqrt = FALSE,
                                 input_transform = "level",
                                 equity_prefix = "^SX") {
  z <- get_vol_input_bk(
    vol_models = vol_models,
    model_name = model_name,
    use_sqrt = use_sqrt,
    input_transform = input_transform
  )
  
  base_names <- .strip_vol(colnames(z))
  
  com_base <- .match_base_name(commodity_requested, base_names)
  if (is.na(com_base)) {
    stop("Commodity not found in ", model_name, ": ", commodity_requested)
  }
  
  eq_idx <- grep(equity_prefix, base_names)
  if (length(eq_idx) == 0) {
    stop("No equity columns found in ", model_name, " using prefix ", equity_prefix)
  }
  
  # Commodity first, then equities
  keep_bases <- unique(c(com_base, base_names[eq_idx]))
  keep_cols <- paste0(keep_bases, "_vol")
  keep_cols <- keep_cols[keep_cols %in% colnames(z)]
  
  z_sub <- z[, keep_cols, drop = FALSE]
  z_sub <- z_sub[complete.cases(z_sub), , drop = FALSE]
  z_sub <- clean_zoo(z_sub, paste0("BK universe ", model_name, " ", commodity_requested))
  
  list(
    data = z_sub,
    commodity_base = com_base,
    equity_bases = setdiff(.strip_vol(colnames(z_sub)), com_base)
  )
}

run_bk_ceci_universe <- function(vol_zoo,
                                 window = 200,
                                 nlag = 1,
                                 nfore = 100,
                                 partition = make_partition_daily(),
                                 no.corr = FALSE) {
  stopifnot(inherits(vol_zoo, "zoo"))
  vol_zoo <- clean_zoo(vol_zoo, "BK rolling input")
  
  frequencyConnectedness::spilloverRollingBK12(
    data = vol_zoo,
    n.ahead = nfore,
    no.corr = no.corr,
    partition = partition,
    func_est = "VAR",
    params_est = list(
      p = nlag,
      type = "const"
    ),
    window = window
  )
}

get_bk_roll_dates <- function(input_zoo, window, n_out = NULL) {
  idx <- zoo::index(input_zoo)
  out <- idx[window:length(idx)]
  
  if (!is.null(n_out) && length(out) != n_out) {
    if (length(out) > n_out) {
      out <- tail(out, n_out)
    } else {
      warning("Date length mismatch in BK roll dates. Returning available dates.")
    }
  }
  
  as.Date(out)
}

as_clean_square_matrix <- function(M) {
  if (is.null(M)) stop("Band table is NULL.")
  
  if (inherits(M, "spillover_table") && !is.null(M$table)) {
    M <- M$table
  }
  
  M <- as.matrix(M)
  
  rn <- rownames(M)
  cn <- colnames(M)
  
  if (!is.null(rn) && !is.null(cn)) {
    common <- intersect(rn, cn)
    if (length(common) >= 2) {
      M <- M[common, common, drop = FALSE]
    }
  }
  
  M
}

.clean_nm <- function(x) {
  x <- sub("_vol$", "", x)
  x <- sub("_ret$", "", x)
  x <- tolower(x)
  gsub("[^a-z0-9]", "", x)
}

compute_cross_block_stats <- function(M,
                                      commodity_base,
                                      equity_bases,
                                      scale = c("share_total", "raw", "avg_pair")) {
  scale <- match.arg(scale)
  M <- as_clean_square_matrix(M)
  
  rn <- rownames(M)
  cn <- colnames(M)
  
  if (is.null(rn) || is.null(cn)) {
    stop("Band matrix needs rownames and colnames.")
  }
  
  rn_clean <- .clean_nm(rn)
  cn_clean <- .clean_nm(cn)
  
  com_clean <- .clean_nm(commodity_base)
  eq_clean <- .clean_nm(equity_bases)
  
  com_r <- match(com_clean, rn_clean)
  com_c <- match(com_clean, cn_clean)
  eq_r  <- match(eq_clean, rn_clean)
  eq_c  <- match(eq_clean, cn_clean)
  
  eq_r <- eq_r[!is.na(eq_r)]
  eq_c <- eq_c[!is.na(eq_c)]
  
  if (is.na(com_r) || is.na(com_c) || length(eq_r) == 0 || length(eq_c) == 0) {
    stop("Could not identify commodity/equity blocks in BK matrix.")
  }
  
  # Matrix convention: M[i,j] = shocks in j contributing to variance of i.
  # Commodity -> Equity: commodity shock contributes to equity variance.
  c2e_raw <- sum(M[eq_r, com_c, drop = FALSE], na.rm = TRUE)
  
  # Equity -> Commodity: equity shocks contribute to commodity variance.
  e2c_raw <- sum(M[com_r, eq_c, drop = FALSE], na.rm = TRUE)
  
  bidir_raw <- c2e_raw + e2c_raw
  total_sum <- sum(M, na.rm = TRUE)
  
  if (scale == "raw") {
    return(c(c2e = c2e_raw, e2c = e2c_raw, bidir = bidir_raw))
  }
  
  if (scale == "share_total") {
    return(100 * c(c2e = c2e_raw, e2c = e2c_raw, bidir = bidir_raw) / total_sum)
  }
  
  n_eq <- length(eq_r)
  
  100 * c(
    c2e = c2e_raw / n_eq,
    e2c = e2c_raw / n_eq,
    bidir = bidir_raw / (2 * n_eq)
  )
}

extract_bk_ceci_series <- function(bk_roll,
                                   input_zoo,
                                   window,
                                   commodity_base,
                                   equity_bases,
                                   freq_labels = FREQ_LABELS,
                                   scale = BK_CECI_SCALE) {
  if (is.null(bk_roll$list_of_tables)) {
    stop("BK rolling object does not contain $list_of_tables.")
  }
  
  lot <- bk_roll$list_of_tables
  nT <- length(lot)
  if (nT == 0) stop("Empty BK rolling object.")
  
  first_tables <- lot[[1]]$tables
  nB <- length(first_tables)
  
  if (length(freq_labels) != nB) {
    freq_labels <- paste("Band", seq_len(nB))
  }
  
  c2e_mat <- matrix(NA_real_, nrow = nT, ncol = nB)
  e2c_mat <- matrix(NA_real_, nrow = nT, ncol = nB)
  bidir_mat <- matrix(NA_real_, nrow = nT, ncol = nB)
  
  for (tt in seq_len(nT)) {
    band_tables <- lot[[tt]]$tables
    
    for (bb in seq_len(nB)) {
      stats_bb <- compute_cross_block_stats(
        M = band_tables[[bb]],
        commodity_base = commodity_base,
        equity_bases = equity_bases,
        scale = scale
      )
      
      c2e_mat[tt, bb] <- stats_bb["c2e"]
      e2c_mat[tt, bb] <- stats_bb["e2c"]
      bidir_mat[tt, bb] <- stats_bb["bidir"]
    }
  }
  
  dates_out <- get_bk_roll_dates(input_zoo, window = window, n_out = nT)
  
  c2e_z <- zoo::zoo(c2e_mat, order.by = dates_out)
  e2c_z <- zoo::zoo(e2c_mat, order.by = dates_out)
  bidir_z <- zoo::zoo(bidir_mat, order.by = dates_out)
  
  colnames(c2e_z) <- freq_labels
  colnames(e2c_z) <- freq_labels
  colnames(bidir_z) <- freq_labels
  
  list(
    c2e = clean_zoo(c2e_z, "BK c2e"),
    e2c = clean_zoo(e2c_z, "BK e2c"),
    bidir = clean_zoo(bidir_z, "BK bidir")
  )
}

run_bk_ceci_all <- function(vol_models,
                            models_used,
                            commodities,
                            use_sqrt = FALSE,
                            input_transform = "level",
                            window = 200,
                            nlag = 1,
                            nfore = 100,
                            partition = make_partition_daily(),
                            no.corr = FALSE,
                            scale = BK_CECI_SCALE,
                            freq_labels = FREQ_LABELS) {
  out <- list()
  
  for (com in commodities) {
    cat("========== Commodity:", com, "==========\n")
    
    out[[com]] <- list(
      bidir = list(),
      c2e = list(),
      e2c = list(),
      raw_bk = list(),
      input = list()
    )
    
    for (m in models_used) {
      cat("Running BK-CECI for:", m, "|", com, "\n")
      
      uni <- get_bk_ceci_universe(
        vol_models = vol_models,
        model_name = m,
        commodity_requested = com,
        use_sqrt = use_sqrt,
        input_transform = input_transform,
        equity_prefix = equity_regex
      )
      
      bk_roll <- run_bk_ceci_universe(
        vol_zoo = uni$data,
        window = window,
        nlag = nlag,
        nfore = nfore,
        partition = partition,
        no.corr = no.corr
      )
      
      agg <- extract_bk_ceci_series(
        bk_roll = bk_roll,
        input_zoo = uni$data,
        window = window,
        commodity_base = uni$commodity_base,
        equity_bases = uni$equity_bases,
        freq_labels = freq_labels,
        scale = scale
      )
      
      out[[com]]$bidir[[m]] <- agg$bidir
      out[[com]]$c2e[[m]] <- agg$c2e
      out[[com]]$e2c[[m]] <- agg$e2c
      out[[com]]$raw_bk[[m]] <- bk_roll
      out[[com]]$input[[m]] <- uni$data
    }
  }
  
  attr(out, "settings") <- list(
    input_transform = input_transform,
    window = window,
    nlag = nlag,
    nfore = nfore,
    scale = scale,
    freq_labels = freq_labels
  )
  
  out
}

# ==============================================================================
# 4) PLOTTING HELPERS
# ==============================================================================

plot_bk_ceci_compare <- function(bk_ceci_res,
                                 commodity_requested,
                                 direction = c("bidir", "c2e", "e2c"),
                                 benchmark = "DCC_full",
                                 main = NULL,
                                 outfile = NULL,
                                 include_dev = TRUE,
                                 shade_episodes = TRUE,
                                 label_episodes = TRUE,
                                 episodes_named = NULL) {
  direction <- match.arg(direction)
  
  if (!(commodity_requested %in% names(bk_ceci_res))) {
    stop("Commodity not found in bk_ceci_res: ", commodity_requested)
  }
  
  Zlist <- bk_ceci_res[[commodity_requested]][[direction]]
  models <- names(Zlist)
  
  if (!(benchmark %in% models)) {
    stop("benchmark not found in BK-CECI models.")
  }
  
  cols <- .default_cols(models)
  ltys <- .default_ltys(models)
  
  totals <- lapply(Zlist, function(z) {
    z <- clean_zoo(z, "BK totals input")
    zoo::zoo(rowSums(zoo::coredata(z), na.rm = TRUE), order.by = zoo::index(z))
  })
  
  Tmat <- do.call(merge, c(totals, all = FALSE))
  colnames(Tmat) <- models
  Tmat <- clean_zoo(Tmat, "BK total merged")
  
  Dmat <- Tmat
  for (nm in models) {
    Dmat[, nm] <- Tmat[, nm] - Tmat[, benchmark]
  }
  
  z0 <- clean_zoo(Zlist[[1]], "BK first band object")
  nb <- NCOL(z0)
  
  k <- nb + ifelse(include_dev, 1, 0)
  
  if (is.null(main)) {
    dir_lab <- switch(
      direction,
      bidir = "Bidirectional BK-CECI",
      c2e = "Commodity -> Equity BK-CECI",
      e2c = "Equity -> Commodity BK-CECI"
    )
    
    main <- paste0(dir_lab, " — ", commodity_requested, " — benchmark: ", benchmark)
  }
  
  save_png(outfile, {
    op <- par(no.readonly = TRUE)
    on.exit(par(op), add = TRUE)
    
    par(
      mfrow = c(k, 1),
      oma = c(8, 0, 3.5, 0),
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
    
    maybe_add_episode_labels <- function(x_range) {
      if (
        isTRUE(label_episodes) &&
        !top_panel_labeled &&
        !is.null(episodes_named) &&
        nrow(episodes_named) > 0
      ) {
        .label_episodes_top(
          episodes = episodes_named,
          label_x_shift_days = label_x_shift_days,
          label_y_frac_from_top = label_y_frac_from_top,
          cex = episode_label_cex,
          col = episode_label_col,
          x_range = x_range
        )
        top_panel_labeled <<- TRUE
      }
    }
    
    draw_order <- unique(c(
      setdiff(models, c(benchmark, "DCC_scalar_stage", "cDCC_Aielli_stage")),
      benchmark,
      "DCC_scalar_stage",
      "cDCC_Aielli_stage"
    ))
    
    draw_order <- draw_order[draw_order %in% models]
    
    for (bb in seq_len(nb)) {
      set_bottom <- if (include_dev) FALSE else (bb == nb)
      
      par(mar = c(ifelse(set_bottom, 4, 0), 5.8, ifelse(bb == 1, 2, 0), 2) + 0.1)
      
      band_list <- lapply(Zlist, function(z) clean_zoo(z, "BK band")[, bb, drop = FALSE])
      Zband <- do.call(merge, c(band_list, all = FALSE))
      colnames(Zband) <- models
      Zband <- clean_zoo(Zband, "BK band merged")
      
      x <- zoo::index(Zband)
      yr <- range(zoo::coredata(Zband), na.rm = TRUE, finite = TRUE)
      
      plot(
        x,
        zoo::coredata(Zband[, 1]),
        type = "n",
        xlab = "",
        xaxt = "n",
        ylab = "BK-CECI",
        main = colnames(z0)[bb],
        ylim = yr
      )
      
      add_global_title_once()
      
      if (isTRUE(shade_episodes) && !is.null(episodes_named) && nrow(episodes_named) > 0) {
        .shade_episodes(episodes_named, shade_col = episode_shade_col, x_range = x)
      }
      
      box()
      
      for (nm in draw_order) {
        if (nm == benchmark) next
        lines(x, zoo::coredata(Zband[, nm]), col = cols[nm], lty = ltys[nm], lwd = lwd)
      }
      
      lines(
        x,
        zoo::coredata(Zband[, benchmark]),
        col = cols[benchmark],
        lty = ltys[benchmark],
        lwd = lwd * 2
      )
      
      maybe_add_episode_labels(x)
      
      if (!include_dev && bb == nb) {
        .axis_years(x, cex_axis = cex_axis)
      }
    }
    
    if (isTRUE(include_dev)) {
      par(mar = c(4, 5.8, 0, 2) + 0.1)
      
      x <- zoo::index(Dmat)
      yr <- range(zoo::coredata(Dmat), na.rm = TRUE, finite = TRUE)
      mdev <- max(abs(yr), na.rm = TRUE)
      
      plot(
        x,
        zoo::coredata(Dmat[, 1]),
        type = "n",
        xlab = "",
        xaxt = "n",
        ylab = "Delta total BK-CECI",
        ylim = c(-mdev, mdev),
        main = NULL
      )
      
      if (isTRUE(shade_episodes) && !is.null(episodes_named) && nrow(episodes_named) > 0) {
        .shade_episodes(episodes_named, shade_col = episode_shade_col, x_range = x)
      }
      
      box()
      abline(h = 0, lty = 3)
      
      for (nm in setdiff(models, benchmark)) {
        lines(x, zoo::coredata(Dmat[, nm]), col = cols[nm], lty = ltys[nm], lwd = lwd)
      }
      
      .axis_years(x, cex_axis = cex_axis)
    }
    
    par(xpd = NA)
    
    legend(
      "bottom",
      inset = c(0, -0.36),
      legend = c(paste0(benchmark, " (benchmark)"), setdiff(models, benchmark)),
      col = c(cols[benchmark], cols[setdiff(models, benchmark)]),
      lty = c(ltys[benchmark], ltys[setdiff(models, benchmark)]),
      lwd = c(lwd * 2, rep(lwd, length(setdiff(models, benchmark)))),
      bty = "n",
      cex = cex_leg,
      ncol = legend_ncol
    )
  }, width = 2800, height = 2400, res = 220)
}

plot_bk_ceci_stacked <- function(bk_ceci_res,
                                 commodity_requested,
                                 model_name = "DCC_full",
                                 direction = c("bidir", "c2e", "e2c"),
                                 outfile = NULL,
                                 main = NULL,
                                 shade_episodes = TRUE,
                                 label_episodes = TRUE,
                                 episodes_named = NULL) {
  direction <- match.arg(direction)
  
  z <- bk_ceci_res[[commodity_requested]][[direction]][[model_name]]
  z <- clean_zoo(z, "BK stacked input")
  
  cn <- colnames(z)
  cn_low <- tolower(cn)
  
  idx_short <- grep("short", cn_low)
  idx_medium <- grep("medium", cn_low)
  idx_long <- grep("long", cn_low)
  
  if (length(idx_short) != 1 || length(idx_medium) != 1 || length(idx_long) != 1) {
    stop("Could not identify Short/Medium/Long columns in BK-CECI object.")
  }
  
  z <- z[, c(idx_short, idx_medium, idx_long), drop = FALSE]
  colnames(z) <- c("Short (1-5d)", "Medium (5-20d)", "Long (20+d)")
  
  X <- zoo::coredata(z)
  x <- zoo::index(z)
  
  lower <- cbind(rep(0, nrow(X)), X[, 1], X[, 1] + X[, 2])
  upper <- cbind(X[, 1], X[, 1] + X[, 2], X[, 1] + X[, 2] + X[, 3])
  
  if (is.null(main)) {
    dlab <- switch(
      direction,
      bidir = "Bidirectional BK-CECI",
      c2e = "Commodity -> Equity BK-CECI",
      e2c = "Equity -> Commodity BK-CECI"
    )
    
    main <- paste0(model_name, ": ", dlab, " — ", commodity_requested)
  }
  
  save_png(outfile, {
    op <- par(no.readonly = TRUE)
    on.exit(par(op), add = TRUE)
    
    par(
      mar = c(4.5, 5.5, 3, 1.5),
      cex.axis = cex_axis,
      cex.lab = cex_lab
    )
    
    ymax <- max(upper[, 3], na.rm = TRUE)
    
    plot(
      x,
      upper[, 3],
      type = "n",
      xlab = "",
      ylab = "BK-CECI",
      main = main,
      xaxt = "n",
      ylim = c(0, max(200, ymax, na.rm = TRUE))
    )
    
    if (isTRUE(shade_episodes) && !is.null(episodes_named) && nrow(episodes_named) > 0) {
      .shade_episodes(episodes_named, shade_col = episode_shade_col, x_range = x)
    }
    
    for (j in 1:3) {
      polygon(
        x = c(x, rev(x)),
        y = c(lower[, j], rev(upper[, j])),
        col = fill_cols[colnames(z)[j]],
        border = NA
      )
    }
    
    lines(x, upper[, 1], lwd = 0.5, col = gray(0.35))
    lines(x, upper[, 2], lwd = 0.5, col = gray(0.35))
    lines(x, upper[, 3], lwd = 0.7, col = gray(0.20))
    
    .axis_years(x, cex_axis = cex_axis)
    box()
    
    if (isTRUE(label_episodes) && !is.null(episodes_named) && nrow(episodes_named) > 0) {
      .label_episodes_top(
        episodes = episodes_named,
        label_x_shift_days = label_x_shift_days,
        label_y_frac_from_top = label_y_frac_from_top,
        cex = episode_label_cex,
        col = episode_label_col,
        x_range = x
      )
    }
    
    legend(
      "topleft",
      legend = colnames(z),
      fill = fill_cols[colnames(z)],
      bty = "n",
      cex = cex_leg
    )
  }, width = 2800, height = 1800, res = 220)
}

.get_bk_share_object <- function(bk_ceci_res, commodity_requested, direction, model_name) {
  if (is.null(bk_ceci_res[[commodity_requested]])) {
    stop("Commodity not found in bk_ceci_res: ", commodity_requested)
  }
  
  if (is.null(bk_ceci_res[[commodity_requested]][[direction]])) {
    stop("Direction not found for commodity in bk_ceci_res: ", direction)
  }
  
  if (is.null(bk_ceci_res[[commodity_requested]][[direction]][[model_name]])) {
    stop("Model not found for commodity/direction in bk_ceci_res: ", model_name)
  }
  
  z <- bk_ceci_res[[commodity_requested]][[direction]][[model_name]]
  z <- clean_zoo(z, paste0("BK share ", commodity_requested, " ", model_name))
  
  cn <- colnames(z)
  cn_low <- tolower(cn)
  
  idx_short <- grep("short", cn_low)
  idx_medium <- grep("medium", cn_low)
  idx_long <- grep("long", cn_low)
  
  if (length(idx_short) != 1 || length(idx_medium) != 1 || length(idx_long) != 1) {
    stop("Could not identify Short/Medium/Long columns in BK-CECI object for model ", model_name)
  }
  
  z <- z[, c(idx_short, idx_medium, idx_long), drop = FALSE]
  colnames(z) <- c("Short (1-5d)", "Medium (5-20d)", "Long (20+d)")
  z
}

.draw_single_share_panel <- function(z,
                                     panel_title,
                                     show_y_axis = TRUE,
                                     show_x_axis = TRUE,
                                     shade_episodes = FALSE,
                                     label_episodes = FALSE,
                                     episodes_named = NULL) {
  X <- zoo::coredata(z)
  total <- rowSums(X, na.rm = TRUE)
  total[total <= 0 | !is.finite(total)] <- NA_real_
  
  S <- X / total
  x <- zoo::index(z)
  
  lower <- cbind(rep(0, nrow(S)), S[, 1], S[, 1] + S[, 2])
  upper <- cbind(S[, 1], S[, 1] + S[, 2], S[, 1] + S[, 2] + S[, 3])
  
  plot(
    x,
    upper[, 3],
    type = "n",
    xlab = "",
    ylab = if (show_y_axis) "Share of total BK-CECI" else "",
    main = panel_title,
    xaxt = "n",
    yaxt = "n",
    bty = "n",
    ylim = c(0, 1)
  )
  
  if (isTRUE(shade_episodes) && !is.null(episodes_named)) {
    .shade_episodes(
      episodes = episodes_named,
      shade_col = episode_shade_col,
      x_range = x
    )
  }
  
  for (j in 1:3) {
    polygon(
      x = c(x, rev(x)),
      y = c(lower[, j], rev(upper[, j])),
      col = fill_cols[colnames(z)[j]],
      border = NA
    )
  }
  
  box()
  
  if (show_y_axis) {
    axis(
      2,
      at = seq(0, 1, by = 0.2),
      labels = paste0(seq(0, 100, by = 20), "%"),
      cex.axis = cex_axis
    )
  }
  
  if (show_x_axis) {
    .axis_years(x, cex_axis = cex_axis)
  }
  
  if (isTRUE(label_episodes) && !is.null(episodes_named)) {
    .label_episodes_top(
      episodes = episodes_named,
      label_x_shift_days = label_x_shift_days,
      label_y_frac_from_top = label_y_frac_from_top,
      cex = episode_label_cex,
      col = episode_label_col,
      x_range = x
    )
  }
}

plot_bk_ceci_stacked_share_6panel <- function(bk_ceci_res,
                                              commodity_requested,
                                              direction = c("bidir", "c2e", "e2c"),
                                              outfile = NULL,
                                              shade_episodes = FALSE,
                                              label_episodes = FALSE,
                                              episodes_named = NULL,
                                              width = 3200,
                                              height = 2600,
                                              res = 220) {
  direction <- match.arg(direction)
  
  if (is.null(outfile)) {
    stop("Please provide outfile.")
  }
  
  title_text <- switch(
    direction,
    bidir = paste0("Frequency composition of bidirectional BK-CECI — ", commodity_requested),
    c2e = paste0("Frequency composition of commodity -> equity BK-CECI — ", commodity_requested),
    e2c = paste0("Frequency composition of equity -> commodity BK-CECI — ", commodity_requested)
  )
  
  subtitle_text <- "Left column: DCC specifications    |    Right column: BEKK specifications"
  
  save_png(outfile, {
    op <- par(no.readonly = TRUE)
    
    on.exit({
      layout(1)
      par(op)
    }, add = TRUE)
    
    layout(
      matrix(
        c(
          1, 2,
          3, 4,
          5, 6,
          7, 7
        ),
        nrow = 4,
        byrow = TRUE
      ),
      heights = c(8.5, 8.5, 8.5, 1.9)
    )
    
    par(
      oma = c(2, 0, 4, 0),
      cex.axis = cex_axis,
      cex.lab = cex_lab
    )
    
    par(mar = c(0.8, left_margin_lines, 3.0, right_margin_lines))
    .draw_single_share_panel(
      z = .get_bk_share_object(bk_ceci_res, commodity_requested, direction, "DCC_full"),
      panel_title = pretty_title[["DCC_full"]],
      show_y_axis = TRUE,
      show_x_axis = FALSE,
      shade_episodes = shade_episodes,
      label_episodes = label_episodes,
      episodes_named = episodes_named
    )
    
    par(mar = c(0.8, left_margin_lines, 3.0, right_margin_lines))
    .draw_single_share_panel(
      z = .get_bk_share_object(bk_ceci_res, commodity_requested, direction, "sBEKK_sym"),
      panel_title = pretty_title[["sBEKK_sym"]],
      show_y_axis = FALSE,
      show_x_axis = FALSE,
      shade_episodes = shade_episodes,
      label_episodes = FALSE,
      episodes_named = episodes_named
    )
    
    par(mar = c(0.8, left_margin_lines, 3.0, right_margin_lines))
    .draw_single_share_panel(
      z = .get_bk_share_object(bk_ceci_res, commodity_requested, direction, "DCC_scalar_stage"),
      panel_title = pretty_title[["DCC_scalar_stage"]],
      show_y_axis = TRUE,
      show_x_axis = FALSE,
      shade_episodes = shade_episodes,
      label_episodes = FALSE,
      episodes_named = episodes_named
    )
    
    par(mar = c(0.8, left_margin_lines, 3.0, right_margin_lines))
    .draw_single_share_panel(
      z = .get_bk_share_object(bk_ceci_res, commodity_requested, direction, "dBEKK_sym"),
      panel_title = pretty_title[["dBEKK_sym"]],
      show_y_axis = FALSE,
      show_x_axis = FALSE,
      shade_episodes = shade_episodes,
      label_episodes = FALSE,
      episodes_named = episodes_named
    )
    
    par(mar = c(4.2, left_margin_lines, 3.0, right_margin_lines))
    .draw_single_share_panel(
      z = .get_bk_share_object(bk_ceci_res, commodity_requested, direction, "cDCC_Aielli_stage"),
      panel_title = pretty_title[["cDCC_Aielli_stage"]],
      show_y_axis = TRUE,
      show_x_axis = TRUE,
      shade_episodes = shade_episodes,
      label_episodes = FALSE,
      episodes_named = episodes_named
    )
    
    par(mar = c(4.2, left_margin_lines, 3.0, right_margin_lines))
    .draw_single_share_panel(
      z = .get_bk_share_object(bk_ceci_res, commodity_requested, direction, "dBEKK_asym"),
      panel_title = pretty_title[["dBEKK_asym"]],
      show_y_axis = FALSE,
      show_x_axis = TRUE,
      shade_episodes = shade_episodes,
      label_episodes = FALSE,
      episodes_named = episodes_named
    )
    
    par(mar = c(0, 0, 0, 0))
    plot.new()
    
    legend(
      "center",
      legend = c("Short (1-5d)", "Medium (5-20d)", "Long (20+d)"),
      fill = fill_cols,
      horiz = TRUE,
      bty = "n",
      cex = cex_leg,
      xpd = NA
    )
    
    mtext(title_text, side = 3, outer = TRUE, line = 1.4, cex = cex_main)
    mtext(subtitle_text, side = 3, outer = TRUE, line = 0.2, cex = cex_sub)
  }, width = width, height = height, res = res)
  
  invisible(outfile)
}

# ==============================================================================
# 5) LOAD vol_models AND RUN / LOAD BK CACHE
# ==============================================================================

if (!exists("vol_models")) {
  if (!file.exists(vol_models_file)) {
    stop("vol_models not found and file does not exist: ", vol_models_file)
  }
  
  vol_models <- readRDS(vol_models_file)
}

vol_models <- stats::setNames(
  lapply(names(vol_models), function(nm) {
    clean_zoo(.ensure_vol_colnames(vol_models[[nm]], nm), nm)
  }),
  names(vol_models)
)

miss <- setdiff(models_used, names(vol_models))
if (length(miss) > 0) {
  stop("vol_models is missing: ", paste(miss, collapse = ", "))
}

if (file.exists(bk_cache_file) && !isTRUE(FORCE_RECOMPUTE_BK)) {
  cat("Loading cached BK-CECI result:\n", bk_cache_file, "\n\n")
  bk_ceci_res <- readRDS(bk_cache_file)
} else {
  cat("Starting BK-CECI estimation...\n\n")
  
  bk_ceci_res <- run_bk_ceci_all(
    vol_models = vol_models,
    models_used = models_used,
    commodities = commodities4,
    use_sqrt = USE_SQRT_ON_INPUT_BK,
    input_transform = BK_INPUT_TRANSFORM,
    window = W_bk,
    nlag = nlag_bk,
    nfore = nfore_bk,
    partition = make_partition_daily(),
    no.corr = FALSE,
    scale = BK_CECI_SCALE,
    freq_labels = FREQ_LABELS
  )
  
  saveRDS(bk_ceci_res, bk_cache_file)
  cat("Saved BK-CECI cache:\n", bk_cache_file, "\n\n")
}

# ==============================================================================
# 6) PLOTS
# ==============================================================================

plot_bk_ceci_compare(
  bk_ceci_res = bk_ceci_res,
  commodity_requested = bk_plot_commodity,
  direction = bk_direction,
  benchmark = benchmark_model,
  outfile = out_bk_ceci_compare_png,
  include_dev = TRUE,
  shade_episodes = shade_episodes,
  label_episodes = label_episodes,
  episodes_named = if (shade_episodes || label_episodes) episodes_named else NULL
)

plot_bk_ceci_stacked(
  bk_ceci_res = bk_ceci_res,
  commodity_requested = bk_plot_commodity,
  model_name = benchmark_model,
  direction = bk_direction,
  outfile = out_bk_ceci_stack_png,
  shade_episodes = shade_episodes,
  label_episodes = label_episodes,
  episodes_named = if (shade_episodes || label_episodes) episodes_named else NULL
)

plot_bk_ceci_stacked_share_6panel(
  bk_ceci_res = bk_ceci_res,
  commodity_requested = bk_plot_commodity,
  direction = bk_direction,
  outfile = out_bk_ceci_share_6panel_png,
  shade_episodes = shade_episodes,
  label_episodes = label_episodes,
  episodes_named = if (shade_episodes || label_episodes) episodes_named else NULL,
  width = png_width,
  height = png_height,
  res = png_res
)

cat("\nDone.\n")
cat("BK cache RDS           :", bk_cache_file, "\n")
cat("BK compare PNG         :", out_bk_ceci_compare_png, "\n")
cat("BK benchmark stack PNG :", out_bk_ceci_stack_png, "\n")
cat("BK 6-panel share PNG   :", out_bk_ceci_share_6panel_png, "\n\n")

cat("Example access:\n")
cat('bk_ceci_res[["TTF (Gas)"]]$bidir[["DCC_full"]]\n')
cat('bk_ceci_res[["TTF (Gas)"]]$c2e[["DCC_full"]]\n')
cat('bk_ceci_res[["TTF (Gas)"]]$e2c[["DCC_full"]]\n')
################################################################################









































################################################################################
# ADD-ON — BK-CECI quantile-state heatmaps
#
# Uses existing bk_ceci_res object; no re-estimation.
#
# Outputs:
#   1) Combined one-model heatmap for all 4 commodities
#   2) Two-model comparison heatmap: DCC_full vs dBEKK_asym
################################################################################

suppressPackageStartupMessages({
  library(zoo)
  library(dplyr)
  library(ggplot2)
  library(grid)
  library(gridExtra)
  library(scales)
})

# ==============================================================================
# 1) USER SETTINGS
# ==============================================================================

bk_heatmap_commodities <- c(
  "TTF (Gas)",
  "Brent (Oil)",
  "API2 (Coal)",
  "MO1 (Carbon)"
)

bk_heatmap_model <- "DCC_full"
bk_two_models <- c("DCC_full", "dBEKK_asym")

bk_heatmap_direction <- "bidir"     # "bidir", "c2e", "e2c"
bk_heatmap_row_to_plot <- "Above 20d"  # "Above 20d" or "Below 20d"

# Use quartile bins by default. Change to seq(0, 1, by = 0.05) for 5% bins.
bk_heatmap_probs <- c(0, 0.25, 0.50, 0.75, 1)

out_bk_ceci_combined_png <- file.path(
  output_dir,
  paste0(
    "BK_CECI_CombinedHeatmap_",
    gsub("[^A-Za-z0-9]+", "", bk_heatmap_row_to_plot),
    "_",
    bk_heatmap_model,
    "_All4_",
    bk_heatmap_direction,
    ".png"
  )
)

out_bk_ceci_combined_csv <- file.path(
  output_dir,
  paste0(
    "BK_CECI_CombinedHeatmap_",
    gsub("[^A-Za-z0-9]+", "", bk_heatmap_row_to_plot),
    "_",
    bk_heatmap_model,
    "_All4_",
    bk_heatmap_direction,
    ".csv"
  )
)

out_png_bk_two_model <- file.path(
  output_dir,
  paste0(
    "BK_CECI_",
    gsub("[^A-Za-z0-9]+", "", bk_heatmap_row_to_plot),
    "_Heatmap_",
    bk_two_models[1],
    "_vs_",
    bk_two_models[2],
    "_",
    bk_heatmap_direction,
    ".png"
  )
)

out_csv_bk_two_model <- file.path(
  output_dir,
  paste0(
    "BK_CECI_",
    gsub("[^A-Za-z0-9]+", "", bk_heatmap_row_to_plot),
    "_Heatmap_",
    bk_two_models[1],
    "_vs_",
    bk_two_models[2],
    "_",
    bk_heatmap_direction,
    ".csv"
  )
)

# ==============================================================================
# 2) CHECKS
# ==============================================================================

if (!exists("bk_ceci_res")) {
  stop("bk_ceci_res not found. Run or load the BK-CECI script first.")
}

if (!exists("output_dir")) {
  output_dir <- getwd()
}

if (!exists("save_png")) {
  save_png <- function(filename, expr, width = 2600, height = 1800, res = 220) {
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
}

if (!exists(".ensure_Date_index")) {
  .ensure_Date_index <- function(z) {
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
    if (all(is.na(idx2))) stop("Could not coerce zoo index to Date.")
    zoo::index(z) <- idx2
    z
  }
}

# ==============================================================================
# 3) CORE COMPUTATION — NO RE-ESTIMATION
# ==============================================================================

compute_bk_ceci_quantile_share_table_fine <- function(bk_ceci_res,
                                                      commodity_requested,
                                                      model_name = "DCC_full",
                                                      direction = c("bidir", "c2e", "e2c"),
                                                      probs = seq(0, 1, by = 0.05)) {
  direction <- match.arg(direction)
  
  if (is.null(bk_ceci_res[[commodity_requested]])) {
    stop("Commodity not found in bk_ceci_res: ", commodity_requested)
  }
  
  if (is.null(bk_ceci_res[[commodity_requested]][[direction]])) {
    stop("Direction not found in bk_ceci_res: ", direction)
  }
  
  if (is.null(bk_ceci_res[[commodity_requested]][[direction]][[model_name]])) {
    stop(
      "Model not found in bk_ceci_res for commodity=",
      commodity_requested,
      ", direction=",
      direction,
      ", model=",
      model_name
    )
  }
  
  z <- bk_ceci_res[[commodity_requested]][[direction]][[model_name]]
  z <- .ensure_Date_index(z)
  
  cn <- colnames(z)
  cn_low <- tolower(cn)
  
  idx_short  <- grep("short",  cn_low)
  idx_medium <- grep("medium", cn_low)
  idx_long   <- grep("long",   cn_low)
  
  if (length(idx_short) != 1 || length(idx_medium) != 1 || length(idx_long) != 1) {
    stop("Could not identify Short/Medium/Long columns in BK-CECI object.")
  }
  
  z <- z[, c(idx_short, idx_medium, idx_long), drop = FALSE]
  colnames(z) <- c("Short", "Medium", "Long")
  
  X <- zoo::coredata(z)
  total <- rowSums(X, na.rm = TRUE)
  
  share_below20 <- (X[, "Short"] + X[, "Medium"]) / total
  share_above20 <- X[, "Long"] / total
  
  df <- data.frame(
    date = as.Date(zoo::index(z)),
    total = total,
    share_below20 = share_below20,
    share_above20 = share_above20,
    stringsAsFactors = FALSE
  )
  
  df <- df[
    is.finite(df$total) &
      df$total > 0 &
      is.finite(df$share_below20) &
      is.finite(df$share_above20),
    ,
    drop = FALSE
  ]
  
  if (nrow(df) == 0) {
    stop("No finite BK-CECI observations for ", commodity_requested, " / ", model_name)
  }
  
  qs <- stats::quantile(df$total, probs = probs, na.rm = TRUE, type = 7)
  
  for (i in 2:length(qs)) {
    if (qs[i] <= qs[i - 1]) {
      qs[i] <- qs[i - 1] + 1e-12
    }
  }
  
  labels <- paste0(
    format(round(100 * probs[-length(probs)], 0), trim = TRUE, scientific = FALSE),
    "-",
    format(round(100 * probs[-1], 0), trim = TRUE, scientific = FALSE)
  )
  
  df$state <- cut(
    df$total,
    breaks = qs,
    include.lowest = TRUE,
    labels = labels
  )
  
  below <- tapply(df$share_below20, df$state, mean, na.rm = TRUE)
  above <- tapply(df$share_above20, df$state, mean, na.rm = TRUE)
  
  out <- rbind(below, above)
  out <- as.matrix(out)
  rownames(out) <- c("Below 20d", "Above 20d")
  
  list(
    table = out,
    detail = df,
    breaks = qs,
    labels = labels
  )
}

# ==============================================================================
# 4) ONE-MODEL COMBINED HEATMAP
# ==============================================================================

build_bk_ceci_combined_heatmap_data <- function(bk_ceci_res,
                                                commodities = c("TTF (Gas)", "Brent (Oil)", "API2 (Coal)", "MO1 (Carbon)"),
                                                model_name = "DCC_full",
                                                direction = "bidir",
                                                row_to_plot = c("Above 20d", "Below 20d"),
                                                probs = seq(0, 1, by = 0.05)) {
  row_to_plot <- match.arg(row_to_plot)
  
  out <- lapply(commodities, function(com) {
    res <- compute_bk_ceci_quantile_share_table_fine(
      bk_ceci_res = bk_ceci_res,
      commodity_requested = com,
      model_name = model_name,
      direction = direction,
      probs = probs
    )
    
    vals <- as.numeric(100 * res$table[row_to_plot, ])
    states <- colnames(res$table)
    
    data.frame(
      commodity = com,
      state = states,
      value = vals,
      stringsAsFactors = FALSE
    )
  })
  
  df <- do.call(rbind, out)
  
  pretty_names <- c(
    "TTF (Gas)"    = "Gas",
    "Brent (Oil)"  = "Oil",
    "API2 (Coal)"  = "Coal",
    "MO1 (Carbon)" = "Carbon"
  )
  
  df$commodity <- ifelse(
    df$commodity %in% names(pretty_names),
    pretty_names[df$commodity],
    df$commodity
  )
  
  df$commodity <- factor(df$commodity, levels = c("Carbon", "Coal", "Oil", "Gas"))
  df$state <- factor(df$state, levels = unique(df$state))
  
  df
}

plot_bk_ceci_combined_heatmap <- function(bk_ceci_res,
                                          commodities = c("TTF (Gas)", "Brent (Oil)", "API2 (Coal)", "MO1 (Carbon)"),
                                          model_name = "DCC_full",
                                          direction = "bidir",
                                          row_to_plot = c("Above 20d", "Below 20d"),
                                          probs = c(0, 0.25, 0.50, 0.75, 1),
                                          outfile = NULL,
                                          main = NULL,
                                          digits = 1,
                                          show_legend = FALSE) {
  row_to_plot <- match.arg(row_to_plot)
  
  df <- build_bk_ceci_combined_heatmap_data(
    bk_ceci_res = bk_ceci_res,
    commodities = commodities,
    model_name = model_name,
    direction = direction,
    row_to_plot = row_to_plot,
    probs = probs
  )
  
  vmin <- min(df$value, na.rm = TRUE)
  vmax <- max(df$value, na.rm = TRUE)
  cut_text <- vmin + 0.62 * (vmax - vmin)
  df$text_col <- ifelse(df$value >= cut_text, "white", "black")
  
  if (is.null(main)) {
    dlab <- switch(
      direction,
      bidir = "Bidirectional BK-CECI",
      c2e   = "Commodity -> Equity BK-CECI",
      e2c   = "Equity -> Commodity BK-CECI"
    )
    
    main <- paste0(
      model_name,
      ": ",
      row_to_plot,
      " share by BK-CECI quantile state\n(",
      dlab,
      ")"
    )
  }
  
  p <- ggplot(df, aes(x = state, y = commodity, fill = value)) +
    geom_tile(color = "white", linewidth = 0.55, width = 0.98, height = 0.98) +
    geom_text(
      aes(
        label = sprintf(paste0("%.", digits, "f%%"), value),
        color = text_col
      ),
      size = 5.0,
      fontface = "plain",
      show.legend = FALSE
    ) +
    scale_color_identity() +
    scale_fill_gradientn(
      colours = c("#fff5f0", "#fcbba1", "#fb6a4a", "#cb181d", "#67000d"),
      limits = c(vmin, vmax),
      name = row_to_plot
    ) +
    labs(
      title = main,
      x = "Total BK-CECI quantile state (%)",
      y = "Commodity"
    ) +
    coord_fixed(ratio = 0.28) +
    theme_minimal(base_size = 16) +
    theme(
      panel.grid = element_blank(),
      axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 13),
      axis.text.y = element_text(size = 14),
      axis.title.x = element_text(face = "bold", size = 16),
      axis.title.y = element_text(face = "bold", size = 16),
      plot.title = element_text(face = "bold", hjust = 0.5, size = 18),
      legend.position = if (show_legend) "right" else "none",
      plot.margin = margin(8, 14, 8, 8)
    )
  
  if (!is.null(outfile)) {
    ggsave(outfile, plot = p, width = 13.5, height = 4.6, dpi = 220)
  }
  
  invisible(list(plot = p, data = df))
}

# ==============================================================================
# 5) TWO-MODEL HEATMAP
# ==============================================================================

get_legend_grob_bk_quantile <- function(p) {
  g <- ggplotGrob(p)
  guides <- which(sapply(g$grobs, function(x) x$name) == "guide-box")
  if (length(guides) == 0) return(grid::nullGrob())
  g$grobs[[guides[1]]]
}

build_bk_ceci_heatmap_df_models <- function(bk_ceci_res,
                                            commodities = c("TTF (Gas)", "Brent (Oil)", "API2 (Coal)", "MO1 (Carbon)"),
                                            models = c("DCC_full", "dBEKK_asym"),
                                            direction = "bidir",
                                            row_to_plot = c("Above 20d", "Below 20d"),
                                            probs = c(0, 0.25, 0.50, 0.75, 1)) {
  row_to_plot <- match.arg(row_to_plot)
  
  out <- list()
  
  for (m in models) {
    for (com in commodities) {
      res <- compute_bk_ceci_quantile_share_table_fine(
        bk_ceci_res = bk_ceci_res,
        commodity_requested = com,
        model_name = m,
        direction = direction,
        probs = probs
      )
      
      vals <- as.numeric(100 * res$table[row_to_plot, ])
      states <- colnames(res$table)
      
      tmp <- data.frame(
        model = m,
        commodity = com,
        tail_bucket = states,
        heat_value = vals,
        stringsAsFactors = FALSE
      )
      
      out[[paste(m, com, sep = "_")]] <- tmp
    }
  }
  
  df <- dplyr::bind_rows(out)
  
  pretty_names <- c(
    "TTF (Gas)"    = "Gas",
    "Brent (Oil)"  = "Oil",
    "API2 (Coal)"  = "Coal",
    "MO1 (Carbon)" = "Carbon"
  )
  
  df$commodity_label <- ifelse(
    df$commodity %in% names(pretty_names),
    pretty_names[df$commodity],
    df$commodity
  )
  
  df$commodity_label <- factor(df$commodity_label, levels = c("Carbon", "Coal", "Oil", "Gas"))
  df$tail_bucket <- factor(df$tail_bucket, levels = unique(df$tail_bucket))
  
  df
}

plot_bk_ceci_two_model_heatmap <- function(bk_ceci_res,
                                           out_png_file,
                                           commodities = c("TTF (Gas)", "Brent (Oil)", "API2 (Coal)", "MO1 (Carbon)"),
                                           models = c("DCC_full", "dBEKK_asym"),
                                           direction = "bidir",
                                           row_to_plot = "Above 20d",
                                           probs = c(0, 0.25, 0.50, 0.75, 1),
                                           png_width = 3200,
                                           png_height = 1500,
                                           png_res = 220) {
  
  heatmap_df <- build_bk_ceci_heatmap_df_models(
    bk_ceci_res = bk_ceci_res,
    commodities = commodities,
    models = models,
    direction = direction,
    row_to_plot = row_to_plot,
    probs = probs
  )
  
  fill_lim_low <- min(heatmap_df$heat_value, na.rm = TRUE)
  fill_lim_high <- max(heatmap_df$heat_value, na.rm = TRUE)
  
  plot_df_left <- heatmap_df |> dplyr::filter(model == models[1])
  plot_df_right <- heatmap_df |> dplyr::filter(model == models[2])
  
  base_heatmap <- function(df_plot,
                           panel_title = "",
                           show_y_title = TRUE,
                           hide_y_text_visually = FALSE,
                           show_legend = TRUE) {
    
    cut_text <- fill_lim_low + 0.62 * (fill_lim_high - fill_lim_low)
    
    # Dark-red high values need white text; light low values need black text.
    df_plot$text_col <- ifelse(df_plot$heat_value >= cut_text, "white", "black")
    
    p <- ggplot(
      df_plot,
      aes(x = tail_bucket, y = commodity_label, fill = heat_value)
    ) +
      geom_tile(color = "white", linewidth = 0.8, width = 0.98, height = 0.98) +
      geom_text(
        aes(
          label = ifelse(is.na(heat_value), "", sprintf("%.1f%%", heat_value)),
          color = text_col
        ),
        size = 5
      ) +
      scale_color_identity() +
      scale_fill_gradientn(
        colours = c("#fff5f0", "#fcbba1", "#fb6a4a", "#cb181d", "#67000d"),
        limits = c(fill_lim_low, fill_lim_high),
        oob = scales::squish,
        name = expression(`+20d` ~ "(%)")
      ) +
      labs(
        title = panel_title,
        x = "BK-CECI Quantile State",
        y = if (show_y_title) "Commodity" else NULL
      ) +
      coord_fixed(ratio = 1) +
      theme_minimal(base_size = 16) +
      theme(
        panel.grid = element_blank(),
        axis.text.x = element_text(size = 14, angle = 0, vjust = 0.5, hjust = 1),
        axis.title.x = element_text(size = 16),
        plot.title = element_text(size = 18, face = "bold", hjust = 0.5),
        legend.position = if (show_legend) "right" else "none",
        legend.title = element_text(size = 14),
        legend.text = element_text(size = 12)
      )
    
    if (!hide_y_text_visually) {
      p <- p +
        theme(
          axis.text.y = element_text(size = 14, angle = 0),
          axis.title.y = if (show_y_title) element_text(size = 16) else element_blank(),
          axis.ticks.y = element_line()
        )
    } else {
      p <- p +
        theme(
          axis.text.y = element_text(size = 14, angle = 0, colour = scales::alpha("black", 0)),
          axis.title.y = element_blank(),
          axis.ticks.y = element_line(colour = scales::alpha("black", 0))
        )
    }
    
    p
  }
  
  p_left <- base_heatmap(
    df_plot = plot_df_left,
    panel_title = models[1],
    show_y_title = TRUE,
    hide_y_text_visually = FALSE,
    show_legend = FALSE
  )
  
  p_right <- base_heatmap(
    df_plot = plot_df_right,
    panel_title = models[2],
    show_y_title = FALSE,
    hide_y_text_visually = TRUE,
    show_legend = TRUE
  )
  
  legend_grob <- get_legend_grob_bk_quantile(p_right)
  p_right_nolegend <- p_right + theme(legend.position = "none")
  
  title_grob <- grid::textGrob(
    "BK-CECI horizon composition by quantile state",
    gp = grid::gpar(fontsize = 22, fontface = "bold")
  )
  
  subtitle_grob <- grid::textGrob(
    paste0(
      row_to_plot,
      " share across BK-CECI quantile states; direction = ",
      direction,
      "; quantile bins = ",
      paste0(100 * probs[-length(probs)], "-", 100 * probs[-1], collapse = ", ")
    ),
    gp = grid::gpar(fontsize = 14)
  )
  
  note_grob <- grid::textGrob(
    paste0(
      "Cells show mean ",
      row_to_plot,
      " share conditional on the total BK-CECI state. ",
      "Rows are commodities; columns are BK-CECI quantile bins."
    ),
    gp = grid::gpar(fontsize = 12)
  )
  
  combined_plots <- gridExtra::arrangeGrob(
    grobs = list(p_left, p_right_nolegend, legend_grob),
    ncol = 3,
    widths = c(1, 1, 0.12)
  )
  
  final_plot <- gridExtra::arrangeGrob(
    title_grob,
    subtitle_grob,
    combined_plots,
    note_grob,
    ncol = 1,
    heights = c(0.09, 0.06, 0.79, 0.06)
  )
  
  save_png(
    out_png_file,
    {
      grid::grid.newpage()
      grid::grid.draw(final_plot)
    },
    width = png_width,
    height = png_height,
    res = png_res
  )
  
  invisible(heatmap_df)
}

# ==============================================================================
# 6) RUN — NO RE-ESTIMATION
# ==============================================================================

cat("\n==================================================================\n")
cat("BK-CECI QUANTILE-STATE HEATMAPS — USING EXISTING bk_ceci_res\n")
cat("No BK estimation is performed in this add-on.\n")
cat("==================================================================\n\n")

res_combined <- plot_bk_ceci_combined_heatmap(
  bk_ceci_res = bk_ceci_res,
  commodities = bk_heatmap_commodities,
  model_name = bk_heatmap_model,
  direction = bk_heatmap_direction,
  row_to_plot = bk_heatmap_row_to_plot,
  probs = bk_heatmap_probs,
  outfile = out_bk_ceci_combined_png
)

utils::write.csv(
  res_combined$data,
  out_bk_ceci_combined_csv,
  row.names = FALSE
)

heatmap_df_bk <- plot_bk_ceci_two_model_heatmap(
  bk_ceci_res = bk_ceci_res,
  out_png_file = out_png_bk_two_model,
  commodities = bk_heatmap_commodities,
  models = bk_two_models,
  direction = bk_heatmap_direction,
  row_to_plot = bk_heatmap_row_to_plot,
  probs = bk_heatmap_probs,
  png_width = 3200,
  png_height = 1500,
  png_res = 220
)

utils::write.csv(
  heatmap_df_bk,
  out_csv_bk_two_model,
  row.names = FALSE
)

cat("Saved BK combined heatmap PNG :", out_bk_ceci_combined_png, "\n")
cat("Saved BK combined heatmap CSV :", out_bk_ceci_combined_csv, "\n")
cat("Saved BK two-model heatmap PNG:", out_png_bk_two_model, "\n")
cat("Saved BK two-model heatmap CSV:", out_csv_bk_two_model, "\n\n")