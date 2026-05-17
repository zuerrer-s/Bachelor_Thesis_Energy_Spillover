################################################################################
# Pairwise SX NET CECI networks
#
# Uses exactly the requested structure:
#   - commodities = API2 (Coal)
#   - models = sBEKK_sym, dBEKK_sym, dBEKK_asym, DCC_full
#   - HIGH regime = top 25% commodity volatility days from DCC_full levels
#   - outputs model-level summary, grouped BEKK_mean vs DCC summary, and JPG networks
#
# INPUT REQUIRED
#   - vol_models in memory, OR vol_models_all_7models.rds in base/output folder
################################################################################

suppressPackageStartupMessages({
  library(zoo)
  library(ConnectednessApproach)
  library(igraph)
})

# ------------------------------ LOAD INPUT -----------------------------------

base_dir <- "C:/Users/sezue/OneDrive/Desktop/BA"

output_dir <- file.path(base_dir, "outputs")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

vol_models_file <- file.path(output_dir, "vol_models_all_7models.rds")
if (!file.exists(vol_models_file)) {
  vol_models_file <- file.path(base_dir, "vol_models_all_7models.rds")
}

if (!exists("vol_models")) {
  if (!file.exists(vol_models_file)) {
    stop("vol_models not found in memory and file does not exist: ", vol_models_file)
  }
  vol_models <- readRDS(vol_models_file)
}

# ------------------------------ USER SETTINGS --------------------------------

commodities <- c("API2 (Coal)")

models_used <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym",
  "DCC_full"
)

# Rolling connectedness settings
W     <- 200
nlag  <- 1
nfore <- 20

# Equity identification in your asset names (before "_vol")
equity_regex <- "^SX"

# Input transform from vols:
# "dlog_var" = Δlog(σ^2) ; "dlog_vol" = Δlog(σ)
input_transform <- "dlog_var"

# Regime definition calendar
regime_model  <- "DCC_full"
regime_q_high <- 0.75   # top 25% commodity vol days (levels, not dlog)

# Outputs
out_csv_ts            <- file.path(output_dir, "CECI_pairwise_SX_NET_timeseries.csv")
out_csv_summ          <- file.path(output_dir, "CECI_pairwise_SX_NET_summary_ALL_vs_HIGH.csv")
out_csv_summ_grouped  <- file.path(output_dir, "CECI_pairwise_SX_NET_summary_grouped_ALL_vs_HIGH.csv")

# Network plot settings
out_dir    <- file.path(output_dir, "NET_network_jpg")
edge_top_k <- 12                 # show strongest |mean NET| edges
jpeg_w     <- 2400
jpeg_h     <- 1600
jpeg_res   <- 240

# IMPORTANT: fixed global scaling for comparability across plots (ALL vs HIGH)
EDGE_MAX_ABS   <- 2.0            # |NET| >= 2 -> max thickness
EDGE_MIN_WIDTH <- 1.2
EDGE_MAX_WIDTH <- 9.0

# --- Arrowhead visibility on thick edges ---
EDGE_ARROW_SIZE_BASE  <- 1.10
EDGE_ARROW_WIDTH_MULT <- 0.38
EDGE_ARROW_WIDTH_MIN  <- 1.25

# --- Label placement: rotate POSITION slightly around center, keep text horizontal ---
EDGE_LABEL_AROUND_CENTER_DEG <- 12
EDGE_LABEL_ROT_BLEND         <- 0.57
EDGE_LABEL_POS_FRAC          <- 0.56

# shift all arrow numbers slightly to the RIGHT (screen/right axis)
EDGE_LABEL_SHIFT_X <- 0.001

# make yellow bubbles bigger so text fits
EQUITY_NODE_SIZE_MULT <- 1.15

# Make sure bubbles are not clipped
PLOT_MARGIN_MULT <- 1.22

# ------------------------------ HELPERS --------------------------------------

.dir_create <- function(d) if (!dir.exists(d)) dir.create(d, recursive=TRUE, showWarnings=FALSE)

.strip_vol <- function(x) sub("_vol$", "", x)

.norm_name <- function(x) {
  x <- tolower(x)
  gsub("[^a-z0-9]+", "", x)
}

.match_one <- function(requested, available_base) {
  req_n <- .norm_name(requested)
  av_n  <- .norm_name(available_base)
  
  hit <- which(av_n == req_n)
  if (length(hit) >= 1) return(available_base[hit[1]])
  
  hit <- which(vapply(av_n, function(a) grepl(req_n, a, fixed=TRUE), logical(1)))
  if (length(hit) >= 1) return(available_base[hit[1]])
  
  hit <- which(vapply(av_n, function(a) grepl(a, req_n, fixed=TRUE), logical(1)))
  if (length(hit) >= 1) return(available_base[hit[1]])
  
  NA_character_
}

.ensure_vol_colnames <- function(vol_zoo, model_name="model") {
  stopifnot(inherits(vol_zoo, "zoo"))
  cn <- colnames(vol_zoo)
  if (is.null(cn)) stop(model_name, ": vol_zoo has no colnames().")
  cn2 <- ifelse(grepl("_vol$", cn), cn, paste0(cn, "_vol"))
  colnames(vol_zoo) <- cn2
  vol_zoo
}

.ensure_Date_index <- function(z) {
  stopifnot(inherits(z, "zoo"))
  idx <- index(z)
  
  if (inherits(idx, "Date")) {
    # still remove duplicate Date entries if any
    if (anyDuplicated(idx)) {
      warning("Duplicate Date index entries detected. Keeping last row per Date.")
      z <- z[!duplicated(idx, fromLast = TRUE), ]
    }
    return(z)
  }
  
  if (inherits(idx, "POSIXct") || inherits(idx, "POSIXt")) {
    index(z) <- as.Date(idx)
  } else if (inherits(idx, "yearmon") || inherits(idx, "yearqtr")) {
    index(z) <- as.Date(idx)
  } else if (is.numeric(idx)) {
    index(z) <- as.Date(idx, origin="1970-01-01")
  } else {
    try_idx <- suppressWarnings(as.Date(idx))
    if (all(is.na(try_idx))) stop("Could not coerce zoo index to Date.")
    index(z) <- try_idx
  }
  
  if (anyDuplicated(index(z))) {
    warning("Duplicate Date index entries detected after Date coercion. Keeping last row per Date.")
    z <- z[!duplicated(index(z), fromLast = TRUE), ]
  }
  
  z
}

.build_v_from_sigma <- function(sig_zoo, transform=c("dlog_var","dlog_vol")) {
  transform <- match.arg(transform)
  sig_zoo <- .ensure_Date_index(sig_zoo)
  
  S <- coredata(sig_zoo)
  if (!is.matrix(S)) S <- as.matrix(S)
  S <- pmax(S, 1e-12)
  
  if (transform == "dlog_var") {
    H <- S^2
    V <- diff(log(H))
  } else {
    V <- diff(log(S))
  }
  
  out <- zoo(V, order.by = index(sig_zoo)[-1])
  colnames(out) <- colnames(sig_zoo)
  .ensure_Date_index(out)
}

.get_theta_cube <- function(ca, v_in) {
  cand <- c("CT","GFEVD","FEVD","Theta","theta","TABLE")
  for (nm in cand) {
    obj <- ca[[nm]]
    
    if (is.array(obj) && length(dim(obj)) == 3 && is.numeric(obj)) {
      theta <- obj
      if (is.null(dimnames(theta)) || is.null(dimnames(theta)[[1]])) {
        dn <- colnames(v_in)
        if (is.null(dn)) stop("v_in has no colnames(); cannot attach dimnames to theta.")
        dimnames(theta) <- list(dn, dn, NULL)
      }
      return(theta)
    }
    
    if (is.list(obj)) {
      inner_cand <- c("CT","GFEVD","FEVD","Theta","theta")
      for (jj in inner_cand) {
        if (!is.null(obj[[jj]]) &&
            is.array(obj[[jj]]) &&
            length(dim(obj[[jj]])) == 3 &&
            is.numeric(obj[[jj]])) {
          theta <- obj[[jj]]
          if (is.null(dimnames(theta)) || is.null(dimnames(theta)[[1]])) {
            dn <- colnames(v_in)
            if (is.null(dn)) stop("v_in has no colnames(); cannot attach dimnames to theta.")
            dimnames(theta) <- list(dn, dn, NULL)
          }
          return(theta)
        }
      }
    }
  }
  stop("No 3D FEVD/GFEVD cube found. Available names: ", paste(names(ca), collapse=", "))
}

.pairwise_from_theta <- function(theta, com_asset_vol, eq_assets_vol) {
  asset_names <- dimnames(theta)[[1]]
  if (is.null(asset_names)) stop("theta has no asset dimnames().")
  
  com_idx <- match(com_asset_vol, asset_names)
  if (is.na(com_idx)) stop("Commodity asset not found in theta: ", com_asset_vol)
  
  eq_idx <- match(eq_assets_vol, asset_names)
  if (any(is.na(eq_idx))) stop("Some equities not in theta: ", paste(eq_assets_vol[is.na(eq_idx)], collapse=", "))
  
  TT <- dim(theta)[3]
  K  <- length(eq_idx)
  
  E_to_C <- matrix(NA_real_, nrow=TT, ncol=K, dimnames=list(NULL, .strip_vol(eq_assets_vol)))
  C_to_E <- matrix(NA_real_, nrow=TT, ncol=K, dimnames=list(NULL, .strip_vol(eq_assets_vol)))
  
  for (t in seq_len(TT)) {
    M <- theta[,,t]
    E_to_C[t,] <- 100 * as.numeric(M[eq_idx, com_idx])
    C_to_E[t,] <- 100 * as.numeric(M[com_idx, eq_idx])
  }
  
  NET <- C_to_E - E_to_C
  list(E_to_C=E_to_C, C_to_E=C_to_E, NET=NET)
}

.get_high_regime_dates <- function(sig_regime, commodity_base, q_high=0.90) {
  sig_regime <- .ensure_Date_index(sig_regime)
  base <- .strip_vol(colnames(sig_regime))
  com_hit <- .match_one(commodity_base, base)
  if (is.na(com_hit)) stop("Commodity not found in regime model vols: ", commodity_base)
  
  z <- sig_regime[, paste0(com_hit, "_vol")]
  x <- as.numeric(coredata(z))
  thr <- as.numeric(stats::quantile(x, q_high, na.rm=TRUE, names=FALSE))
  dates <- index(z)[is.finite(x) & x >= thr]
  list(threshold=thr, dates=dates)
}

.summ_vec <- function(x) {
  x <- x[is.finite(x)]
  if (length(x) == 0) return(c(mean=NA_real_, sd=NA_real_, N=0))
  c(mean=mean(x), sd=stats::sd(x), N=length(x))
}

.pretty_label <- function(name_base) {
  if (grepl("\\(", name_base)) {
    main <- trimws(sub("\\(.*$", "", name_base))
    br   <- regmatches(name_base, regexpr("\\([^\\)]+\\)", name_base))
    paste0(main, "\n", br)
  } else {
    name_base
  }
}

.scale_width_fixed <- function(w_abs) {
  w_abs <- pmin(pmax(w_abs, 0), EDGE_MAX_ABS)
  EDGE_MIN_WIDTH + (w_abs / EDGE_MAX_ABS) * (EDGE_MAX_WIDTH - EDGE_MIN_WIDTH)
}

.plot_limits_from_layout <- function(lay, mult=PLOT_MARGIN_MULT) {
  xr <- range(lay[,1], finite=TRUE)
  yr <- range(lay[,2], finite=TRUE)
  cx <- mean(xr); cy <- mean(yr)
  hx <- diff(xr) / 2; hy <- diff(yr) / 2
  hx <- ifelse(hx <= 0, 1, hx)
  hy <- ifelse(hy <= 0, 1, hy)
  xlim <- c(cx - mult*hx, cx + mult*hx)
  ylim <- c(cy - mult*hy, cy + mult*hy)
  list(xlim=xlim, ylim=ylim)
}

.draw_edge_labels_rotate_around_center <- function(g, layout_xy, labels,
                                                   center_name,
                                                   pos=EDGE_LABEL_POS_FRAC,
                                                   rot_deg=EDGE_LABEL_AROUND_CENTER_DEG,
                                                   blend=EDGE_LABEL_ROT_BLEND,
                                                   shift_x=EDGE_LABEL_SHIFT_X,
                                                   cex=0.85,
                                                   col="black") {
  if (length(labels) == 0 || igraph::ecount(g) == 0) return(invisible(NULL))
  
  vnames <- igraph::V(g)$name
  cidx <- match(center_name, vnames)
  if (is.na(cidx)) return(invisible(NULL))
  cx <- layout_xy[cidx, 1]; cy <- layout_xy[cidx, 2]
  
  el <- igraph::ends(g, igraph::E(g), names=TRUE)
  idx_from <- match(el[,1], vnames)
  idx_to   <- match(el[,2], vnames)
  if (anyNA(idx_from) || anyNA(idx_to)) return(invisible(NULL))
  
  x1 <- layout_xy[idx_from, 1]; y1 <- layout_xy[idx_from, 2]
  x2 <- layout_xy[idx_to,   1]; y2 <- layout_xy[idx_to,   2]
  
  x0 <- (1 - pos) * x1 + pos * x2
  y0 <- (1 - pos) * y1 + pos * y2
  
  th <- rot_deg * pi / 180
  xr <- x0 - cx; yr <- y0 - cy
  xR <- cx + (xr * cos(th) - yr * sin(th))
  yR <- cy + (xr * sin(th) + yr * cos(th))
  
  b <- pmin(pmax(blend, 0), 1)
  xF <- (1 - b) * x0 + b * xR
  yF <- (1 - b) * y0 + b * yR
  
  xF <- xF + shift_x
  
  graphics::text(xF, yF, labels=labels, srt=0, cex=cex, col=col, xpd=NA)
  invisible(NULL)
}

# Deterministic (fixed) star layout: equities placed in a stable alphabetical order
.layout_star_fixed <- function(g, center_name) {
  vnames <- igraph::V(g)$name
  cidx <- match(center_name, vnames)
  if (is.na(cidx)) stop("Center node not found in graph vertices: ", center_name)
  
  others <- vnames[vnames != center_name]
  others_sorted <- sort(others)
  n <- length(others_sorted)
  
  lay <- matrix(0, nrow=length(vnames), ncol=2, dimnames=list(vnames, c("x","y")))
  lay[center_name,] <- c(0, 0)
  
  if (n > 0) {
    ang <- seq(from=pi/2, length.out=n, by=2*pi/n)
    r <- 1.0
    lay[others_sorted, 1] <- r * cos(ang)
    lay[others_sorted, 2] <- r * sin(ang)
  }
  
  lay[vnames, , drop=FALSE]
}

# Build plotting means:
# - DCC_full uses just DCC_full
# - BEKK_mean averages sBEKK_sym, dBEKK_sym, dBEKK_asym by Date x Equity first,
#   then averages over time for the plot
.get_plot_means <- function(ts_df, commodity_requested, model_choice, high_dates=NULL) {
  
  if (model_choice == "BEKK_mean") {
    sub <- ts_df[
      ts_df$Model %in% c("sBEKK_sym", "dBEKK_sym", "dBEKK_asym") &
        ts_df$Commodity == commodity_requested,
      ,
      drop = FALSE
    ]
    
    sub2 <- aggregate(NET ~ Date + Equity, data=sub, FUN=function(x) mean(x, na.rm=TRUE))
    
  } else if (model_choice == "DCC_full") {
    sub2 <- ts_df[
      ts_df$Model == "DCC_full" &
        ts_df$Commodity == commodity_requested,
      c("Date", "Equity", "NET"),
      drop = FALSE
    ]
    
  } else {
    stop("Unsupported model_choice: ", model_choice)
  }
  
  mean_all <- aggregate(NET ~ Equity, data=sub2, FUN=function(x) mean(x, na.rm=TRUE))
  colnames(mean_all) <- c("Equity", "meanNET")
  
  sub_h <- sub2[as.Date(sub2$Date) %in% as.Date(high_dates), , drop=FALSE]
  mean_high <- aggregate(NET ~ Equity, data=sub_h, FUN=function(x) mean(x, na.rm=TRUE))
  colnames(mean_high) <- c("Equity", "meanNET")
  
  list(mean_all=mean_all, mean_high=mean_high)
}

# ------------------------------ PLOT -----------------------------------------

.plot_net_network <- function(df_means, commodity_label, main_title, top_k=12) {
  
  d <- df_means[is.finite(df_means$meanNET), , drop=FALSE]
  if (nrow(d) == 0) {
    plot.new()
    title(main=paste0(main_title, "\n(no finite values)"))
    return(invisible(NULL))
  }
  
  d$absw <- abs(d$meanNET)
  d <- d[order(d$absw, decreasing=TRUE), , drop=FALSE]
  if (!is.null(top_k) && nrow(d) > top_k) d <- d[1:top_k, , drop=FALSE]
  
  edges <- data.frame(from=character(0), to=character(0),
                      w=numeric(0), dir=character(0),
                      stringsAsFactors=FALSE)
  
  for (i in seq_len(nrow(d))) {
    eq <- d$Equity[i]
    w  <- d$meanNET[i]
    if (!is.finite(w) || w == 0) next
    
    if (w > 0) edges <- rbind(edges, data.frame(from=commodity_label, to=eq, w=w, dir="OUT"))
    if (w < 0) edges <- rbind(edges, data.frame(from=eq, to=commodity_label, w=abs(w), dir="IN"))
  }
  
  if (nrow(edges) == 0) {
    plot.new()
    title(main=paste0(main_title, "\n(no nonzero edges)"))
    return(invisible(NULL))
  }
  
  nodes <- unique(c(edges$from, edges$to))
  g <- igraph::graph_from_data_frame(edges, directed=TRUE, vertices=data.frame(name=nodes))
  
  igraph::V(g)$size <- ifelse(
    igraph::V(g)$name == commodity_label,
    56,
    38 * EQUITY_NODE_SIZE_MULT
  )
  igraph::V(g)$color <- ifelse(igraph::V(g)$name == commodity_label, "#7FB3D5", "#F4D03F")
  igraph::V(g)$label <- vapply(igraph::V(g)$name, .pretty_label, character(1))
  igraph::V(g)$label.cex <- 1.05
  igraph::V(g)$label.color <- "black"
  
  igraph::E(g)$color <- ifelse(igraph::E(g)$dir == "OUT", "#1F78B4", "#D73027")
  igraph::E(g)$width <- .scale_width_fixed(igraph::E(g)$w)
  
  igraph::E(g)$arrow.size  <- EDGE_ARROW_SIZE_BASE
  igraph::E(g)$arrow.width <- pmax(EDGE_ARROW_WIDTH_MIN, EDGE_ARROW_WIDTH_MULT * igraph::E(g)$width)
  
  edge_num_labels <- format(round(igraph::E(g)$w, 3), nsmall=3)
  igraph::E(g)$label <- NA_character_
  
  lay0 <- .layout_star_fixed(g, center_name=commodity_label)
  lay  <- igraph::norm_coords(lay0, xmin=-1, xmax=1, ymin=-1, ymax=1)
  
  lims <- .plot_limits_from_layout(lay, mult=PLOT_MARGIN_MULT)
  
  op <- par(mar=c(1.3, 1.3, 3.5, 1.3), xpd=NA)
  on.exit(par(op), add=TRUE)
  
  plot(g,
       layout=lay,
       rescale=FALSE,
       xlim=lims$xlim,
       ylim=lims$ylim,
       main=main_title,
       edge.label=NA)
  
  .draw_edge_labels_rotate_around_center(
    g=g,
    layout_xy=lay,
    labels=edge_num_labels,
    center_name=commodity_label,
    pos=EDGE_LABEL_POS_FRAC,
    rot_deg=EDGE_LABEL_AROUND_CENTER_DEG,
    blend=EDGE_LABEL_ROT_BLEND,
    shift_x=EDGE_LABEL_SHIFT_X,
    cex=0.85,
    col="black"
  )
  
  invisible(g)
}

.save_jpeg_plot <- function(filepath, expr) {
  grDevices::jpeg(filename=filepath, width=jpeg_w, height=jpeg_h, res=jpeg_res, quality=96)
  tryCatch(eval.parent(substitute(expr)),
           error=function(e) {
             plot.new()
             text(0.01, 0.99, paste("ERROR:", conditionMessage(e)), adj=c(0,1), cex=0.9)
           },
           finally=grDevices::dev.off())
  invisible(filepath)
}

# ------------------------------ MAIN -----------------------------------------

cat("START\n")

if (!exists("vol_models")) stop("vol_models not found. Load/build vol_models first.")
miss <- setdiff(unique(c(models_used, regime_model)), names(vol_models))
if (length(miss) > 0) stop("vol_models is missing: ", paste(miss, collapse=", "))

vol_models <- stats::setNames(
  lapply(names(vol_models), function(nm) {
    .ensure_Date_index(.ensure_vol_colnames(vol_models[[nm]], nm))
  }),
  names(vol_models)
)

vol_models[[regime_model]] <- .ensure_Date_index(.ensure_vol_colnames(vol_models[[regime_model]], regime_model))
sig_regime <- vol_models[[regime_model]]

high_regime <- setNames(vector("list", length(commodities)), commodities)
for (cname in commodities) {
  high_regime[[cname]] <- .get_high_regime_dates(sig_regime, cname, q_high=regime_q_high)
  cat("Regime calendar:", cname,
      "| thr=", round(high_regime[[cname]]$threshold, 6),
      "| days=", length(high_regime[[cname]]$dates), "\n")
}

rows_ts <- list()

for (m in models_used) {
  cat("\nModel:", m, "\n")
  
  sig <- .ensure_Date_index(.ensure_vol_colnames(vol_models[[m]], m))
  base <- .strip_vol(colnames(sig))
  
  eq_base <- base[grep(equity_regex, base)]
  if (length(eq_base) == 0) stop("No equities matched equity_regex='", equity_regex, "' in model ", m)
  eq_assets_vol <- paste0(eq_base, "_vol")
  
  for (commodity_requested in commodities) {
    cat("  Commodity:", commodity_requested, "\n")
    
    com_base_hit <- .match_one(commodity_requested, base)
    if (is.na(com_base_hit)) stop("Commodity '", commodity_requested, "' not found in model ", m)
    
    use_base <- unique(c(eq_base, com_base_hit))
    sig_use  <- sig[, paste0(use_base, "_vol"), drop=FALSE]
    sig_use  <- sig_use[complete.cases(sig_use), , drop=FALSE]
    
    v_in <- .ensure_Date_index(.build_v_from_sigma(sig_use, transform=input_transform))
    
    if (NROW(v_in) <= W + nlag + 5) {
      stop("Too few observations for ", m, " / ", commodity_requested,
           ". N=", NROW(v_in), ", W=", W, ", nlag=", nlag)
    }
    
    ca <- ConnectednessApproach(
      x = v_in,
      nlag = nlag,
      nfore = nfore,
      window.size = W,
      model = "VAR",
      connectedness = "Time"
    )
    
    theta <- .get_theta_cube(ca, v_in)
    
    com_asset_vol <- paste0(com_base_hit, "_vol")
    pw <- .pairwise_from_theta(theta, com_asset_vol, eq_assets_vol)
    
    TT <- nrow(pw$NET)
    end_dates <- tail(index(v_in), TT)
    
    for (k in seq_along(eq_base)) {
      rows_ts[[length(rows_ts)+1]] <- data.frame(
        Date      = as.character(end_dates),
        Model     = m,
        Commodity = commodity_requested,
        Equity    = eq_base[k],
        E_to_C    = pw$E_to_C[, k],
        C_to_E    = pw$C_to_E[, k],
        NET       = pw$NET[,   k],
        stringsAsFactors = FALSE
      )
    }
  }
}

ts_df <- do.call(rbind, rows_ts)
ts_df$E_to_C <- as.numeric(ts_df$E_to_C)
ts_df$C_to_E <- as.numeric(ts_df$C_to_E)
ts_df$NET    <- as.numeric(ts_df$NET)

ts_df$ModelGroup <- ifelse(
  ts_df$Model %in% c("sBEKK_sym", "dBEKK_sym", "dBEKK_asym"),
  "BEKK",
  ifelse(ts_df$Model == "DCC_full", "DCC", NA_character_)
)

write.csv(ts_df, out_csv_ts, row.names=FALSE)
cat("\nSaved time series:", out_csv_ts, "\n")

# ------------------------------ MODEL-LEVEL SUMMARY --------------------------

keys <- unique(ts_df[, c("Model","Commodity","Equity")])
summ_rows <- vector("list", nrow(keys))

for (i in seq_len(nrow(keys))) {
  ki <- keys[i, ]
  sub <- ts_df[ts_df$Model==ki$Model & ts_df$Commodity==ki$Commodity & ts_df$Equity==ki$Equity, ]
  
  s_all <- .summ_vec(sub$NET)
  
  hd <- high_regime[[ki$Commodity]]$dates
  sub_h <- sub[as.Date(sub$Date) %in% as.Date(hd), , drop=FALSE]
  s_high <- .summ_vec(sub_h$NET)
  
  summ_rows[[i]] <- data.frame(
    Model=ki$Model, Commodity=ki$Commodity, Equity=ki$Equity,
    NET_mean_all=s_all["mean"], NET_sd_all=s_all["sd"], N_all=s_all["N"],
    NET_mean_high=s_high["mean"], NET_sd_high=s_high["sd"], N_high=s_high["N"],
    stringsAsFactors = FALSE
  )
}

summ_df <- do.call(rbind, summ_rows)
write.csv(summ_df, out_csv_summ, row.names=FALSE)
cat("Saved summary:", out_csv_summ, "\n")

# ------------------------------ GROUPED SUMMARY ------------------------------
# BEKK_mean = date-by-date mean across the three BEKK models
# DCC = DCC_full only

group_keys <- unique(ts_df[, c("ModelGroup","Commodity","Equity")])
group_summ_rows <- vector("list", nrow(group_keys))

for (i in seq_len(nrow(group_keys))) {
  ki <- group_keys[i, ]
  
  sub <- ts_df[
    ts_df$ModelGroup == ki$ModelGroup &
      ts_df$Commodity == ki$Commodity &
      ts_df$Equity == ki$Equity,
    ,
    drop = FALSE
  ]
  
  if (ki$ModelGroup == "BEKK") {
    sub_by_date <- aggregate(NET ~ Date, data=sub, FUN=function(x) mean(x, na.rm=TRUE))
    out_model <- "BEKK_mean"
  } else if (ki$ModelGroup == "DCC") {
    sub_by_date <- sub[, c("Date", "NET"), drop=FALSE]
    out_model <- "DCC"
  } else {
    next
  }
  
  s_all <- .summ_vec(sub_by_date$NET)
  
  hd <- high_regime[[ki$Commodity]]$dates
  sub_h <- sub_by_date[as.Date(sub_by_date$Date) %in% as.Date(hd), , drop=FALSE]
  s_high <- .summ_vec(sub_h$NET)
  
  group_summ_rows[[i]] <- data.frame(
    Model=out_model, Commodity=ki$Commodity, Equity=ki$Equity,
    NET_mean_all=s_all["mean"], NET_sd_all=s_all["sd"], N_all=s_all["N"],
    NET_mean_high=s_high["mean"], NET_sd_high=s_high["sd"], N_high=s_high["N"],
    stringsAsFactors = FALSE
  )
}

group_summ_rows <- Filter(Negate(is.null), group_summ_rows)
summ_group_df <- do.call(rbind, group_summ_rows)

write.csv(summ_group_df, out_csv_summ_grouped, row.names=FALSE)
cat("Saved grouped summary:", out_csv_summ_grouped, "\n")

# ------------------------------ JPG NETWORK PLOTS -----------------------------

.dir_create(out_dir)

for (commodity_requested in commodities) {
  
  hd <- high_regime[[commodity_requested]]$dates
  
  # ----- DCC_full plots -----
  dcc_means <- .get_plot_means(
    ts_df = ts_df,
    commodity_requested = commodity_requested,
    model_choice = "DCC_full",
    high_dates = hd
  )
  
  f_all_dcc <- file.path(out_dir, paste0(
    "NET_network_ALL__",
    gsub("[^A-Za-z0-9]+","_", commodity_requested),
    "__DCC_full.jpg"
  ))
  
  f_high_dcc <- file.path(out_dir, paste0(
    "NET_network_HIGH__",
    gsub("[^A-Za-z0-9]+","_", commodity_requested),
    "__DCC_full.jpg"
  ))
  
  .save_jpeg_plot(f_all_dcc, {
    .plot_net_network(
      df_means = dcc_means$mean_all,
      commodity_label = commodity_requested,
      main_title = paste0(
        "NET network (mean over ALL days)\nModel=DCC_full",
        " | Commodity=", commodity_requested,
        " | fixed scale: |NET| in [0, ", EDGE_MAX_ABS, "]"
      ),
      top_k = edge_top_k
    )
  })
  
  .save_jpeg_plot(f_high_dcc, {
    .plot_net_network(
      df_means = dcc_means$mean_high,
      commodity_label = commodity_requested,
      main_title = paste0(
        "NET network (mean over HIGH regime days: top 25% vol in ", regime_model, ")\nModel=DCC_full",
        " | Commodity=", commodity_requested,
        " | fixed scale: |NET| in [0, ", EDGE_MAX_ABS, "]"
      ),
      top_k = edge_top_k
    )
  })
  
  # ----- BEKK mean plots -----
  bekk_means <- .get_plot_means(
    ts_df = ts_df,
    commodity_requested = commodity_requested,
    model_choice = "BEKK_mean",
    high_dates = hd
  )
  
  f_all_bekk <- file.path(out_dir, paste0(
    "NET_network_ALL__",
    gsub("[^A-Za-z0-9]+","_", commodity_requested),
    "__BEKK_mean.jpg"
  ))
  
  f_high_bekk <- file.path(out_dir, paste0(
    "NET_network_HIGH__",
    gsub("[^A-Za-z0-9]+","_", commodity_requested),
    "__BEKK_mean.jpg"
  ))
  
  .save_jpeg_plot(f_all_bekk, {
    .plot_net_network(
      df_means = bekk_means$mean_all,
      commodity_label = commodity_requested,
      main_title = paste0(
        "NET network (mean over ALL days)\nModel=BEKK mean",
        " | Commodity=", commodity_requested,
        " | fixed scale: |NET| in [0, ", EDGE_MAX_ABS, "]"
      ),
      top_k = edge_top_k
    )
  })
  
  .save_jpeg_plot(f_high_bekk, {
    .plot_net_network(
      df_means = bekk_means$mean_high,
      commodity_label = commodity_requested,
      main_title = paste0(
        "NET network (mean over HIGH regime days: top 25% vol in ", regime_model, ")\nModel=BEKK mean",
        " | Commodity=", commodity_requested,
        " | fixed scale: |NET| in [0, ", EDGE_MAX_ABS, "]"
      ),
      top_k = edge_top_k
    )
  })
  
  cat("Saved JPGs:\n  ",
      normalizePath(f_all_dcc, winslash="/", mustWork=FALSE), "\n  ",
      normalizePath(f_high_dcc, winslash="/", mustWork=FALSE), "\n  ",
      normalizePath(f_all_bekk, winslash="/", mustWork=FALSE), "\n  ",
      normalizePath(f_high_bekk, winslash="/", mustWork=FALSE), "\n")
}

cat("\nDONE\n")
cat("Time-series CSV      :", out_csv_ts, "\n")
cat("Model summary CSV    :", out_csv_summ, "\n")
cat("Grouped summary CSV  :", out_csv_summ_grouped, "\n")
cat("Network JPG folder   :", out_dir, "\n")
################################################################################
























################################################################################
# ADD-ON — Alpha sensitivity (HIGH regime): export EACH equity panel as own JPG
#
# Use this updated full block in the SAME script, after the NET/network section.
#
# Fix included:
#   - Explicitly rematches commodity_base from alpha_sig_regime.
#   - Does not rely on .get_high_regime_dates() returning commodity_base.
#   - Prevents ego_node_vol from becoming "_vol".
#
# Terminology:
#   - y-axis/file names/titles use CECI_qw
################################################################################

# ------------------------------ ALPHA SETTINGS --------------------------------

alpha_models_used <- c(
  "sBEKK_sym",
  "dBEKK_sym",
  "dBEKK_asym",
  "DCC_full",
  "DCC_scalar_stage",
  "cDCC_Aielli_stage"
)

alpha_commodities <- c("Brent (Oil)")

# Shared HIGH regime for alpha sensitivity
alpha_regime_model  <- "DCC_full"
alpha_regime_q_high <- 0.90

# Connectedness / alpha sensitivity
alpha_nlag <- 1
alpha_nfore <- 20
alpha_window_size_qw <- 60
alpha_grid <- 0:8
alpha_input_transform <- "dlog_var"

# Output
alpha_out_dir <- file.path(output_dir, "CECI_alpha_panels_jpg")
dir.create(alpha_out_dir, recursive = TRUE, showWarnings = FALSE)

alpha_jpg_w   <- 1400
alpha_jpg_h   <- 1100
alpha_jpg_res <- 240

# ------------------------------ ALPHA HELPERS ---------------------------------

.subset_by_base <- function(vol_zoo, base_names) {
  vol_zoo <- .ensure_vol_colnames(vol_zoo)
  base <- .strip_vol(colnames(vol_zoo))
  keep <- intersect(base, base_names)
  if (length(keep) == 0) stop("subset_by_base: nothing matched.")
  vol_zoo[, paste0(keep, "_vol"), drop = FALSE]
}

.subset_by_dates_zoo <- function(z, dates) {
  z <- .ensure_Date_index(z)
  dates <- as.Date(dates)
  z[index(z) %in% dates, , drop = FALSE]
}

.build_x_from_sigma <- function(sig_zoo, transform = c("dlog_var", "dlog_vol")) {
  transform <- match.arg(transform)
  sig_zoo <- .ensure_Date_index(sig_zoo)
  
  S <- coredata(sig_zoo)
  if (!is.matrix(S)) S <- as.matrix(S)
  S <- pmax(S, 1e-12)
  
  if (transform == "dlog_var") {
    H <- S^2
    X <- diff(log(H))
  } else {
    X <- diff(log(S))
  }
  
  out <- zoo(X, order.by = index(sig_zoo)[-1])
  colnames(out) <- colnames(sig_zoo)
  .ensure_Date_index(out)
}

run_conn_npdc_roll <- function(X, nlag = 1, nfore = 20, window.size = 60) {
  ca <- ConnectednessApproach(
    x = X,
    nlag = nlag,
    nfore = nfore,
    window.size = window.size,
    model = "VAR",
    connectedness = "Time"
  )
  
  cand <- c("NPDC", "npdc", "CT", "GFEVD", "FEVD", "Theta", "theta", "TABLE")
  
  for (nm in cand) {
    obj <- ca[[nm]]
    
    if (is.array(obj) && length(dim(obj)) == 3 && is.numeric(obj)) {
      if (is.null(dimnames(obj)) || is.null(dimnames(obj)[[1]])) {
        dn <- colnames(X)
        dimnames(obj) <- list(dn, dn, NULL)
      }
      return(list(NPDC = obj, ca = ca, source = nm))
    }
    
    if (is.list(obj)) {
      inner_cand <- c("NPDC", "npdc", "CT", "GFEVD", "FEVD", "Theta", "theta")
      
      for (jj in inner_cand) {
        if (!is.null(obj[[jj]]) &&
            is.array(obj[[jj]]) &&
            length(dim(obj[[jj]])) == 3 &&
            is.numeric(obj[[jj]])) {
          
          inner <- obj[[jj]]
          
          if (is.null(dimnames(inner)) || is.null(dimnames(inner)[[1]])) {
            dn <- colnames(X)
            dimnames(inner) <- list(dn, dn, NULL)
          }
          
          return(list(NPDC = inner, ca = ca, source = paste(nm, jj, sep = "$")))
        }
      }
    }
  }
  
  stop("Could not find rolling 3D NPDC/FEVD cube. Names: ", paste(names(ca), collapse = ", "))
}

.quantile_rank <- function(x) {
  r <- rank(x, ties.method = "average", na.last = "keep")
  r / max(r, na.rm = TRUE)
}

# Quantile-weighted directional connectedness difference.
# Returned column is CECI_qw.
quantile_weighted_ceci <- function(X_full,
                                   npdc_roll,
                                   ego_node_vol,
                                   regime_col_vol,
                                   alpha = 4) {
  npdc <- npdc_roll$NPDC
  
  if (length(dim(npdc)) != 3) stop("Need rolling NPDC/FEVD cube, 3D.")
  
  vars <- dimnames(npdc)[[1]]
  if (is.null(vars)) stop("NPDC/FEVD cube has no dimnames().")
  if (!(ego_node_vol %in% vars)) stop("ego_node_vol missing in NPDC/FEVD cube: ", ego_node_vol)
  if (!(regime_col_vol %in% colnames(X_full))) stop("regime_col_vol missing in X_full: ", regime_col_vol)
  
  ego_idx <- match(ego_node_vol, vars)
  others <- setdiff(vars, ego_node_vol)
  Tn <- dim(npdc)[3]
  
  x_roll <- tail(as.numeric(coredata(X_full[, regime_col_vol])), Tn)
  Fv <- .quantile_rank(x_roll)
  
  w <- Fv ^ alpha
  w[!is.finite(w)] <- 0
  
  if (sum(w) <= 0) stop("All alpha weights are zero/NA.")
  
  w <- w / sum(w)
  
  out <- lapply(others, function(v) {
    j <- match(v, vars)
    
    out_series <- vapply(
      seq_len(Tn),
      function(t) as.numeric(npdc[j, ego_idx, t]),
      numeric(1)
    )
    
    in_series <- vapply(
      seq_len(Tn),
      function(t) as.numeric(npdc[ego_idx, j, t]),
      numeric(1)
    )
    
    OUT_qw  <- sum(w * out_series, na.rm = TRUE)
    IN_qw   <- sum(w * in_series,  na.rm = TRUE)
    CECI_qw <- OUT_qw - IN_qw
    
    data.frame(
      Asset = .strip_vol(v),
      CECI_qw = CECI_qw,
      stringsAsFactors = FALSE
    )
  })
  
  do.call(rbind, out)
}

alpha_sensitivity_ceci_high <- function(vol_models,
                                        model,
                                        reg_dates,
                                        assets_keep,
                                        regime_asset_base,
                                        ego_asset_base,
                                        transform = "dlog_var",
                                        window.size_qw = 60,
                                        alpha_grid = 0:8,
                                        nlag = 1,
                                        nfore = 20) {
  
  sig <- .ensure_Date_index(.ensure_vol_colnames(vol_models[[model]], model))
  
  X_full <- .build_x_from_sigma(
    .subset_by_base(sig, assets_keep),
    transform = transform
  )
  
  Xr <- .subset_by_dates_zoo(X_full, reg_dates)
  
  if (nrow(Xr) < (nlag + window.size_qw + 5)) {
    stop("Too few HIGH-regime rows for ", model, ": nrow=", nrow(Xr))
  }
  
  npdc_roll <- run_conn_npdc_roll(
    X = Xr,
    nlag = nlag,
    nfore = nfore,
    window.size = window.size_qw
  )
  
  regime_col_vol <- paste0(regime_asset_base, "_vol")
  ego_node_vol   <- paste0(ego_asset_base, "_vol")
  
  out_list <- lapply(alpha_grid, function(a) {
    tab <- quantile_weighted_ceci(
      X_full = Xr,
      npdc_roll = npdc_roll,
      ego_node_vol = ego_node_vol,
      regime_col_vol = regime_col_vol,
      alpha = a
    )
    
    tab$alpha <- a
    tab
  })
  
  res <- do.call(rbind, out_list)
  res$Model <- model
  res
}

.plot_alpha_panel_one_asset_to_jpg <- function(alpha_df,
                                               asset,
                                               models,
                                               main_title_lines,
                                               out_file,
                                               width = 1400,
                                               height = 1100,
                                               res = 240) {
  d <- alpha_df[
    alpha_df$Asset == asset &
      alpha_df$Model %in% models,
    ,
    drop = FALSE
  ]
  
  if (nrow(d) == 0) return(invisible(FALSE))
  
  models <- intersect(models, unique(d$Model))
  if (length(models) < 1) return(invisible(FALSE))
  
  cols <- grDevices::hcl.colors(length(models), palette = "Dark 3")
  names(cols) <- models
  
  ltys <- rep(1:6, length.out = length(models))
  names(ltys) <- models
  
  yr <- range(d$CECI_qw, finite = TRUE)
  if (!all(is.finite(yr))) yr <- c(-1, 1)
  
  grDevices::jpeg(out_file, width = width, height = height, res = res, quality = 96)
  on.exit(grDevices::dev.off(), add = TRUE)
  
  op <- par(no.readonly = TRUE)
  on.exit(par(op), add = TRUE)
  
  layout(matrix(c(1, 2), nrow = 2), heights = c(4.8, 0.9))
  
  par(mar = c(3.4, 3.8, 3.0, 1.0), mgp = c(2.1, 0.55, 0))
  
  m0 <- models[1]
  d0 <- d[d$Model == m0, , drop = FALSE]
  d0 <- d0[order(d0$alpha), ]
  
  main_txt <- paste(c(main_title_lines, paste0(asset)), collapse = "\n")
  
  plot(
    d0$alpha,
    d0$CECI_qw,
    type = "b",
    xlab = "alpha",
    ylab = "CECI_qw",
    main = main_txt,
    cex.main = 0.95,
    ylim = yr,
    lty = ltys[m0],
    col = cols[m0]
  )
  
  abline(h = 0, lty = 2)
  
  if (length(models) > 1) {
    for (m in models[-1]) {
      dm <- d[d$Model == m, , drop = FALSE]
      dm <- dm[order(dm$alpha), ]
      
      lines(
        dm$alpha,
        dm$CECI_qw,
        type = "b",
        lty = ltys[m],
        col = cols[m]
      )
    }
  }
  
  par(mar = c(0.1, 0.2, 0.1, 0.2))
  plot.new()
  
  legend(
    "center",
    legend = models,
    bty = "n",
    ncol = 3,
    lty = ltys[models],
    col = cols[models],
    cex = 0.9
  )
  
  invisible(TRUE)
}

# ------------------------------ ALPHA MAIN ------------------------------------

cat("\nSTART: CECI_qw alpha panels\n")

alpha_miss <- setdiff(unique(c(alpha_models_used, alpha_regime_model)), names(vol_models))
if (length(alpha_miss) > 0) {
  stop("vol_models is missing for alpha add-on: ", paste(alpha_miss, collapse = ", "))
}

vol_models[[alpha_regime_model]] <- .ensure_Date_index(
  .ensure_vol_colnames(vol_models[[alpha_regime_model]], alpha_regime_model)
)

alpha_sig_regime <- vol_models[[alpha_regime_model]]

for (commodity_requested in alpha_commodities) {
  
  high <- .get_high_regime_dates(
    alpha_sig_regime,
    commodity_requested,
    q_high = alpha_regime_q_high
  )
  
  reg_dates_high <- high$dates
  
  # Critical fix: explicitly match commodity base.
  alpha_base_cols <- .strip_vol(colnames(alpha_sig_regime))
  commodity_base <- .match_one(commodity_requested, alpha_base_cols)
  
  if (is.na(commodity_base) || !nzchar(commodity_base)) {
    stop(
      "Could not match alpha commodity '",
      commodity_requested,
      "' in alpha_regime_model columns. Available bases: ",
      paste(head(alpha_base_cols, 20), collapse = ", ")
    )
  }
  
  cat(
    "Commodity:",
    commodity_requested,
    "| matched=",
    commodity_base,
    "| threshold=",
    round(high$threshold, 6),
    "| HIGH days=",
    length(reg_dates_high),
    "\n"
  )
  
  base_cols <- .strip_vol(colnames(alpha_sig_regime))
  equities <- sort(base_cols[grepl(equity_regex, base_cols)])
  
  if (length(equities) == 0) {
    stop("No equities matched equity_regex in alpha_regime_model.")
  }
  
  assets_keep <- unique(c(commodity_base, equities))
  
  alpha_all <- list()
  
  for (m in alpha_models_used) {
    cat("  model:", m, "...\n")
    
    tmp <- try(
      alpha_sensitivity_ceci_high(
        vol_models = vol_models,
        model = m,
        reg_dates = reg_dates_high,
        assets_keep = assets_keep,
        regime_asset_base = commodity_base,
        ego_asset_base = commodity_base,
        transform = alpha_input_transform,
        window.size_qw = alpha_window_size_qw,
        alpha_grid = alpha_grid,
        nlag = alpha_nlag,
        nfore = alpha_nfore
      ),
      silent = TRUE
    )
    
    if (inherits(tmp, "try-error")) {
      message("    skipped (error): ", as.character(tmp))
    } else {
      alpha_all[[m]] <- tmp
    }
  }
  
  if (length(alpha_all) < 2) {
    warning("Need at least 2 successful models for overlays. Skipping: ", commodity_requested)
    next
  }
  
  alpha_df <- do.call(rbind, alpha_all)
  alpha_df <- alpha_df[alpha_df$Asset %in% equities, , drop = FALSE]
  
  alpha_csv <- file.path(
    alpha_out_dir,
    paste0(
      "CECI_qw_alpha_sensitivity_HIGH__",
      gsub("[^A-Za-z0-9]+", "_", commodity_requested),
      ".csv"
    )
  )
  
  write.csv(alpha_df, alpha_csv, row.names = FALSE)
  cat("  saved alpha CSV:", normalizePath(alpha_csv, winslash = "/", mustWork = FALSE), "\n")
  
  title_lines <- c(
    "CECI_qw vs alpha — OVERLAY — HIGH",
    paste0(
      "(shared regime=",
      alpha_regime_model,
      ", asset=",
      commodity_requested,
      ", window.size=",
      alpha_window_size_qw,
      ")"
    )
  )
  
  for (as in equities) {
    out_file <- file.path(
      alpha_out_dir,
      paste0(
        "CECI_qw_vs_alpha_OVERLAY_HIGH__",
        gsub("[^A-Za-z0-9]+", "_", commodity_requested),
        "__",
        as,
        ".jpg"
      )
    )
    
    ok <- .plot_alpha_panel_one_asset_to_jpg(
      alpha_df = alpha_df,
      asset = as,
      models = alpha_models_used,
      main_title_lines = title_lines,
      out_file = out_file,
      width = alpha_jpg_w,
      height = alpha_jpg_h,
      res = alpha_jpg_res
    )
    
    if (isTRUE(ok)) {
      cat("  saved:", normalizePath(out_file, winslash = "/", mustWork = FALSE), "\n")
    }
  }
}

cat("DONE: CECI_qw alpha panels\n")
################################################################################
































################################################################################
# Commodity-specific 6-variable systems: 1 commodity + 5 equity sectors
# Based on existing `vol_models`
#
# AGGREGATION RULE
#   - BEKK mean = arithmetic mean across:
#       sBEKK_sym, dBEKK_sym, dBEKK_asym
#   - DCC = DCC_full
#
# PURPOSE
#   For each commodity separately, estimate connectedness on:
#     {commodity, SX6P, SXEP, SXQP, SX7P, SXNP}
#
#   Then build two figures:
#
#   FIGURE 1: Heatmap of tail amplification
#       Delta_tail = NET_qw(alpha_max) - NET_qw(alpha_min)
#
#   FIGURE 2: Compact small-multiples alpha profiles
#       rows = commodities, cols = sectors
#       lines = BEKK mean and DCC
#
# REGIME
#   Only top 25% volatility state is used, defined within each commodity-specific
#   6-variable system using the regime-defining commodity volatility series.
#
# FILES SAVED TO output_dir if it exists, otherwise current directory:
#   figure1_heatmap_top25_subsystem_BEKKmean_DCC.pdf
#   figure1_heatmap_top25_subsystem_BEKKmean_DCC.png
#   figure2_smallmultiples_top25_subsystem_BEKKmean_DCC.pdf
#   figure2_smallmultiples_top25_subsystem_BEKKmean_DCC.png
#   alpha_all_top25_subsystem_raw_models.csv
#   alpha_all_top25_subsystem_BEKKmean_DCC.csv
#   heatmap_delta_top25_subsystem_BEKKmean_DCC.csv
################################################################################

suppressPackageStartupMessages({
  library(zoo)
  library(ConnectednessApproach)
  library(ggplot2)
  library(dplyr)
  library(tidyr)
})

################################################################################
# 0) USER SETTINGS
################################################################################

# raw models actually estimated
models_raw <- c("sBEKK_sym", "dBEKK_sym", "dBEKK_asym", "DCC_full")

bekk_models <- c("sBEKK_sym", "dBEKK_sym", "dBEKK_asym")
dcc_model   <- "DCC_full"

commodities <- c("TTF (Gas)", "Brent (Oil)", "API2 (Coal)", "MO1 (Carbon)")

sector_order <- c(
  "SX6P (Utilities)",
  "SXEP (Oil & Gas)",
  "SXQP (Consumer)",
  "SX7P (Financials)",
  "SXNP (Industrials)"
)

assets_required <- c(commodities, sector_order)

# IMPORTANT:
#   dlog_var = Delta log(sigma^2), same convention as your directional CECI work.
#   dlog     = Delta log(sigma), included for compatibility with your old script.
mode_used       <- "dlog_var"   # "level", "loglevel", "dlog", "dlog_var", "dlog_vol"
q_high_top25    <- 0.75
alpha_grid_used <- 0:8
alpha_min_used  <- 0
alpha_max_used  <- 8
nlag_used       <- 1
nfore_used      <- 20
window_qw_used  <- 60

out_base_dir <- if (exists("output_dir")) output_dir else getwd()
if (!dir.exists(out_base_dir)) dir.create(out_base_dir, recursive = TRUE)

################################################################################
# 1) HELPERS
################################################################################

.ensure_vol_colnames <- function(vol_zoo, model_name = "model") {
  stopifnot(inherits(vol_zoo, "zoo"))
  cn <- colnames(vol_zoo)
  if (is.null(cn)) stop(model_name, ": vol_zoo has no colnames().")
  colnames(vol_zoo) <- ifelse(grepl("_vol$", cn), cn, paste0(cn, "_vol"))
  vol_zoo
}

.ensure_Date_index <- function(z) {
  stopifnot(inherits(z, "zoo"))
  idx <- index(z)
  
  if (inherits(idx, "Date")) {
    if (anyDuplicated(idx)) {
      warning("Duplicate Date index entries detected. Keeping last row per Date.")
      z <- z[!duplicated(idx, fromLast = TRUE), ]
    }
    return(z)
  }
  
  if (inherits(idx, c("POSIXct", "POSIXt"))) {
    index(z) <- as.Date(idx)
  } else if (inherits(idx, c("yearmon", "yearqtr"))) {
    index(z) <- as.Date(idx)
  } else if (is.numeric(idx)) {
    index(z) <- as.Date(idx, origin = "1970-01-01")
  } else {
    idx2 <- suppressWarnings(as.Date(idx))
    if (all(is.na(idx2))) {
      stop("Could not coerce zoo index to Date. Current class: ", paste(class(idx), collapse = ", "))
    }
    index(z) <- idx2
  }
  
  if (anyDuplicated(index(z))) {
    warning("Duplicate Date index entries detected after Date coercion. Keeping last row per Date.")
    z <- z[!duplicated(index(z), fromLast = TRUE), ]
  }
  
  z
}

.strip_vol <- function(x) sub("_vol$", "", x)

prep_input <- function(vol_zoo,
                       mode = c("level", "loglevel", "dlog", "dlog_var", "dlog_vol")) {
  mode <- match.arg(mode)
  
  vol_zoo <- .ensure_Date_index(vol_zoo)
  
  Xmat <- coredata(vol_zoo)
  if (!is.matrix(Xmat)) Xmat <- as.matrix(Xmat)
  Xmat <- pmax(Xmat, 1e-12)
  
  if (mode == "level") {
    Xout <- Xmat
    idx <- index(vol_zoo)
  } else if (mode == "loglevel") {
    Xout <- log(Xmat)
    idx <- index(vol_zoo)
  } else if (mode %in% c("dlog", "dlog_vol")) {
    Xout <- diff(log(Xmat))
    idx <- index(vol_zoo)[-1]
  } else if (mode == "dlog_var") {
    Xout <- diff(log(Xmat^2))
    idx <- index(vol_zoo)[-1]
  }
  
  colnames(Xout) <- colnames(vol_zoo)
  X <- zoo(Xout, order.by = idx)
  X <- X[complete.cases(X), , drop = FALSE]
  .ensure_Date_index(X)
}

.subset_by_base <- function(vol_zoo, base_keep) {
  vol_zoo <- .ensure_vol_colnames(vol_zoo)
  cols <- paste0(base_keep, "_vol")
  missing <- setdiff(cols, colnames(vol_zoo))
  if (length(missing) > 0) {
    stop(
      "Missing columns: ", paste(missing, collapse = ", "),
      "\nAvailable: ", paste(colnames(vol_zoo), collapse = ", ")
    )
  }
  vol_zoo[, cols, drop = FALSE]
}

.subset_by_dates_zoo <- function(Z, dates) {
  stopifnot(inherits(Z, "zoo"))
  Z <- .ensure_Date_index(Z)
  if (length(dates) == 0) return(Z[0, , drop = FALSE])
  keep <- as.character(index(Z)) %in% as.character(as.Date(dates))
  Z[keep, , drop = FALSE]
}

.quantile_rank <- function(x) {
  r <- rank(x, ties.method = "average", na.last = "keep")
  mx <- max(r, na.rm = TRUE)
  if (!is.finite(mx) || mx <= 0) return(rep(NA_real_, length(x)))
  r / mx
}

run_conn <- function(X, nlag = 1, nfore = 20, window.size = NULL) {
  ConnectednessApproach(
    X,
    nlag = nlag,
    nfore = nfore,
    model = "VAR",
    connectedness = "Time",
    window.size = window.size,
    corrected = TRUE
  )
}

.get_npdc_cube <- function(dca_roll, X) {
  cand <- c("NPDC", "npdc", "CT", "GFEVD", "FEVD", "Theta", "theta", "TABLE")
  
  for (nm in cand) {
    obj <- dca_roll[[nm]]
    
    if (is.array(obj) && length(dim(obj)) == 3 && is.numeric(obj)) {
      cube <- obj
      if (is.null(dimnames(cube)) || is.null(dimnames(cube)[[1]])) {
        dimnames(cube) <- list(colnames(X), colnames(X), NULL)
      }
      return(cube)
    }
    
    if (is.list(obj)) {
      inner_cand <- c("NPDC", "npdc", "CT", "GFEVD", "FEVD", "Theta", "theta")
      for (jj in inner_cand) {
        if (!is.null(obj[[jj]]) &&
            is.array(obj[[jj]]) &&
            length(dim(obj[[jj]])) == 3 &&
            is.numeric(obj[[jj]])) {
          cube <- obj[[jj]]
          if (is.null(dimnames(cube)) || is.null(dimnames(cube)[[1]])) {
            dimnames(cube) <- list(colnames(X), colnames(X), NULL)
          }
          return(cube)
        }
      }
    }
  }
  
  stop("Rolling NPDC/FEVD cube not available. Names: ", paste(names(dca_roll), collapse = ", "))
}

commodity_subsystem <- function(commodity, sectors = sector_order) {
  c(commodity, sectors)
}

.aggregate_model_groups <- function(alpha_df) {
  d_bekk <- alpha_df %>%
    dplyr::filter(Model %in% c("sBEKK_sym", "dBEKK_sym", "dBEKK_asym")) %>%
    dplyr::group_by(Commodity, Asset, alpha) %>%
    dplyr::summarise(
      OUT_qw = mean(OUT_qw, na.rm = TRUE),
      IN_qw  = mean(IN_qw,  na.rm = TRUE),
      NET_qw = mean(NET_qw, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::mutate(Model = "BEKK mean")
  
  d_dcc <- alpha_df %>%
    dplyr::filter(Model == "DCC_full") %>%
    dplyr::mutate(Model = "DCC")
  
  dplyr::bind_rows(d_bekk, d_dcc) %>%
    dplyr::mutate(
      Model = factor(Model, levels = c("DCC", "BEKK mean"))
    )
}

################################################################################
# 2) VALIDATION
################################################################################

if (!exists("vol_models")) {
  stop("`vol_models` does not exist in your environment.")
}

missing_models <- setdiff(models_raw, names(vol_models))
if (length(missing_models) > 0) {
  stop("vol_models is missing: ", paste(missing_models, collapse = ", "))
}

vol_models <- stats::setNames(
  lapply(names(vol_models), function(nm) {
    .ensure_Date_index(.ensure_vol_colnames(vol_models[[nm]], nm))
  }),
  names(vol_models)
)

for (m in models_raw) {
  if (!inherits(vol_models[[m]], "zoo")) {
    stop("vol_models[['", m, "']] must be a zoo object.")
  }
  miss_cols <- setdiff(paste0(assets_required, "_vol"), colnames(vol_models[[m]]))
  if (length(miss_cols) > 0) {
    stop(
      "Model ", m, " is missing required columns: ",
      paste(miss_cols, collapse = ", ")
    )
  }
}

################################################################################
# 3) CORE FUNCTION:
#    alpha sensitivity in TOP-25% regime for one commodity-specific 6-variable system
################################################################################

alpha_sensitivity_top25_one_commodity_subsystem <- function(vol_models,
                                                            model,
                                                            regime_model,
                                                            regime_asset,
                                                            ego_asset = regime_asset,
                                                            sectors = sector_order,
                                                            mode = "dlog_var",
                                                            q_high = 0.75,
                                                            window.size_qw = 60,
                                                            alpha_grid = 0:8,
                                                            nlag = 1,
                                                            nfore = 20) {
  assets_keep <- commodity_subsystem(regime_asset, sectors)
  
  vol_m_reg <- .ensure_Date_index(.ensure_vol_colnames(vol_models[[regime_model]], regime_model))
  vol_m     <- .ensure_Date_index(.ensure_vol_colnames(vol_models[[model]], model))
  
  regime_col_vol <- paste0(regime_asset, "_vol")
  ego_node_vol   <- paste0(ego_asset, "_vol")
  
  X_reg_full <- prep_input(.subset_by_base(vol_m_reg, assets_keep), mode = mode)
  if (!(regime_col_vol %in% colnames(X_reg_full))) {
    stop("Regime column not found: ", regime_col_vol)
  }
  
  x_reg <- as.numeric(coredata(X_reg_full[, regime_col_vol]))
  qh <- stats::quantile(x_reg, q_high, na.rm = TRUE)
  reg_dates <- index(X_reg_full)[which(is.finite(x_reg) & x_reg >= qh)]
  
  X_full <- prep_input(.subset_by_base(vol_m, assets_keep), mode = mode)
  Xr <- .subset_by_dates_zoo(X_full, reg_dates)
  
  if (nrow(Xr) < (nlag + window.size_qw + 5)) {
    stop(
      "Too few top-25% observations for model ", model,
      " and commodity subsystem ", regime_asset, ": ", nrow(Xr),
      ". Need at least ", nlag + window.size_qw + 5, "."
    )
  }
  
  dca_roll <- run_conn(Xr, nlag = nlag, nfore = nfore, window.size = window.size_qw)
  npdc <- .get_npdc_cube(dca_roll, Xr)
  
  vars <- dimnames(npdc)[[1]]
  if (is.null(vars)) stop("Rolling NPDC/FEVD cube has no dimnames.")
  if (!(ego_node_vol %in% vars)) stop("ego_node_vol not found in NPDC: ", ego_node_vol)
  
  ego_idx <- match(ego_node_vol, vars)
  others <- setdiff(vars, ego_node_vol)
  Tn <- dim(npdc)[3]
  
  x_roll <- tail(as.numeric(coredata(Xr[, regime_col_vol])), Tn)
  Fv <- .quantile_rank(x_roll)
  
  out_list <- lapply(alpha_grid, function(a) {
    w <- Fv^a
    w[!is.finite(w)] <- 0
    if (sum(w) <= 0) stop("All alpha weights are zero/NA.")
    w <- w / sum(w)
    
    out <- lapply(others, function(v) {
      j <- match(v, vars)
      
      out_series <- vapply(seq_len(Tn), function(t) as.numeric(npdc[j, ego_idx, t]), numeric(1))
      in_series  <- vapply(seq_len(Tn), function(t) as.numeric(npdc[ego_idx, j, t]), numeric(1))
      
      OUT_qw <- sum(w * out_series, na.rm = TRUE)
      IN_qw  <- sum(w * in_series,  na.rm = TRUE)
      
      data.frame(
        Commodity = regime_asset,
        Model     = model,
        Asset     = .strip_vol(v),
        alpha     = a,
        OUT_qw    = OUT_qw,
        IN_qw     = IN_qw,
        NET_qw    = OUT_qw - IN_qw,
        stringsAsFactors = FALSE
      )
    })
    
    dplyr::bind_rows(out)
  })
  
  dplyr::bind_rows(out_list)
}

################################################################################
# 4) BUILD FULL ALPHA TABLE FOR ALL 4 COMMODITIES
################################################################################

build_alpha_all_top25_subsystems <- function(vol_models,
                                             models_use,
                                             regime_model = "DCC_full",
                                             commodities,
                                             sectors = sector_order,
                                             mode = "dlog_var",
                                             q_high = 0.75,
                                             window.size_qw = 60,
                                             alpha_grid = 0:8,
                                             nlag = 1,
                                             nfore = 20) {
  out <- list()
  k <- 1
  
  for (com in commodities) {
    for (m in models_use) {
      message("Running subsystem commodity = ", com, " | model = ", m)
      
      tmp <- try(
        alpha_sensitivity_top25_one_commodity_subsystem(
          vol_models      = vol_models,
          model           = m,
          regime_model    = regime_model,
          regime_asset    = com,
          ego_asset       = com,
          sectors         = sectors,
          mode            = mode,
          q_high          = q_high,
          window.size_qw  = window.size_qw,
          alpha_grid      = alpha_grid,
          nlag            = nlag,
          nfore           = nfore
        ),
        silent = TRUE
      )
      
      if (inherits(tmp, "try-error")) {
        message("  skipped: ", as.character(tmp))
      } else {
        out[[k]] <- tmp
        k <- k + 1
      }
    }
  }
  
  if (length(out) == 0) {
    stop("No alpha subsystem results were produced.")
  }
  
  dplyr::bind_rows(out)
}

################################################################################
# 5) FIGURE 1 DATA: HEATMAP OF TAIL AMPLIFICATION
################################################################################

make_tail_delta_heatmap_data <- function(alpha_df,
                                         alpha_min = 0,
                                         alpha_max = 8,
                                         commodity_order = NULL,
                                         sector_order = NULL) {
  df <- alpha_df %>%
    dplyr::filter(!is.na(Commodity), !is.na(Asset)) %>%
    dplyr::filter(Asset %in% sector_order) %>%
    dplyr::filter(alpha %in% c(alpha_min, alpha_max)) %>%
    dplyr::mutate(alpha_lab = ifelse(alpha == alpha_min, "start", "end")) %>%
    dplyr::select(Commodity, Model, Asset, alpha_lab, NET_qw) %>%
    tidyr::pivot_wider(names_from = alpha_lab, values_from = NET_qw) %>%
    dplyr::mutate(delta_tail = end - start)
  
  if (!is.null(commodity_order)) {
    df$Commodity <- factor(df$Commodity, levels = commodity_order)
  }
  if (!is.null(sector_order)) {
    df$Asset <- factor(df$Asset, levels = sector_order)
  }
  if ("Model" %in% names(df)) {
    df$Model <- factor(df$Model, levels = c("DCC", "BEKK mean"))
  }
  
  droplevels(df)
}

plot_tail_delta_heatmap <- function(df_heat) {
  ggplot(df_heat, aes(x = Asset, y = Commodity, fill = delta_tail)) +
    geom_tile(color = "white", linewidth = 0.7) +
    geom_text(aes(label = sprintf("%.2f", delta_tail)), size = 3.1) +
    facet_wrap(~ Model, nrow = 1) +
    scale_fill_gradient2(
      low = "#B2182B",
      mid = "white",
      high = "#2166AC",
      midpoint = 0,
      name = expression(Delta["tail"])
    ) +
    labs(
      title = "Tail amplification of quantile-weighted net spillovers",
      subtitle = "Top 25% volatility regime; positive values indicate stronger outward spillover as tail emphasis increases",
      x = NULL,
      y = NULL
    ) +
    theme_minimal(base_size = 12) +
    theme(
      panel.grid = element_blank(),
      axis.text.x = element_text(angle = 35, hjust = 1),
      strip.text = element_text(face = "bold"),
      plot.title = element_text(face = "bold"),
      legend.position = "right"
    )
}

################################################################################
# 6) FIGURE 2: COMPACT SMALL-MULTIPLES
################################################################################

plot_small_multiples_alpha <- function(alpha_df,
                                       commodity_order = NULL,
                                       sector_order = NULL,
                                       models_show = c("DCC", "BEKK mean")) {
  d <- alpha_df %>%
    dplyr::filter(Model %in% models_show) %>%
    dplyr::filter(!is.na(Commodity), !is.na(Asset)) %>%
    dplyr::filter(Asset %in% sector_order)
  
  if (!is.null(commodity_order)) {
    d$Commodity <- factor(d$Commodity, levels = commodity_order)
  }
  if (!is.null(sector_order)) {
    d$Asset <- factor(d$Asset, levels = sector_order)
  }
  
  d$Model <- factor(d$Model, levels = c("DCC", "BEKK mean"))
  d <- droplevels(d)
  
  ggplot(d, aes(x = alpha, y = NET_qw, color = Model, linetype = Model)) +
    geom_hline(yintercept = 0, linetype = 2, color = "grey60", linewidth = 0.4) +
    geom_line(linewidth = 0.75) +
    geom_point(size = 1.15) +
    facet_grid(Commodity ~ Asset, scales = "free_y", drop = TRUE) +
    labs(
      title = "Quantile-weighted net spillovers by commodity and equity sector",
      subtitle = "Top 25% volatility regime; 6-variable subsystems (1 commodity + 5 sectors)",
      x = expression(alpha),
      y = expression(CECI[qw]^NET(alpha))
    ) +
    theme_minimal(base_size = 11) +
    theme(
      strip.text = element_text(face = "bold"),
      plot.title = element_text(face = "bold"),
      legend.position = "bottom",
      panel.grid.minor = element_blank()
    )
}

################################################################################
# 7) RUN EVERYTHING
################################################################################

alpha_all_raw <- build_alpha_all_top25_subsystems(
  vol_models      = vol_models,
  models_use      = models_raw,
  regime_model    = "DCC_full",
  commodities     = commodities,
  sectors         = sector_order,
  mode            = mode_used,
  q_high          = q_high_top25,
  window.size_qw  = window_qw_used,
  alpha_grid      = alpha_grid_used,
  nlag            = nlag_used,
  nfore           = nfore_used
)

alpha_all <- .aggregate_model_groups(alpha_all_raw)

heat_df <- make_tail_delta_heatmap_data(
  alpha_df         = alpha_all,
  alpha_min        = alpha_min_used,
  alpha_max        = alpha_max_used,
  commodity_order  = commodities,
  sector_order     = sector_order
)

fig1_heatmap <- plot_tail_delta_heatmap(heat_df)

fig2_smallmult <- plot_small_multiples_alpha(
  alpha_df        = alpha_all,
  commodity_order = commodities,
  sector_order    = sector_order,
  models_show     = c("DCC", "BEKK mean")
)

print(fig1_heatmap)
print(fig2_smallmult)

################################################################################
# 8) SAVE PLOTS
################################################################################

ggsave(
  file.path(out_base_dir, "figure1_heatmap_top25_subsystem_BEKKmean_DCC.pdf"),
  plot = fig1_heatmap,
  width = 10,
  height = 4.8
)

ggsave(
  file.path(out_base_dir, "figure1_heatmap_top25_subsystem_BEKKmean_DCC.png"),
  plot = fig1_heatmap,
  width = 10,
  height = 4.8,
  dpi = 300
)

ggsave(
  file.path(out_base_dir, "figure2_smallmultiples_top25_subsystem_BEKKmean_DCC.pdf"),
  plot = fig2_smallmult,
  width = 13,
  height = 7.5
)

ggsave(
  file.path(out_base_dir, "figure2_smallmultiples_top25_subsystem_BEKKmean_DCC.png"),
  plot = fig2_smallmult,
  width = 13,
  height = 7.5,
  dpi = 300
)

################################################################################
# 9) SAVE TABLES
################################################################################

write.csv(
  alpha_all_raw,
  file.path(out_base_dir, "alpha_all_top25_subsystem_raw_models.csv"),
  row.names = FALSE
)

write.csv(
  alpha_all,
  file.path(out_base_dir, "alpha_all_top25_subsystem_BEKKmean_DCC.csv"),
  row.names = FALSE
)

write.csv(
  heat_df,
  file.path(out_base_dir, "heatmap_delta_top25_subsystem_BEKKmean_DCC.csv"),
  row.names = FALSE
)

cat("\nDONE\n")
cat("Saved outputs to:", normalizePath(out_base_dir, winslash = "/", mustWork = FALSE), "\n")
################################################################################
# END
################################################################################