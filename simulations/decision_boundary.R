###############################################################################
# decision_boundary.R
#
# Given ONLY q0 as a known variable, find the decision boundary that
# determines whether PCA or DIIM key sectors will yield better economic
# loss reduction.
#
# Approach:
#   1. Monte Carlo: 2000 random q0 vectors, fixed c_star (base), fixed A_star
#   2. Compute q0-derived features for each scenario
#   3. Fit logistic regression on q0 features only
#   4. Find optimal threshold via ROC analysis
#   5. Derive a simple, interpretable decision rule
#
# Run from the project root:
#   Rscript simulations/decision_boundary.R
###############################################################################

library(openxlsx)
library(ggplot2)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

# ─────────────────────────────────────────────────────────────────────────────
# Load base data
# ─────────────────────────────────────────────────────────────────────────────
data <- download_data()
A <- data$A
x <- data$x
c_star_base <- data$c_star
A_star <- data$A_star
q0_base <- data$q0

n_sectors <- length(q0_base)
pca_sectors <- c(3, 2, 5, 8, 10)
non_pca_sectors <- setdiff(1:n_sectors, pca_sectors)

cat("=== Decision Boundary Analysis (q0 only) ===\n")
cat("PCA sectors:", pca_sectors, "\n")
cat("Non-PCA sectors:", non_pca_sectors, "\n\n")

# ─────────────────────────────────────────────────────────────────────────────
# Helper: compare methods (same as run_simulations.R)
# ─────────────────────────────────────────────────────────────────────────────
compare_methods <- function(q0, A_star, c_star, x,
                            lockdown_duration, total_duration,
                            pca_sectors_arg = pca_sectors) {
    model_base <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration)
    EL_base <- model_base$EL_evolution
    max_econ_loss <- apply(EL_base, 1, max)
    diim_sectors <- order(max_econ_loss, decreasing = TRUE)[1:5]

    model_diim <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
        key_sectors = diim_sectors
    )
    model_pca <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
        key_sectors = pca_sectors_arg
    )

    loss_base <- model_base$total_economic_loss
    loss_diim <- model_diim$total_economic_loss
    loss_pca <- model_pca$total_economic_loss

    reduction_diim <- loss_base - loss_diim
    reduction_pca <- loss_base - loss_pca

    return(list(
        loss_base = loss_base,
        loss_diim = loss_diim,
        loss_pca = loss_pca,
        reduction_diim = reduction_diim,
        reduction_pca = reduction_pca,
        pca_wins = reduction_pca > reduction_diim,
        diim_sectors = diim_sectors,
        overlap = length(intersect(pca_sectors_arg, diim_sectors))
    ))
}

# Gini coefficient
gini <- function(x) {
    x <- sort(abs(x))
    n <- length(x)
    if (sum(x) == 0) {
        return(0)
    }
    index <- 1:n
    return((2 * sum(index * x) / (n * sum(x))) - (n + 1) / n)
}

# ─────────────────────────────────────────────────────────────────────────────
# Compute q0-derived features
# ─────────────────────────────────────────────────────────────────────────────
compute_q0_features <- function(q0, pca_idx, non_pca_idx) {
    q0_pca <- q0[pca_idx]
    q0_non_pca <- q0[non_pca_idx]

    # Share of total q0 in PCA sectors
    q0_pca_share <- sum(q0_pca) / sum(q0)

    # Mean ratio: PCA sector mean / non-PCA sector mean
    mean_pca <- mean(q0_pca)
    mean_non_pca <- mean(q0_non_pca)
    q0_mean_ratio <- mean_pca / max(mean_non_pca, 1e-10)

    # Max ratio: max in PCA / max in non-PCA
    q0_max_ratio <- max(q0_pca) / max(max(q0_non_pca), 1e-10)

    # How many of the top-5 q0 sectors are PCA sectors?
    top5_q0 <- order(q0, decreasing = TRUE)[1:5]
    q0_top5_overlap <- length(intersect(top5_q0, pca_idx))

    # Rank-based: average rank of PCA sectors (lower rank = higher q0)
    # Rank 1 = highest q0
    ranks <- rank(-q0)
    q0_pca_avg_rank <- mean(ranks[pca_idx])

    # Coefficient of variation within PCA sectors
    q0_pca_cv <- sd(q0_pca) / max(mean_pca, 1e-10)

    # Overall distribution metrics
    q0_gini <- gini(q0)
    q0_cv <- sd(q0) / max(mean(q0), 1e-10)

    return(data.frame(
        q0_pca_share = q0_pca_share,
        q0_mean_ratio = q0_mean_ratio,
        q0_max_ratio = q0_max_ratio,
        q0_top5_overlap = q0_top5_overlap,
        q0_pca_avg_rank = q0_pca_avg_rank,
        q0_pca_cv = q0_pca_cv,
        q0_gini = q0_gini,
        q0_cv = q0_cv
    ))
}

# ─────────────────────────────────────────────────────────────────────────────
# PART 1: Monte Carlo — 2000 random q0 vectors
# ─────────────────────────────────────────────────────────────────────────────
cat("--- Part 1: Monte Carlo (2000 random q0 vectors) ---\n")

set.seed(123)
n_mc <- 2000
max_q0 <- max(q0_base) * 1.5

mc_results <- vector("list", n_mc)

for (i in 1:n_mc) {
    q0_rand <- runif(n_sectors, min = 1e-6, max = max_q0)

    res <- compare_methods(q0_rand, A_star, c_star_base, x,
        lockdown_duration = 55, total_duration = 751
    )
    features <- compute_q0_features(q0_rand, pca_sectors, non_pca_sectors)

    mc_results[[i]] <- data.frame(
        sim_id = i,
        features,
        reduction_diim = res$reduction_diim,
        reduction_pca = res$reduction_pca,
        pca_advantage = res$reduction_pca - res$reduction_diim,
        pca_wins = res$pca_wins,
        overlap = res$overlap,
        stringsAsFactors = FALSE
    )

    if (i %% 200 == 0) {
        pca_rate <- mean(sapply(mc_results[1:i], function(r) r$pca_wins))
        cat(sprintf("  %d / %d done (PCA win rate: %.1f%%)\n", i, n_mc, pca_rate * 100))
    }
}

mc_df <- do.call(rbind, mc_results)

write.csv(mc_df, "simulations/results/decision_boundary_mc.csv", row.names = FALSE)
cat(sprintf(
    "  Total: PCA wins %d / %d (%.1f%%)\n\n",
    sum(mc_df$pca_wins), nrow(mc_df), mean(mc_df$pca_wins) * 100
))


# ─────────────────────────────────────────────────────────────────────────────
# PART 2: Logistic Regression — q0 features only
# ─────────────────────────────────────────────────────────────────────────────
cat("--- Part 2: Logistic Regression (q0 features only) ---\n")

mc_df$pca_wins_int <- as.integer(mc_df$pca_wins)

# Full model with all q0-derived features
logit_full <- glm(
    pca_wins_int ~ q0_pca_share + q0_mean_ratio + q0_max_ratio +
        q0_top5_overlap + q0_pca_avg_rank + q0_pca_cv +
        q0_gini + q0_cv,
    data = mc_df, family = binomial
)

cat("\nFull Model:\n")
print(summary(logit_full))

# Simplified model — use only the single best predictor for interpretability
# Try q0_pca_share first (was the strongest predictor in previous analysis)
logit_simple <- glm(pca_wins_int ~ q0_pca_share,
    data = mc_df, family = binomial
)

cat("\n\nSimple Model (q0_pca_share only):\n")
print(summary(logit_simple))

# Extract the decision boundary from the simple model:
# P(PCA wins) = 0.5 when intercept + coef * q0_pca_share = 0
# => q0_pca_share_threshold = -intercept / coef
coefs_simple <- coef(logit_simple)
threshold_share <- -coefs_simple[1] / coefs_simple[2]
cat(sprintf("\n** q0_pca_share threshold (P=0.5 boundary): %.4f **\n", threshold_share))
cat(sprintf(
    "   Rule: If PCA sectors hold > %.1f%% of total q0 => use PCA, else use DIIM\n\n",
    threshold_share * 100
))

# Also fit with q0_mean_ratio for an alternative interpretable rule
logit_ratio <- glm(pca_wins_int ~ q0_mean_ratio,
    data = mc_df, family = binomial
)

coefs_ratio <- coef(logit_ratio)
threshold_ratio <- -coefs_ratio[1] / coefs_ratio[2]
cat(sprintf("** q0_mean_ratio threshold (P=0.5 boundary): %.4f **\n", threshold_ratio))
cat(sprintf(
    "   Rule: If mean(q0_PCA) / mean(q0_nonPCA) > %.4f => use PCA, else DIIM\n\n",
    threshold_ratio
))

# Also fit with q0_top5_overlap
logit_overlap <- glm(pca_wins_int ~ q0_top5_overlap,
    data = mc_df, family = binomial
)

coefs_overlap <- coef(logit_overlap)
threshold_overlap <- -coefs_overlap[1] / coefs_overlap[2]
cat(sprintf("** q0_top5_overlap threshold (P=0.5 boundary): %.4f **\n", threshold_overlap))
cat(sprintf(
    "   Rule: If >= %d of the top-5 q0 sectors are PCA sectors => use PCA\n\n",
    ceiling(threshold_overlap)
))


# ─────────────────────────────────────────────────────────────────────────────
# PART 3: ROC Curve and Optimal Thresholds
# ─────────────────────────────────────────────────────────────────────────────
cat("--- Part 3: ROC Analysis & Optimal Thresholds ---\n")

# Manual ROC for q0_pca_share
compute_roc <- function(scores, labels, n_thresholds = 500) {
    thresholds <- seq(min(scores), max(scores), length.out = n_thresholds)
    roc_data <- data.frame(
        threshold = numeric(n_thresholds),
        tpr = numeric(n_thresholds),
        fpr = numeric(n_thresholds),
        accuracy = numeric(n_thresholds),
        precision = numeric(n_thresholds),
        f1 = numeric(n_thresholds)
    )

    for (i in seq_along(thresholds)) {
        t <- thresholds[i]
        pred <- scores >= t
        tp <- sum(pred & labels)
        fp <- sum(pred & !labels)
        fn <- sum(!pred & labels)
        tn <- sum(!pred & !labels)

        tpr <- tp / max(tp + fn, 1)
        fpr <- fp / max(fp + tn, 1)
        acc <- (tp + tn) / length(labels)
        prec <- tp / max(tp + fp, 1)
        f1 <- if (prec + tpr > 0) 2 * prec * tpr / (prec + tpr) else 0

        roc_data[i, ] <- c(t, tpr, fpr, acc, prec, f1)
    }
    return(roc_data)
}

# AUC helper (trapezoidal)
compute_auc <- function(roc_data) {
    ord <- order(roc_data$fpr)
    fpr_sorted <- roc_data$fpr[ord]
    tpr_sorted <- roc_data$tpr[ord]
    auc <- sum(diff(fpr_sorted) * (head(tpr_sorted, -1) + tail(tpr_sorted, -1)) / 2)
    return(abs(auc))
}

# --- q0_pca_share ROC ---
roc_share <- compute_roc(mc_df$q0_pca_share, mc_df$pca_wins)
auc_share <- compute_auc(roc_share)

# Youden's J optimal threshold
roc_share$youden_j <- roc_share$tpr - roc_share$fpr
best_idx_share <- which.max(roc_share$youden_j)
optimal_threshold_share <- roc_share$threshold[best_idx_share]

cat(sprintf(
    "q0_pca_share: AUC=%.4f, Optimal threshold (Youden)=%.4f\n",
    auc_share, optimal_threshold_share
))
cat(sprintf(
    "  At optimal: TPR=%.3f, FPR=%.3f, Accuracy=%.3f, F1=%.3f\n",
    roc_share$tpr[best_idx_share], roc_share$fpr[best_idx_share],
    roc_share$accuracy[best_idx_share], roc_share$f1[best_idx_share]
))

# Max accuracy threshold
best_acc_idx <- which.max(roc_share$accuracy)
cat(sprintf(
    "  Max accuracy threshold: %.4f (accuracy=%.3f)\n",
    roc_share$threshold[best_acc_idx], roc_share$accuracy[best_acc_idx]
))

# --- q0_mean_ratio ROC ---
roc_ratio <- compute_roc(mc_df$q0_mean_ratio, mc_df$pca_wins)
auc_ratio <- compute_auc(roc_ratio)

roc_ratio$youden_j <- roc_ratio$tpr - roc_ratio$fpr
best_idx_ratio <- which.max(roc_ratio$youden_j)
optimal_threshold_ratio <- roc_ratio$threshold[best_idx_ratio]

cat(sprintf(
    "\nq0_mean_ratio: AUC=%.4f, Optimal threshold (Youden)=%.4f\n",
    auc_ratio, optimal_threshold_ratio
))
cat(sprintf(
    "  At optimal: TPR=%.3f, FPR=%.3f, Accuracy=%.3f, F1=%.3f\n",
    roc_ratio$tpr[best_idx_ratio], roc_ratio$fpr[best_idx_ratio],
    roc_ratio$accuracy[best_idx_ratio], roc_ratio$f1[best_idx_ratio]
))

# --- q0_top5_overlap ROC ---
roc_overlap <- compute_roc(mc_df$q0_top5_overlap, mc_df$pca_wins)
auc_overlap <- compute_auc(roc_overlap)

roc_overlap$youden_j <- roc_overlap$tpr - roc_overlap$fpr
best_idx_overlap <- which.max(roc_overlap$youden_j)
optimal_threshold_ovl <- roc_overlap$threshold[best_idx_overlap]

cat(sprintf(
    "\nq0_top5_overlap: AUC=%.4f, Optimal threshold (Youden)=%.4f\n",
    auc_overlap, optimal_threshold_ovl
))
cat(sprintf(
    "  At optimal: TPR=%.3f, FPR=%.3f, Accuracy=%.3f, F1=%.3f\n",
    roc_overlap$tpr[best_idx_overlap], roc_overlap$fpr[best_idx_overlap],
    roc_overlap$accuracy[best_idx_overlap], roc_overlap$f1[best_idx_overlap]
))


# ─────────────────────────────────────────────────────────────────────────────
# PART 4: Systematic Sweep of q0_pca_share
#
# Instead of random q0, we systematically control q0_pca_share to map out
# the exact transition zone.
# ─────────────────────────────────────────────────────────────────────────────
cat("\n--- Part 4: Systematic q0_pca_share Sweep ---\n")

target_shares <- seq(0.10, 0.70, by = 0.01)
n_reps <- 50 # repetitions per share level (random within constraint)

sweep_results <- data.frame(
    target_share = numeric(),
    actual_share = numeric(),
    pca_wins = logical(),
    pca_advantage = numeric(),
    stringsAsFactors = FALSE
)

set.seed(456)
for (share in target_shares) {
    for (rep in 1:n_reps) {
        # Create random q0 with a target PCA share
        # Total q0 is drawn uniformly
        total_q0 <- runif(1, min = sum(q0_base) * 0.5, max = sum(q0_base) * 1.5)

        pca_total <- total_q0 * share
        non_pca_total <- total_q0 * (1 - share)

        # Distribute within PCA/non-PCA using Dirichlet-like (random splits)
        pca_weights <- runif(length(pca_sectors))
        pca_weights <- pca_weights / sum(pca_weights)

        non_pca_weights <- runif(length(non_pca_sectors))
        non_pca_weights <- non_pca_weights / sum(non_pca_weights)

        q0_test <- numeric(n_sectors)
        q0_test[pca_sectors] <- pca_total * pca_weights
        q0_test[non_pca_sectors] <- non_pca_total * non_pca_weights
        q0_test[q0_test < 1e-6] <- 1e-6

        actual_share <- sum(q0_test[pca_sectors]) / sum(q0_test)

        res <- compare_methods(q0_test, A_star, c_star_base, x,
            lockdown_duration = 55, total_duration = 751
        )

        sweep_results <- rbind(sweep_results, data.frame(
            target_share = share,
            actual_share = actual_share,
            pca_wins = res$pca_wins,
            pca_advantage = res$reduction_pca - res$reduction_diim,
            stringsAsFactors = FALSE
        ))
    }

    if (round(share * 100) %% 10 == 0) {
        pca_rate <- mean(sweep_results$pca_wins[sweep_results$target_share == share])
        cat(sprintf("  share=%.2f: PCA win rate = %.0f%%\n", share, pca_rate * 100))
    }
}

write.csv(sweep_results, "simulations/results/decision_boundary_sweep.csv", row.names = FALSE)
cat(sprintf("  -> Saved decision_boundary_sweep.csv (%d rows)\n\n", nrow(sweep_results)))

# Compute PCA win rate at each target share
win_rate_by_share <- aggregate(pca_wins ~ target_share, data = sweep_results, FUN = mean)
colnames(win_rate_by_share) <- c("q0_pca_share", "pca_win_rate")

# Find the transition point (where win rate crosses 50%)
transition_idx <- which(win_rate_by_share$pca_win_rate >= 0.5)
if (length(transition_idx) > 0) {
    transition_share <- win_rate_by_share$q0_pca_share[min(transition_idx)]
    cat(sprintf("** Transition point (50%% win rate): q0_pca_share = %.2f **\n", transition_share))
} else {
    cat("  PCA never reaches 50% win rate in the tested range.\n")
    transition_share <- NA
}

# Find the 90% win rate threshold
high_win_idx <- which(win_rate_by_share$pca_win_rate >= 0.9)
if (length(high_win_idx) > 0) {
    high_threshold <- win_rate_by_share$q0_pca_share[min(high_win_idx)]
    cat(sprintf("** 90%% confidence threshold: q0_pca_share = %.2f **\n", high_threshold))
} else {
    high_threshold <- NA
}


# ─────────────────────────────────────────────────────────────────────────────
# PART 5: Plots
# ─────────────────────────────────────────────────────────────────────────────
cat("\n--- Part 5: Generating Plots ---\n")

# Plot 1: PCA win rate vs q0_pca_share (the S-curve)
p1 <- ggplot(win_rate_by_share, aes(x = q0_pca_share, y = pca_win_rate)) +
    geom_line(size = 1.3, color = "#2C3E50") +
    geom_point(size = 2, color = "#2C3E50") +
    geom_hline(yintercept = 0.5, linetype = "dashed", color = "#E74C3C", size = 0.8) +
    {
        if (!is.na(transition_share)) {
            geom_vline(
                xintercept = transition_share, linetype = "dashed",
                color = "#3498DB", size = 0.8
            )
        }
    } +
    {
        if (!is.na(transition_share)) {
            annotate("text",
                x = transition_share + 0.02, y = 0.55,
                label = sprintf("Threshold = %.2f", transition_share),
                color = "#3498DB", fontface = "bold", hjust = 0
            )
        }
    } +
    annotate("rect",
        xmin = 0.10, xmax = ifelse(is.na(transition_share), 0.70, transition_share),
        ymin = -Inf, ymax = Inf, fill = "#E74C3C", alpha = 0.05
    ) +
    annotate("rect",
        xmin = ifelse(is.na(transition_share), 0.70, transition_share),
        xmax = 0.70, ymin = -Inf, ymax = Inf, fill = "#2ECC71", alpha = 0.05
    ) +
    scale_y_continuous(labels = scales::percent, limits = c(0, 1)) +
    scale_x_continuous(labels = scales::percent) +
    labs(
        title = "Decision Boundary: PCA Win Rate vs q0 Share in PCA Sectors",
        subtitle = "How likely is PCA to outperform DIIM given the share of inoperability in PCA sectors?",
        x = expression(paste("q"[0], " PCA Share = ", Sigma, "q"[0]^PCA, " / ", Sigma, "q"[0]^total)),
        y = "P(PCA Wins)"
    ) +
    theme_minimal() +
    theme(
        plot.title = element_text(face = "bold", size = 13),
        plot.subtitle = element_text(size = 10, color = "gray40")
    )

ggsave(file.path(results_dir, "decision_boundary_curve.png"), p1, width = 10, height = 6)
cat("  -> Saved decision_boundary_curve.png\n")


# Plot 2: ROC curves for the three q0 features
roc_share$feature <- "q0_pca_share"
roc_ratio$feature <- "q0_mean_ratio"
roc_overlap$feature <- "q0_top5_overlap"
roc_all <- rbind(
    roc_share[, c("fpr", "tpr", "feature")],
    roc_ratio[, c("fpr", "tpr", "feature")],
    roc_overlap[, c("fpr", "tpr", "feature")]
)

p2 <- ggplot(roc_all, aes(x = fpr, y = tpr, color = feature)) +
    geom_line(size = 1.2) +
    geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "gray60") +
    scale_color_manual(
        values = c(
            "q0_pca_share" = "#E74C3C", "q0_mean_ratio" = "#3498DB",
            "q0_top5_overlap" = "#2ECC71"
        ),
        labels = c(
            sprintf("q0_pca_share (AUC=%.3f)", auc_share),
            sprintf("q0_mean_ratio (AUC=%.3f)", auc_ratio),
            sprintf("q0_top5_overlap (AUC=%.3f)", auc_overlap)
        ),
        name = "Feature"
    ) +
    labs(
        title = "ROC Curves: q0-Based Predictors of PCA Outperformance",
        x = "False Positive Rate",
        y = "True Positive Rate"
    ) +
    theme_minimal() +
    theme(
        plot.title = element_text(face = "bold", size = 13),
        legend.position = "bottom"
    ) +
    coord_equal()

ggsave(file.path(results_dir, "decision_boundary_roc.png"), p2, width = 8, height = 8)
cat("  -> Saved decision_boundary_roc.png\n")


# Plot 3: PCA advantage distribution by q0_pca_share bin
mc_df$share_bin <- cut(mc_df$q0_pca_share,
    breaks = seq(0, 1, by = 0.05),
    include.lowest = TRUE
)

p3 <- ggplot(mc_df, aes(x = share_bin, y = pca_advantage, fill = pca_wins)) +
    geom_boxplot(alpha = 0.7, outlier.alpha = 0.3) +
    geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
    scale_fill_manual(
        values = c("TRUE" = "#2ECC71", "FALSE" = "#E74C3C"),
        labels = c("TRUE" = "PCA Wins", "FALSE" = "DIIM Wins"),
        name = ""
    ) +
    labs(
        title = "PCA Advantage by q0 PCA Share Bin",
        subtitle = "Positive = PCA reduces more loss than DIIM",
        x = "q0 PCA Share Bin",
        y = "PCA Advantage (reduction_PCA - reduction_DIIM)"
    ) +
    theme_minimal() +
    theme(
        plot.title = element_text(face = "bold", size = 13),
        axis.text.x = element_text(angle = 45, hjust = 1, size = 7),
        legend.position = "bottom"
    )

ggsave(file.path(results_dir, "decision_boundary_advantage_boxplot.png"), p3, width = 12, height = 6)
cat("  -> Saved decision_boundary_advantage_boxplot.png\n")


# Plot 4: Scatter plot of q0_pca_share vs q0_mean_ratio colored by winner
p4 <- ggplot(mc_df, aes(x = q0_pca_share, y = q0_mean_ratio, color = pca_wins)) +
    geom_point(alpha = 0.4, size = 1.5) +
    {
        if (!is.na(threshold_share)) {
            geom_vline(
                xintercept = optimal_threshold_share, linetype = "dashed",
                color = "black", size = 0.8
            )
        }
    } +
    scale_color_manual(
        values = c("TRUE" = "#2ECC71", "FALSE" = "#E74C3C"),
        labels = c("TRUE" = "PCA Wins", "FALSE" = "DIIM Wins"),
        name = ""
    ) +
    labs(
        title = "Decision Space: q0_pca_share vs q0_mean_ratio",
        x = "q0 PCA Share",
        y = "q0 Mean Ratio (PCA / non-PCA)"
    ) +
    theme_minimal() +
    theme(
        plot.title = element_text(face = "bold", size = 13),
        legend.position = "bottom"
    )

ggsave(file.path(results_dir, "decision_boundary_scatter.png"), p4, width = 9, height = 7)
cat("  -> Saved decision_boundary_scatter.png\n")


# ─────────────────────────────────────────────────────────────────────────────
# PART 6: Final Summary
# ─────────────────────────────────────────────────────────────────────────────
cat("\n")
cat("══════════════════════════════════════════════════════\n")
cat("  DECISION BOUNDARY SUMMARY (q0 only)\n")
cat("══════════════════════════════════════════════════════\n\n")

cat("BEST SINGLE PREDICTOR: q0_pca_share\n")
cat(sprintf("  AUC: %.4f\n", auc_share))
cat(sprintf("  Logistic regression threshold (P=0.5): %.4f\n", threshold_share))
cat(sprintf("  ROC optimal threshold (Youden's J):    %.4f\n", optimal_threshold_share))
if (!is.na(transition_share)) {
    cat(sprintf("  Systematic sweep transition (50%% win): %.4f\n", transition_share))
}
if (!is.na(high_threshold)) {
    cat(sprintf("  90%% confidence threshold:              %.4f\n", high_threshold))
}

cat("\nALTERNATIVE PREDICTORS:\n")
cat(sprintf("  q0_mean_ratio: AUC=%.4f, threshold=%.4f\n", auc_ratio, optimal_threshold_ratio))
cat(sprintf("  q0_top5_overlap: AUC=%.4f, threshold=%d sectors\n", auc_overlap, ceiling(optimal_threshold_ovl)))

cat("\nDECISION RULES (pick one):\n")
cat(sprintf("  Rule 1: Use PCA if sum(q0[PCA]) / sum(q0) > %.2f\n", optimal_threshold_share))
cat(sprintf(
    "  Rule 2: Use PCA if mean(q0[PCA]) / mean(q0[non-PCA]) > %.2f\n",
    optimal_threshold_ratio
))
cat(sprintf(
    "  Rule 3: Use PCA if >= %d of the top-5 inoperable sectors are PCA sectors\n",
    ceiling(optimal_threshold_ovl)
))

cat("\nFULL MODEL (for highest accuracy):\n")
sig_vars <- summary(logit_full)$coefficients[, 4] < 0.05
sig_names <- names(which(sig_vars[-1])) # exclude intercept
if (length(sig_names) > 0) {
    cat(sprintf("  Significant q0 features: %s\n", paste(sig_names, collapse = ", ")))
} else {
    cat("  No individually significant features in the full model (multicollinearity)\n")
}
cat(sprintf(
    "  Full model AIC: %.1f vs Simple model AIC: %.1f\n",
    AIC(logit_full), AIC(logit_simple)
))

# Accuracy of the simple threshold rule
simple_pred <- mc_df$q0_pca_share >= optimal_threshold_share
simple_accuracy <- mean(simple_pred == mc_df$pca_wins)
cat(sprintf("\nSIMPLE RULE ACCURACY: %.1f%%\n", simple_accuracy * 100))

# Verify on original data
original_share <- sum(q0_base[pca_sectors]) / sum(q0_base)
cat(sprintf("\nVERIFICATION (original COVID data):\n"))
cat(sprintf("  Original q0_pca_share = %.4f\n", original_share))
cat(sprintf("  Threshold = %.4f\n", optimal_threshold_share))
cat(sprintf("  Prediction: %s\n", ifelse(original_share >= optimal_threshold_share,
    "Use PCA (correct!)", "Use DIIM"
)))

cat("\n══════════════════════════════════════════════════════\n")

# Save summary to text file
sink("simulations/results/decision_boundary_summary.txt")
cat("DECISION BOUNDARY SUMMARY\n")
cat("=========================\n\n")
cat(sprintf("Best predictor: q0_pca_share (AUC = %.4f)\n", auc_share))
cat(sprintf("Optimal threshold (Youden's J): %.4f\n", optimal_threshold_share))
cat(sprintf("Logistic regression threshold: %.4f\n", threshold_share))
cat(sprintf("Simple rule accuracy: %.1f%%\n\n", simple_accuracy * 100))

cat("Decision Rules:\n")
cat(sprintf("  1. Use PCA if sum(q0[PCA_sectors]) / sum(q0) > %.4f\n", optimal_threshold_share))
cat(sprintf("  2. Use PCA if mean(q0[PCA]) / mean(q0[non-PCA]) > %.4f\n", optimal_threshold_ratio))
cat(sprintf("  3. Use PCA if >= %d of top-5 q0 sectors are PCA sectors\n\n", ceiling(optimal_threshold_ovl)))

cat("Full logistic regression model:\n")
print(summary(logit_full))
cat("\nSimple logistic regression model:\n")
print(summary(logit_simple))
cat("\nPCA win rate by q0_pca_share:\n")
print(win_rate_by_share)
sink()

cat("  -> Saved decision_boundary_summary.txt\n")
cat("\nDone!\n")
