if (!file.exists("functions.R")) setwd("..")

library(openxlsx)

source("functions.R")
dir.create("simulations/results", showWarnings = FALSE, recursive = TRUE)

data <- download_data()
A      <- data$A
x      <- data$x
q0_base    <- data$q0
c_star_base <- data$c_star
A_star <- data$A_star

n_sectors <- length(q0_base)

# PCA x xi key sectors
pca_results <- pca_rank_sectors(A, x, n_pcs = 2)
pca_key_sectors <- pca_results$ranked_sectors[1:5]

cat("Number of sectors:", n_sectors, "\n")
cat("PCA x xi key sectors:", pca_key_sectors, "\n")
cat("q0 range:", range(q0_base), "\n")
cat("c_star range:", range(c_star_base), "\n")

multipliers <- seq(0, 2, by = 0.1)

exp1_results <- data.frame(
    multiplier = numeric(), loss_base = numeric(), loss_diim = numeric(),
    loss_pca = numeric(), reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), diim_sectors = character(), overlap = integer(),
    stringsAsFactors = FALSE
)

for (c_star_multiplier in multipliers) {
    c_star_scaled <- c_star_base * c_star_multiplier
    res <- compare_methods(q0_base, A_star, c_star_scaled, x,
        lockdown_duration = 55, total_duration = 751, pca_sectors = pca_key_sectors)

    exp1_results <- rbind(exp1_results, data.frame(
        multiplier = c_star_multiplier, loss_base = res$loss_base, loss_diim = res$loss_diim,
        loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
        reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
        diim_sectors = res$diim_sectors, overlap = res$overlap,
        stringsAsFactors = FALSE))

    cat(sprintf("  Multiplier %.1f: PCA wins = %s\n", c_star_multiplier, res$pca_wins))
}

write.csv(exp1_results, "simulations/results/exp1_cstar_magnitude.csv", row.names = FALSE)
cat("Saved exp1_cstar_magnitude.csv\n")

total_c_star <- sum(c_star_base)

# --- Exp 2: c* distribution patterns ---
# concentrated_k: all c_star to top-k sectors; uniform: equal spread; inverse: reversed ranking

make_concentrated <- function(top_k) {
    cs <- rep(0, n_sectors)
    top_idx <- order(c_star_base, decreasing = TRUE)[1:top_k]
    cs[top_idx] <- total_c_star / top_k
    return(cs)
}

inverse_cs <- {
    cs <- rep(0, n_sectors)
    sorted_idx <- order(c_star_base, decreasing = FALSE)
    weights <- rev(sort(c_star_base))
    if (sum(weights) > 0) cs[sorted_idx] <- weights / sum(weights) * total_c_star
    else cs <- rep(total_c_star / n_sectors, n_sectors) # fallback if all sectors are zero
    cs
}

distribution_patterns <- list(
    "original"        = c_star_base,
    "concentrated_3"  = make_concentrated(3),
    "concentrated_5"  = make_concentrated(5),
    "moderate_7"      = make_concentrated(7),
    "moderate_10"     = make_concentrated(10),
    "uniform"         = rep(total_c_star / n_sectors, n_sectors),
    "inverse"         = inverse_cs
)

exp2_results <- data.frame(
    pattern = character(), loss_base = numeric(), loss_diim = numeric(),
    loss_pca = numeric(), reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), diim_sectors = character(), overlap = integer(),
    stringsAsFactors = FALSE
)

for (pattern_name in names(distribution_patterns)) {
    cs <- distribution_patterns[[pattern_name]]
    res <- compare_methods(q0_base, A_star, cs, x,
        lockdown_duration = 55, total_duration = 751, pca_sectors = pca_key_sectors)

    exp2_results <- rbind(exp2_results, data.frame(
        pattern = pattern_name, loss_base = res$loss_base, loss_diim = res$loss_diim,
        loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
        reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
        diim_sectors = res$diim_sectors, overlap = res$overlap,
        stringsAsFactors = FALSE))

    cat(sprintf("  Pattern '%s': PCA wins = %s\n", pattern_name, res$pca_wins))
}

write.csv(exp2_results, "simulations/results/exp2_cstar_distribution.csv", row.names = FALSE)
cat("Saved exp2_cstar_distribution.csv\n")

total_q0 <- sum(q0_base)

# --- Exp 3: q0 distribution patterns ---
q0_patterns <- list(
    "original" = q0_base,
    "uniform"  = rep(mean(q0_base), n_sectors),
    "concentrated_top3" = {
        q <- rep(total_q0 * 0.2 / (n_sectors - 3), n_sectors)
        top_idx <- order(q0_base, decreasing = TRUE)[1:3]
        q[top_idx] <- total_q0 * 0.8 / 3
        q
    },
    "concentrated_bottom3" = {
        q <- rep(total_q0 * 0.2 / (n_sectors - 3), n_sectors)
        bot_idx <- order(q0_base, decreasing = FALSE)[1:3]
        q[bot_idx] <- total_q0 * 0.8 / 3
        q
    }
)

c_star_levels <- c(0.5, 1.0, 1.5)

exp3_results <- data.frame(
    q0_pattern = character(), c_star_multiplier = numeric(),
    loss_base = numeric(), loss_diim = numeric(), loss_pca = numeric(),
    reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), diim_sectors = character(),
    overlap = integer(), stringsAsFactors = FALSE
)

for (q0_name in names(q0_patterns)) {
    for (c_star_multiplier in c_star_levels) {
        current_q0 <- q0_patterns[[q0_name]]
        current_c_star <- c_star_base * c_star_multiplier
        res <- compare_methods(current_q0, A_star, current_c_star, x,
            lockdown_duration = 55, total_duration = 751, pca_sectors = pca_key_sectors)

        exp3_results <- rbind(exp3_results, data.frame(
            q0_pattern = q0_name, c_star_multiplier = c_star_multiplier,
            loss_base = res$loss_base, loss_diim = res$loss_diim,
            loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
            reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
            diim_sectors = res$diim_sectors, overlap = res$overlap,
            stringsAsFactors = FALSE))

        cat(sprintf("  q0='%s', c_star_mult=%.1f: PCA wins = %s\n",
            q0_name, c_star_multiplier, res$pca_wins))
    }
}

write.csv(exp3_results, "simulations/results/exp3_q0_uniformity.csv", row.names = FALSE)
cat("Saved exp3_q0_uniformity.csv\n")

lockdown_vals  <- seq(5, 100, by = 5)
cstar_mult_vals <- seq(0, 2, by = 0.2)

exp4_results <- data.frame(
    lockdown_duration = integer(), c_star_multiplier = numeric(),
    loss_base = numeric(), loss_diim = numeric(), loss_pca = numeric(),
    reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), overlap = integer(), stringsAsFactors = FALSE
)

total_combos <- length(lockdown_vals) * length(cstar_mult_vals)
combo_count <- 0

for (duration in lockdown_vals) {
    for (multiplier in cstar_mult_vals) {
        current_c_star <- c_star_base * multiplier
        res <- compare_methods(q0_base, A_star, current_c_star, x,
            lockdown_duration = duration, total_duration = 751, pca_sectors = pca_key_sectors)

        exp4_results <- rbind(exp4_results, data.frame(
            lockdown_duration = duration, c_star_multiplier = multiplier,
            loss_base = res$loss_base, loss_diim = res$loss_diim,
            loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
            reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
            overlap = res$overlap, stringsAsFactors = FALSE))

        combo_count <- combo_count + 1
        if (combo_count %% 50 == 0) cat(sprintf("  Progress: %d / %d\n", combo_count, total_combos))
    }
}

write.csv(exp4_results, "simulations/results/exp4_lockdown_cstar_grid.csv", row.names = FALSE)
cat(sprintf("Saved exp4_lockdown_cstar_grid.csv (%d rows)\n", nrow(exp4_results)))

set.seed(42)
n_mc <- 500
max_c_star <- max(abs(c_star_base)) * 2
max_q0 <- max(q0_base) * 1.5

exp5_results <- data.frame(
    sim_id = integer(), c_star_mean = numeric(), c_star_sd = numeric(),
    c_star_max = numeric(), c_star_sum = numeric(),
    q0_mean = numeric(), q0_sd = numeric(), q0_max = numeric(),
    q0_sum = numeric(), cstar_q0_cor = numeric(),
    c_star_pca_share = numeric(), q0_pca_share = numeric(),
    lockdown_duration = integer(), loss_base = numeric(),
    loss_diim = numeric(), loss_pca = numeric(),
    reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), overlap = integer(), stringsAsFactors = FALSE
)

for (i in 1:n_mc) {
    random_c_star <- runif(n_sectors, min = 0, max = max_c_star)
    random_q0 <- runif(n_sectors, min = 1e-6, max = max_q0)
    random_lockdown_duration <- sample(10:80, 1)

    res <- compare_methods(random_q0, A_star, random_c_star, x,
        lockdown_duration = random_lockdown_duration, total_duration = 751, pca_sectors = pca_key_sectors)

    # Summary statistics as predictors for logistic regression
    cstar_q0_cor <- if (sd(random_c_star) > 0 && sd(random_q0) > 0) cor(random_c_star, random_q0) else 0
    c_star_pca_share <- sum(random_c_star[pca_key_sectors]) / sum(random_c_star)
    q0_pca_share <- sum(random_q0[pca_key_sectors]) / sum(random_q0)

    exp5_results <- rbind(exp5_results, data.frame(
        sim_id = i, c_star_mean = mean(random_c_star), c_star_sd = sd(random_c_star),
        c_star_max = max(random_c_star), c_star_sum = sum(random_c_star),
        q0_mean = mean(random_q0), q0_sd = sd(random_q0), q0_max = max(random_q0),
        q0_sum = sum(random_q0),
        cstar_q0_cor = cstar_q0_cor, c_star_pca_share = c_star_pca_share,
        q0_pca_share = q0_pca_share, lockdown_duration = random_lockdown_duration,
        loss_base = res$loss_base, loss_diim = res$loss_diim,
        loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
        reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
        overlap = res$overlap, stringsAsFactors = FALSE))

    if (i %% 50 == 0) {
        pca_win_rate <- mean(exp5_results$pca_wins[1:i])
        cat(sprintf("  MC sim %d / %d (PCA win rate: %.1f%%)\n", i, n_mc, pca_win_rate * 100))
    }
}

write.csv(exp5_results, "simulations/results/exp5_monte_carlo.csv", row.names = FALSE)
cat(sprintf("Saved exp5_monte_carlo.csv (%d rows)\n", nrow(exp5_results)))

exp5_results$pca_wins_int <- as.integer(exp5_results$pca_wins)

logit_model <- glm(
    pca_wins_int ~ c_star_mean + c_star_sd +
        q0_mean + q0_sd +
        cstar_q0_cor + c_star_pca_share + q0_pca_share +
        lockdown_duration + overlap,
    data = exp5_results, family = binomial
)

cat("Logistic Regression Summary:\n")
print(summary(logit_model))
cat("\nOverall PCA Win Rate:", mean(exp5_results$pca_wins_int) * 100, "%\n")
