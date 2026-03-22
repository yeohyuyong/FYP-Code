# Set working directory to project root
setwd("..")

library(openxlsx)
# Helper: source R code from a Jupyter notebook (.ipynb)
source_notebook <- function(nb_path) {
  nb <- jsonlite::fromJSON(nb_path, simplifyVector = FALSE)
  code_cells <- Filter(function(cell) cell$cell_type == "code", nb$cells)
  code_lines <- unlist(lapply(code_cells, function(cell) {
    paste(unlist(cell$source), collapse = "")
  }))
  all_code <- paste(code_lines, collapse = "\n")
  eval(parse(text = all_code), envir = globalenv())
}

source_notebook("functions.ipynb")
dir.create("results", showWarnings = FALSE, recursive = TRUE)

data <- download_data()
A <- data$A
x <- data$x
q0_base <- data$q0
c_star_base <- data$c_star
A_star <- data$A_star

n_sectors <- length(q0_base)
pca_key_sectors <- c(3, 2, 5, 8, 10)

cat("=== Base Data Summary ===\n")
cat("Number of sectors:", n_sectors, "\n")
cat("PCA key sectors:", pca_key_sectors, "\n")
cat("q0 range:", range(q0_base), "\n")
cat("c_star range:", range(c_star_base), "\n")
cat("x range:", range(x), "\n")

compare_methods <- function(q0, A_star, c_star, x,
                            lockdown_duration, total_duration,
                            pca_sectors = pca_key_sectors) {
    model_base <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration)
    EL_base <- model_base$EL_evolution
    max_econ_loss <- apply(EL_base, 1, max)
    diim_sectors <- order(max_econ_loss, decreasing = TRUE)[1:5]
    model_diim <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = diim_sectors)
    model_pca <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = pca_sectors)
    loss_base <- model_base$total_economic_loss
    loss_diim <- model_diim$total_economic_loss
    loss_pca <- model_pca$total_economic_loss
    reduction_diim <- loss_base - loss_diim
    reduction_pca <- loss_base - loss_pca
    pca_wins <- reduction_pca > reduction_diim
    overlap <- length(intersect(pca_sectors, diim_sectors))
    return(list(
        loss_base = loss_base, loss_diim = loss_diim, loss_pca = loss_pca,
        reduction_diim = reduction_diim, reduction_pca = reduction_pca,
        pca_wins = pca_wins, diim_sectors = paste(diim_sectors, collapse = ","),
        overlap = overlap
    ))
}

gini <- function(x) {
    x <- sort(abs(x))
    n <- length(x)
    if (sum(x) == 0) return(0)
    index <- 1:n
    return((2 * sum(index * x) / (n * sum(x))) - (n + 1) / n)
}

cat("=== Experiment 1: c_star Magnitude Scaling ===\n")

multipliers <- seq(0, 2, by = 0.1)
exp1_results <- data.frame(
    multiplier = numeric(), loss_base = numeric(), loss_diim = numeric(),
    loss_pca = numeric(), reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), diim_sectors = character(), overlap = integer(),
    stringsAsFactors = FALSE
)

for (m in multipliers) {
    c_star_scaled <- c_star_base * m
    res <- compare_methods(q0_base, A_star, c_star_scaled, x,
        lockdown_duration = 55, total_duration = 751)
    exp1_results <- rbind(exp1_results, data.frame(
        multiplier = m, loss_base = res$loss_base, loss_diim = res$loss_diim,
        loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
        reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
        diim_sectors = res$diim_sectors, overlap = res$overlap,
        stringsAsFactors = FALSE))
    cat(sprintf("  Multiplier %.1f: PCA wins = %s\n", m, res$pca_wins))
}

write.csv(exp1_results, "results/exp1_cstar_magnitude.csv", row.names = FALSE)
cat("  -> Saved\n")

cat("=== Experiment 2: c_star Concentration vs Spread ===\n")

total_c_star <- sum(c_star_base)

create_concentrated <- function(total, n, top_k = 3) {
    cs <- rep(0, n)
    top_idx <- order(c_star_base, decreasing = TRUE)[1:top_k]
    cs[top_idx] <- total / top_k
    return(cs)
}

create_moderate <- function(total, n, spread_k = 7) {
    cs <- rep(0, n)
    top_idx <- order(c_star_base, decreasing = TRUE)[1:spread_k]
    cs[top_idx] <- total / spread_k
    return(cs)
}

create_uniform <- function(total, n) { return(rep(total / n, n)) }

create_inverse <- function(total, n) {
    cs <- rep(0, n)
    sorted_idx <- order(c_star_base, decreasing = FALSE)
    weights <- rev(sort(c_star_base))
    if (sum(weights) > 0) { cs[sorted_idx] <- weights / sum(weights) * total }
    else { cs <- rep(total / n, n) }
    return(cs)
}

distribution_patterns <- list(
    "original" = c_star_base,
    "concentrated_3" = create_concentrated(total_c_star, n_sectors, 3),
    "concentrated_5" = create_concentrated(total_c_star, n_sectors, 5),
    "moderate_7" = create_moderate(total_c_star, n_sectors, 7),
    "moderate_10" = create_moderate(total_c_star, n_sectors, 10),
    "uniform" = create_uniform(total_c_star, n_sectors),
    "inverse" = create_inverse(total_c_star, n_sectors)
)

exp2_results <- data.frame(
    pattern = character(), loss_base = numeric(), loss_diim = numeric(),
    loss_pca = numeric(), reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), diim_sectors = character(), overlap = integer(),
    c_star_gini = numeric(), stringsAsFactors = FALSE
)

for (pattern_name in names(distribution_patterns)) {
    cs <- distribution_patterns[[pattern_name]]
    res <- compare_methods(q0_base, A_star, cs, x, lockdown_duration = 55, total_duration = 751)
    g <- gini(cs)
    exp2_results <- rbind(exp2_results, data.frame(
        pattern = pattern_name, loss_base = res$loss_base, loss_diim = res$loss_diim,
        loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
        reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
        diim_sectors = res$diim_sectors, overlap = res$overlap,
        c_star_gini = g, stringsAsFactors = FALSE))
    cat(sprintf("  Pattern '%s' (Gini=%.3f): PCA wins = %s\n", pattern_name, g, res$pca_wins))
}

write.csv(exp2_results, "results/exp2_cstar_distribution.csv", row.names = FALSE)
cat("  -> Saved\n")

cat("=== Experiment 3: q0 Uniformity ===\n")

total_q0 <- sum(q0_base)

q0_patterns <- list(
    "original" = q0_base,
    "uniform" = rep(mean(q0_base), n_sectors),
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
    overlap = integer(), q0_gini = numeric(), stringsAsFactors = FALSE
)

for (q0_name in names(q0_patterns)) {
    for (cs_mult in c_star_levels) {
        q0_cur <- q0_patterns[[q0_name]]
        cs_cur <- c_star_base * cs_mult
        res <- compare_methods(q0_cur, A_star, cs_cur, x, lockdown_duration = 55, total_duration = 751)
        g_q0 <- gini(q0_cur)
        exp3_results <- rbind(exp3_results, data.frame(
            q0_pattern = q0_name, c_star_multiplier = cs_mult,
            loss_base = res$loss_base, loss_diim = res$loss_diim,
            loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
            reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
            diim_sectors = res$diim_sectors, overlap = res$overlap,
            q0_gini = g_q0, stringsAsFactors = FALSE))
        cat(sprintf("  q0='%s', c_star_mult=%.1f: PCA wins = %s\n", q0_name, cs_mult, res$pca_wins))
    }
}

write.csv(exp3_results, "results/exp3_q0_uniformity.csv", row.names = FALSE)
cat("  -> Saved\n")

cat("=== Experiment 4: Lockdown Duration x c_star Interaction ===\n")

lockdown_vals <- seq(5, 100, by = 5)
cstar_mult_vals <- seq(0, 2, by = 0.2)

exp4_results <- data.frame(
    lockdown_duration = integer(), c_star_multiplier = numeric(),
    loss_base = numeric(), loss_diim = numeric(), loss_pca = numeric(),
    reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), overlap = integer(), stringsAsFactors = FALSE
)

total_combos <- length(lockdown_vals) * length(cstar_mult_vals)
combo_count <- 0

for (ld in lockdown_vals) {
    for (cm in cstar_mult_vals) {
        cs <- c_star_base * cm
        res <- compare_methods(q0_base, A_star, cs, x,
            lockdown_duration = ld, total_duration = 751)
        exp4_results <- rbind(exp4_results, data.frame(
            lockdown_duration = ld, c_star_multiplier = cm,
            loss_base = res$loss_base, loss_diim = res$loss_diim,
            loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
            reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
            overlap = res$overlap, stringsAsFactors = FALSE))
        combo_count <- combo_count + 1
        if (combo_count %% 50 == 0) cat(sprintf("  Progress: %d / %d\n", combo_count, total_combos))
    }
}

write.csv(exp4_results, "results/exp4_lockdown_cstar_grid.csv", row.names = FALSE)
cat(sprintf("  -> Saved (%d rows)\n", nrow(exp4_results)))

cat("=== Experiment 5: Monte Carlo Simulation ===\n")

set.seed(42)
n_mc <- 500
max_c_star <- max(abs(c_star_base)) * 2
max_q0 <- max(q0_base) * 1.5

exp5_results <- data.frame(
    sim_id = integer(), c_star_mean = numeric(), c_star_sd = numeric(),
    c_star_max = numeric(), c_star_gini = numeric(), c_star_sum = numeric(),
    q0_mean = numeric(), q0_sd = numeric(), q0_max = numeric(),
    q0_gini = numeric(), q0_sum = numeric(), cstar_q0_cor = numeric(),
    c_star_pca_share = numeric(), q0_pca_share = numeric(),
    lockdown_duration = integer(), loss_base = numeric(),
    loss_diim = numeric(), loss_pca = numeric(),
    reduction_diim = numeric(), reduction_pca = numeric(),
    pca_wins = logical(), overlap = integer(), stringsAsFactors = FALSE
)

for (i in 1:n_mc) {
    cs_rand <- runif(n_sectors, min = 0, max = max_c_star)
    q0_rand <- runif(n_sectors, min = 1e-6, max = max_q0)
    ld_rand <- sample(10:80, 1)
    res <- compare_methods(q0_rand, A_star, cs_rand, x,
        lockdown_duration = ld_rand, total_duration = 751)
    cs_gini <- gini(cs_rand)
    q0_gini_val <- gini(q0_rand)
    cstar_q0_corr <- if (sd(cs_rand) > 0 && sd(q0_rand) > 0) cor(cs_rand, q0_rand) else 0
    cs_pca_share <- sum(cs_rand[pca_key_sectors]) / sum(cs_rand)
    q0_pca_share <- sum(q0_rand[pca_key_sectors]) / sum(q0_rand)
    exp5_results <- rbind(exp5_results, data.frame(
        sim_id = i, c_star_mean = mean(cs_rand), c_star_sd = sd(cs_rand),
        c_star_max = max(cs_rand), c_star_gini = cs_gini, c_star_sum = sum(cs_rand),
        q0_mean = mean(q0_rand), q0_sd = sd(q0_rand), q0_max = max(q0_rand),
        q0_gini = q0_gini_val, q0_sum = sum(q0_rand),
        cstar_q0_cor = cstar_q0_corr, c_star_pca_share = cs_pca_share,
        q0_pca_share = q0_pca_share, lockdown_duration = ld_rand,
        loss_base = res$loss_base, loss_diim = res$loss_diim,
        loss_pca = res$loss_pca, reduction_diim = res$reduction_diim,
        reduction_pca = res$reduction_pca, pca_wins = res$pca_wins,
        overlap = res$overlap, stringsAsFactors = FALSE))
    if (i %% 50 == 0) {
        pca_win_rate <- mean(exp5_results$pca_wins[1:i])
        cat(sprintf("  MC sim %d / %d (PCA win rate: %.1f%%)\n", i, n_mc, pca_win_rate * 100))
    }
}

write.csv(exp5_results, "results/exp5_monte_carlo.csv", row.names = FALSE)
cat(sprintf("  -> Saved (%d rows)\n", nrow(exp5_results)))

exp5_results$pca_wins_int <- as.integer(exp5_results$pca_wins)

logit_model <- glm(
    pca_wins_int ~ c_star_mean + c_star_sd + c_star_gini +
        q0_mean + q0_sd + q0_gini +
        cstar_q0_cor + c_star_pca_share + q0_pca_share +
        lockdown_duration + overlap,
    data = exp5_results, family = binomial
)

cat("\nLogistic Regression Summary:\n")
print(summary(logit_model))
cat("\nOverall PCA Win Rate:", mean(exp5_results$pca_wins_int) * 100, "%\n")
