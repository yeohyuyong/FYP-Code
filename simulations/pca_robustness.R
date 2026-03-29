# Tests whether n_pcs=2 is optimal for PCA x xi method.
# 500-trial MC for n_pcs in {1, 2, 3, 4} at k=5.

if (!file.exists("functions.R")) setwd("..")

library(openxlsx)
library(ggplot2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

n_pcs_values <- c(1, 2, 3, 4)
k <- 5
n_mc <- 500

# PCA ranking with variable n_pcs (unlike functions.R which fixes n_pcs=2)
pca_rank_with_npcs <- function(A, x, n_pcs) {
    n <- nrow(A)
    I_minus_A <- diag(n) - A
    L <- solve(I_minus_A)
    H <- A %*% L
    eig <- eigen(H)
    loadings <- Re(eig$vectors[, 1:min(n_pcs, n), drop = FALSE])
    distances <- sqrt(rowSums(loadings^2))
    weighted_distances <- distances * as.vector(x)
    ranked_sectors <- order(weighted_distances, decreasing = TRUE)

    eigenvalues <- Re(eig$values)
    var_explained <- sum(abs(eigenvalues[1:min(n_pcs, n)])) / sum(abs(eigenvalues))

    return(list(ranked_sectors = ranked_sectors,
                var_explained = var_explained))
}

run_pca_robustness <- function(scenario_name, data_loader,
                                lockdown_duration, total_duration,
                                days_in_year) {
    cat(sprintf("\n--- %s PCA Robustness (k=%d) ---\n", scenario_name, k))
    data <- data_loader()
    A <- data$A; x <- data$x; c_star <- data$c_star; A_star <- data$A_star
    q0_base <- data$q0; q0_base[q0_base == 0] <- 1e-8
    num_sectors <- length(q0_base)

    pca_rankings <- list()
    var_explained_vals <- numeric(length(n_pcs_values))
    for (i in seq_along(n_pcs_values)) {
        np <- n_pcs_values[i]
        result <- pca_rank_with_npcs(A, x, np)
        pca_rankings[[i]] <- result$ranked_sectors
        var_explained_vals[i] <- result$var_explained
        cat(sprintf("  n_pcs=%d: variance explained=%.3f, top-%d=[%s]\n",
            np, result$var_explained, k,
            paste(result$ranked_sectors[1:k], collapse = ", ")))
    }

    set.seed(42)
    all_results <- list()

    for (npc_idx in seq_along(n_pcs_values)) {
        np <- n_pcs_values[npc_idx]
        pca_topk <- pca_rankings[[npc_idx]][1:k]
        cat(sprintf("  Testing n_pcs=%d: ", np))

        ratios <- numeric(0)
        valid_trials <- 0

        for (trial in 1:n_mc) {
            noise <- exp(rnorm(num_sectors, mean = 0, sd = 0.5))
            random_q0 <- pmin(pmax(q0_base * noise, 1e-6), 1)

            base_model <- DIIM(random_q0, A_star, c_star, x,
                               lockdown_duration, total_duration,
                               days_in_year = days_in_year)
            base_loss <- base_model$total_economic_loss
            if (base_loss < 1e-6) next

            max_el <- apply(base_model$EL_evolution, 1, max)
            diim_topk <- order(max_el, decreasing = TRUE)[1:k]
            diim_model <- DIIM(random_q0, A_star, c_star, x,
                               lockdown_duration, total_duration,
                               key_sectors = diim_topk,
                               days_in_year = days_in_year)
            diim_reduction <- base_loss - diim_model$total_economic_loss
            if (diim_reduction < 1e-10) next
            valid_trials <- valid_trials + 1

            pca_model <- DIIM(random_q0, A_star, c_star, x,
                              lockdown_duration, total_duration,
                              key_sectors = pca_topk,
                              days_in_year = days_in_year)
            pca_reduction <- base_loss - pca_model$total_economic_loss
            ratios <- c(ratios, pca_reduction / diim_reduction)
        }

        cat(sprintf("%d valid trials\n", valid_trials))

        all_results[[length(all_results) + 1]] <- data.frame(
            scenario = scenario_name,
            n_pcs = np,
            var_explained = var_explained_vals[npc_idx],
            mean_ratio = mean(ratios),
            sd_ratio = sd(ratios),
            close_rate = mean(ratios >= 0.80),
            n_valid = length(ratios),
            stringsAsFactors = FALSE
        )
    }

    return(bind_rows(all_results))
}

# --- Run both scenarios ---
covid_results <- run_pca_robustness("COVID-19", download_data, 55, 751, 366)
manpower_results <- run_pca_robustness("Manpower", download_manpower_data, 55, 751, 365)

combined <- bind_rows(covid_results, manpower_results)
write.csv(combined, file.path(results_dir, "pca_robustness.csv"), row.names = FALSE)
cat("\nSaved pca_robustness.csv\n")

# --- Summary ---
cat("\nPCA Robustness Summary:\n")
print(combined[, c("scenario", "n_pcs", "var_explained", "mean_ratio", "close_rate")])

# --- Plots ---

p <- ggplot(combined, aes(x = factor(n_pcs), y = mean_ratio,
                           color = scenario, group = scenario)) +
    geom_line(linewidth = 1.2) +
    geom_point(size = 4) +
    geom_errorbar(aes(ymin = mean_ratio - 1.96 * sd_ratio / sqrt(n_valid),
                      ymax = mean_ratio + 1.96 * sd_ratio / sqrt(n_valid)),
                  width = 0.15, linewidth = 0.8) +
    geom_hline(yintercept = 1.0, linetype = "dashed", color = "grey50") +
    scale_color_manual(values = c("COVID-19" = "#3498db", "Manpower" = "#e67e22"),
                       name = "Scenario") +
    labs(title = "PCA x xi Robustness to Number of Principal Components",
         subtitle = sprintf("Monte Carlo with %d trials at k=%d (mean ratio +/- 95%% CI)", n_mc, k),
         x = "Number of Principal Components",
         y = "Mean Performance Ratio (PCA x xi / DIIM)") +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14),
          legend.position = "bottom")

ggsave(file.path(results_dir, "pca_robustness_plot.png"), p, width = 10, height = 7, bg = "white")
cat("Saved pca_robustness_plot.png\n")

p_var <- ggplot(combined, aes(x = factor(n_pcs), y = var_explained * 100,
                               color = scenario, group = scenario)) +
    geom_line(linewidth = 1.2) +
    geom_point(size = 4) +
    scale_color_manual(values = c("COVID-19" = "#3498db", "Manpower" = "#e67e22"),
                       name = "Scenario") +
    labs(title = "Variance Explained by Principal Components",
         x = "Number of Principal Components",
         y = "Cumulative Variance Explained (%)") +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14),
          legend.position = "bottom")

ggsave(file.path(results_dir, "pca_variance_explained.png"), p_var, width = 10, height = 7, bg = "white")
cat("Saved pca_variance_explained.png\n")

cat("Done!\n")
