# ============================================================================
# intervention_sensitivity.R
# Tests whether the three-regimes finding is robust to intervention magnitude
# Runs Monte Carlo (500 trials) for magnitudes {5%, 10%, 20%, 30%} across k values
# ============================================================================

if (!file.exists("functions.R")) setwd("..")

library(openxlsx)
library(igraph)
library(ggplot2)
library(dplyr)
library(reshape2)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

intervention_magnitudes <- c(0.05, 0.10, 0.20, 0.30)
k_values <- c(3, 5, 7, 10, 12)
n_mc <- 500

method_prefixes <- c(
    "Total Output"  = "TotalOutput",
    "PCA x xi"      = "PCAxi",
    "PageRank x xi" = "PageRankxi",
    "BL x xi"       = "BLxi",
    "FL x xi"       = "FLxi"
)

run_sensitivity <- function(scenario_name, data_loader,
                             lockdown_duration, total_duration,
                             days_in_year) {
    cat(sprintf("\n=== %s Sensitivity Analysis ===\n", scenario_name))
    data <- data_loader()
    A <- data$A; x <- data$x; c_star <- data$c_star; A_star <- data$A_star
    q0_base <- data$q0; q0_base[q0_base == 0] <- 1e-8
    num_sectors <- length(q0_base)
    simplified_rankings <- compute_simplified_rankings(A, A_star, x)

    set.seed(42)
    all_results <- list()

    for (mag in intervention_magnitudes) {
        for (k in k_values) {
            cat(sprintf("  mag=%.0f%%, k=%d: ", mag * 100, k))
            ratios <- list()
            for (m in names(method_prefixes)) ratios[[m]] <- numeric(0)
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
                                   days_in_year = days_in_year,
                                   intervention_magnitude = mag)
                diim_reduction <- base_loss - diim_model$total_economic_loss
                if (diim_reduction < 1e-10) next
                valid_trials <- valid_trials + 1

                for (method_name in names(simplified_rankings)) {
                    topk <- simplified_rankings[[method_name]][1:k]
                    m_model <- DIIM(random_q0, A_star, c_star, x,
                                    lockdown_duration, total_duration,
                                    key_sectors = topk,
                                    days_in_year = days_in_year,
                                    intervention_magnitude = mag)
                    m_reduction <- base_loss - m_model$total_economic_loss
                    ratios[[method_name]] <- c(ratios[[method_name]],
                                               m_reduction / diim_reduction)
                }
            }

            cat(sprintf("%d valid trials\n", valid_trials))

            for (method_name in names(ratios)) {
                r <- ratios[[method_name]]
                if (length(r) == 0) next
                all_results[[length(all_results) + 1]] <- data.frame(
                    scenario = scenario_name, magnitude = mag, k = k,
                    method = method_name,
                    mean_ratio = mean(r), sd_ratio = sd(r),
                    close_rate = mean(r >= 0.80), n_valid = length(r),
                    stringsAsFactors = FALSE
                )
            }
        }
    }
    return(bind_rows(all_results))
}

# --- Run for both scenarios ---
covid_results <- run_sensitivity("COVID-19", download_data, 55, 751, 366)
manpower_results <- run_sensitivity("Manpower", download_manpower_data, 55, 751, 365)

combined <- bind_rows(covid_results, manpower_results)
write.csv(combined, file.path(results_dir, "sensitivity_intervention.csv"), row.names = FALSE)
cat("\nSaved sensitivity_intervention.csv\n")

# --- Heatmap: Best simplified method ratio by magnitude and k ---
cat("Generating heatmap...\n")

best_method_data <- combined %>%
    group_by(scenario, magnitude, k) %>%
    summarise(best_ratio = max(mean_ratio), .groups = "drop")

p_heat <- ggplot(best_method_data, aes(x = factor(k),
                                        y = sprintf("%.0f%%", magnitude * 100),
                                        fill = best_ratio)) +
    geom_tile(color = "white", linewidth = 0.5) +
    geom_text(aes(label = sprintf("%.3f", best_ratio)), size = 3.5) +
    facet_wrap(~scenario) +
    scale_fill_gradient2(low = "#e74c3c", mid = "#f39c12", high = "#27ae60",
                         midpoint = 0.95, limits = c(0.70, 1.05),
                         name = "Best Mean Ratio") +
    labs(title = "Intervention Magnitude Sensitivity: Best Simplified Method",
         subtitle = "Mean performance ratio of best simplified method vs DIIM",
         x = "Number of Sectors Intervened (k)",
         y = "Intervention Magnitude") +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14),
          panel.grid = element_blank())

ggsave(file.path(results_dir, "sensitivity_intervention_heatmap.png"), p_heat,
       width = 12, height = 6)
cat("Saved sensitivity_intervention_heatmap.png\n")

# --- Line plot: all methods across magnitudes at k=5 ---
k5_data <- combined[combined$k == 5, ]
if (nrow(k5_data) > 0) {
    p_line <- ggplot(k5_data, aes(x = magnitude * 100, y = mean_ratio,
                                   color = method, shape = method)) +
        geom_line(linewidth = 1.1) +
        geom_point(size = 3) +
        geom_errorbar(aes(ymin = mean_ratio - 1.96 * sd_ratio / sqrt(n_valid),
                          ymax = mean_ratio + 1.96 * sd_ratio / sqrt(n_valid)),
                      width = 1, alpha = 0.6) +
        facet_wrap(~scenario) +
        geom_hline(yintercept = 1.0, linetype = "dashed", color = "grey40") +
        scale_color_brewer(palette = "Set2", name = "Method") +
        labs(title = "Intervention Magnitude Sensitivity at k=5",
             subtitle = "Mean performance ratio +/- 95% CI",
             x = "Intervention Magnitude (%)",
             y = "Mean Performance Ratio (simplified / DIIM)") +
        theme_minimal() +
        theme(plot.title = element_text(face = "bold", size = 14),
              legend.position = "bottom")

    ggsave(file.path(results_dir, "sensitivity_intervention_k5.png"), p_line,
           width = 12, height = 6)
    cat("Saved sensitivity_intervention_k5.png\n")
}

cat("Done!\n")
