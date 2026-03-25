# ============================================================================
# intervention_timing.R
# Tests how delayed intervention affects simplified method performance
# Simulates intervention at t=0, 7, 14, 30, 60 days into the disruption
# ============================================================================

if (!file.exists("functions.R")) setwd("..")

library(openxlsx)
library(igraph)
library(ggplot2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

delay_values <- c(0, 7, 14, 30, 60)
k <- 5
n_mc <- 500

method_prefixes <- c(
    "Total Output"  = "TotalOutput",
    "PCA x xi"      = "PCAxi",
    "PageRank x xi" = "PageRankxi",
    "BL x xi"       = "BLxi",
    "FL x xi"       = "FLxi"
)
method_labels <- names(method_prefixes)

# Modified DIIM that applies intervention at a specified delay (not at t=0)
# Runs the model in two phases: pre-intervention and post-intervention
DIIM_delayed <- function(q0, A_star, c_star, x, lockdown_duration, total_duration,
                          key_sectors, delay, intervention_magnitude = 0.10,
                          days_in_year = 366) {
    a_ii <- diag(A_star)
    num_sectors <- length(q0)

    # Phase 1: Run without intervention for 'delay' timesteps
    q0[q0 == 0] <- 1e-8
    qT <- q0 * 1/100
    k_rate <- log(q0 / qT) / (total_duration * (1 - a_ii))
    K <- diag(as.vector(k_rate))

    inoperability <- matrix(NA, nrow = num_sectors, ncol = total_duration)
    inoperability[, 1] <- q0

    for (t in 2:total_duration) {
        if (t == (delay + 1) && delay > 0) {
            # Apply intervention at the delay point
            inoperability[key_sectors, t - 1] <-
                inoperability[key_sectors, t - 1] * (1 - intervention_magnitude)
        }

        if (t <= lockdown_duration) {
            inoperability[, t] <- inoperability[, t - 1] +
                K %*% (A_star %*% inoperability[, t - 1] + c_star - inoperability[, t - 1])
        } else {
            inoperability[, t] <- inoperability[, t - 1] +
                K %*% (A_star %*% inoperability[, t - 1] - inoperability[, t - 1])
        }
    }

    # Handle delay=0 case (intervention at t=0, same as original DIIM)
    if (delay == 0) {
        # Re-run from scratch with intervention at t=0
        q0_int <- q0
        q0_int[key_sectors] <- q0_int[key_sectors] * (1 - intervention_magnitude)
        q0_int[q0_int == 0] <- 1e-8
        qT_int <- q0_int * 1/100
        k_rate_int <- log(q0_int / qT_int) / (total_duration * (1 - a_ii))
        K_int <- diag(as.vector(k_rate_int))

        inoperability[, 1] <- q0_int
        for (t in 2:total_duration) {
            if (t <= lockdown_duration) {
                inoperability[, t] <- inoperability[, t - 1] +
                    K_int %*% (A_star %*% inoperability[, t - 1] + c_star - inoperability[, t - 1])
            } else {
                inoperability[, t] <- inoperability[, t - 1] +
                    K_int %*% (A_star %*% inoperability[, t - 1] - inoperability[, t - 1])
            }
        }
    }

    x_daily <- x / days_in_year
    EL_evolution <- t(apply(inoperability, 1, cumsum)) * as.vector(x_daily)
    total_economic_loss <- sum(EL_evolution[, ncol(EL_evolution)])

    return(list(
        inoperability_evolution = inoperability,
        EL_evolution = EL_evolution,
        total_economic_loss = total_economic_loss
    ))
}

run_timing_analysis <- function(scenario_name, data_loader,
                                 lockdown_duration, total_duration,
                                 days_in_year) {
    cat(sprintf("\n=== %s Intervention Timing Analysis ===\n", scenario_name))
    data <- data_loader()
    A <- data$A; x <- data$x; c_star <- data$c_star; A_star <- data$A_star
    q0_base <- data$q0; q0_base[q0_base == 0] <- 1e-8
    num_sectors <- length(q0_base)

    simplified_rankings <- compute_simplified_rankings(A, A_star, x)

    set.seed(42)
    all_results <- list()

    for (delay in delay_values) {
        cat(sprintf("  delay=%d days: ", delay))

        # Store per-method results
        method_reductions <- list()
        diim_reductions <- numeric(0)
        for (m in names(method_prefixes)) method_reductions[[m]] <- numeric(0)
        valid_trials <- 0

        for (trial in 1:n_mc) {
            noise <- exp(rnorm(num_sectors, mean = 0, sd = 0.5))
            random_q0 <- pmin(pmax(q0_base * noise, 1e-6), 1)

            # Baseline (no intervention at all)
            base_model <- DIIM(random_q0, A_star, c_star, x,
                               lockdown_duration, total_duration,
                               days_in_year = days_in_year)
            base_loss <- base_model$total_economic_loss
            if (base_loss < 1e-6) next

            # DIIM gold standard (use baseline to identify top-k)
            max_el <- apply(base_model$EL_evolution, 1, max)
            diim_topk <- order(max_el, decreasing = TRUE)[1:k]

            diim_delayed <- DIIM_delayed(random_q0, A_star, c_star, x,
                                          lockdown_duration, total_duration,
                                          key_sectors = diim_topk,
                                          delay = delay,
                                          days_in_year = days_in_year)
            diim_reduction <- base_loss - diim_delayed$total_economic_loss
            if (diim_reduction < 1e-10) next
            valid_trials <- valid_trials + 1
            diim_reductions <- c(diim_reductions, diim_reduction)

            # Each simplified method with delayed intervention
            for (method_name in names(simplified_rankings)) {
                topk <- simplified_rankings[[method_name]][1:k]
                m_delayed <- DIIM_delayed(random_q0, A_star, c_star, x,
                                           lockdown_duration, total_duration,
                                           key_sectors = topk,
                                           delay = delay,
                                           days_in_year = days_in_year)
                m_reduction <- base_loss - m_delayed$total_economic_loss
                method_reductions[[method_name]] <- c(
                    method_reductions[[method_name]],
                    m_reduction / diim_reduction
                )
            }
        }

        cat(sprintf("%d valid trials\n", valid_trials))

        # Record aggregated results
        for (method_name in names(method_reductions)) {
            r <- method_reductions[[method_name]]
            if (length(r) == 0) next
            all_results[[length(all_results) + 1]] <- data.frame(
                scenario = scenario_name,
                delay = delay,
                method = method_name,
                mean_ratio = mean(r),
                sd_ratio = sd(r),
                close_rate = mean(r >= 0.80),
                mean_diim_reduction = mean(diim_reductions),
                n_valid = length(r),
                stringsAsFactors = FALSE
            )
        }
    }

    return(bind_rows(all_results))
}

# --- Run for both scenarios ---
covid_results <- run_timing_analysis("COVID-19", download_data, 55, 751, 366)
manpower_results <- run_timing_analysis("Manpower", download_manpower_data, 55, 751, 365)

combined <- bind_rows(covid_results, manpower_results)
write.csv(combined, file.path(results_dir, "intervention_timing.csv"), row.names = FALSE)
cat("\nSaved intervention_timing.csv\n")

# --- Print summary ---
cat("\nIntervention Timing Summary:\n")
for (sc in unique(combined$scenario)) {
    cat(sprintf("\n--- %s ---\n", sc))
    sc_data <- combined[combined$scenario == sc, ]
    for (d in delay_values) {
        d_data <- sc_data[sc_data$delay == d, ]
        best <- d_data[which.max(d_data$mean_ratio), ]
        cat(sprintf("  delay=%2d days: best=%s (ratio=%.3f, close=%.1f%%)\n",
            d, best$method, best$mean_ratio, best$close_rate * 100))
    }
}

# --- Plot 1: Mean ratio vs delay for each method ---
cat("\nGenerating plots...\n")

p1 <- ggplot(combined, aes(x = delay, y = mean_ratio,
                             color = method, shape = method)) +
    geom_line(linewidth = 1.1) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = mean_ratio - 1.96 * sd_ratio / sqrt(n_valid),
                      ymax = mean_ratio + 1.96 * sd_ratio / sqrt(n_valid)),
                  width = 2, alpha = 0.5) +
    facet_wrap(~scenario) +
    geom_hline(yintercept = 1.0, linetype = "dashed", color = "grey50") +
    scale_color_brewer(palette = "Set2", name = "Method") +
    labs(title = "Dynamic Intervention Timing: Method Performance",
         subtitle = sprintf("Mean performance ratio +/- 95%% CI (%d MC trials, k=%d)", n_mc, k),
         x = "Intervention Delay (days)",
         y = "Mean Performance Ratio (simplified / DIIM)") +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14),
          legend.position = "bottom") +
    guides(color = guide_legend(nrow = 2))

ggsave(file.path(results_dir, "intervention_timing_plot.png"), p1, width = 12, height = 7)
cat("Saved intervention_timing_plot.png\n")

# --- Plot 2: DIIM absolute reduction vs delay (shows how much benefit is lost with delay) ---
diim_decay <- combined %>%
    group_by(scenario, delay) %>%
    summarise(mean_diim_red = mean(mean_diim_reduction), .groups = "drop")

p2 <- ggplot(diim_decay, aes(x = delay, y = mean_diim_red,
                               color = scenario, group = scenario)) +
    geom_line(linewidth = 1.2) +
    geom_point(size = 4) +
    scale_color_manual(values = c("COVID-19" = "#3498db", "Manpower" = "#e67e22"),
                       name = "Scenario") +
    labs(title = "Cost of Delayed Intervention",
         subtitle = "Mean DIIM loss reduction decreases with intervention delay",
         x = "Intervention Delay (days)",
         y = "Mean DIIM Loss Reduction") +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14),
          legend.position = "bottom")

ggsave(file.path(results_dir, "intervention_timing_decay.png"), p2, width = 10, height = 7)
cat("Saved intervention_timing_decay.png\n")

cat("Done!\n")
