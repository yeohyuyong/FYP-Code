# ============================================================================
# scatter_comparison.R
# Generates scatter plots: simplified method loss reduction vs DIIM loss reduction
# Uses existing decision_rules_mc_data.csv if available, otherwise runs 500 MC trials
# ============================================================================

if (!file.exists("functions.R")) setwd("..")

library(ggplot2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
mc_file <- file.path(results_dir, "decision_rules_mc_data.csv")

method_prefixes <- c(
    "Total Output"  = "TotalOutput",
    "PCA x xi"      = "PCAxi",
    "PageRank x xi" = "PageRankxi",
    "BL x xi"       = "BLxi",
    "FL x xi"       = "FLxi"
)

method_labels <- c("Total Output", "PCA x xi", "PageRank x xi", "BL x xi", "FL x xi")

# --- Load or generate Monte Carlo data ---
if (file.exists(mc_file)) {
    cat("Loading existing MC data from", mc_file, "\n")
    mc_data <- read.csv(mc_file, stringsAsFactors = FALSE)
} else {
    cat("No existing MC data found. Running 500-trial MC at k=5...\n")

    run_quick_mc <- function(scenario_name, data_loader, lockdown_duration, total_duration,
                             days_in_year, n_mc = 500, k = 5) {
        data <- data_loader()
        A <- data$A; x <- data$x; c_star <- data$c_star; A_star <- data$A_star
        q0_base <- data$q0; q0_base[q0_base == 0] <- 1e-8
        num_sectors <- length(q0_base)
        simplified_rankings <- compute_simplified_rankings(A, A_star, x)

        set.seed(42)
        results <- list()

        for (trial in 1:n_mc) {
            noise <- exp(rnorm(num_sectors, mean = 0, sd = 0.5))
            random_q0 <- pmin(pmax(q0_base * noise, 1e-6), 1)

            base_model <- DIIM(random_q0, A_star, c_star, x, lockdown_duration, total_duration,
                               days_in_year = days_in_year)
            base_loss <- base_model$total_economic_loss
            if (base_loss < 1e-6) next

            max_el <- apply(base_model$EL_evolution, 1, max)
            diim_topk <- order(max_el, decreasing = TRUE)[1:k]
            diim_model <- DIIM(random_q0, A_star, c_star, x, lockdown_duration, total_duration,
                               key_sectors = diim_topk, days_in_year = days_in_year)
            diim_reduction <- base_loss - diim_model$total_economic_loss
            if (diim_reduction < 1e-10) next

            row <- data.frame(trial = trial, scenario = scenario_name, k = k,
                              base_loss = base_loss, diim_reduction = diim_reduction,
                              stringsAsFactors = FALSE)

            for (method_name in names(simplified_rankings)) {
                topk <- simplified_rankings[[method_name]][1:k]
                m_model <- DIIM(random_q0, A_star, c_star, x, lockdown_duration, total_duration,
                                key_sectors = topk, days_in_year = days_in_year)
                m_reduction <- base_loss - m_model$total_economic_loss
                prefix <- method_prefixes[method_name]
                row[[paste0(prefix, "_reduction")]] <- m_reduction
                row[[paste0(prefix, "_ratio")]] <- m_reduction / diim_reduction
            }
            results[[length(results) + 1]] <- row
            if (trial %% 100 == 0) cat(sprintf("  %s: %d/%d\n", scenario_name, trial, n_mc))
        }
        return(bind_rows(results))
    }

    covid_mc <- run_quick_mc("COVID-19", download_data, 55, 751, 366)
    manpower_mc <- run_quick_mc("Manpower", download_manpower_data, 55, 751, 365)
    mc_data <- bind_rows(covid_mc, manpower_mc)
}

# --- Generate scatter plots ---
cat("Generating scatter plots...\n")

for (sc in unique(mc_data$scenario)) {
    sc_data <- mc_data[mc_data$scenario == sc, ]

    # Use k=5 data (or the first k available)
    if ("k" %in% names(sc_data)) {
        available_k <- sort(unique(sc_data$k))
        target_k <- ifelse(5 %in% available_k, 5, available_k[1])
        sc_data <- sc_data[sc_data$k == target_k, ]
    }

    # Build long-form data for faceted plot
    plot_rows <- list()
    for (i in seq_along(method_labels)) {
        prefix <- method_prefixes[method_labels[i]]
        red_col <- paste0(prefix, "_reduction")
        if (!red_col %in% names(sc_data)) next

        plot_rows[[i]] <- data.frame(
            method = method_labels[i],
            diim_reduction = sc_data$diim_reduction,
            method_reduction = sc_data[[red_col]],
            stringsAsFactors = FALSE
        )
    }
    plot_df <- bind_rows(plot_rows)

    if (nrow(plot_df) == 0) next

    # Determine axis limits for the 45-degree line
    max_val <- max(c(plot_df$diim_reduction, plot_df$method_reduction), na.rm = TRUE)

    p <- ggplot(plot_df, aes(x = diim_reduction, y = method_reduction)) +
        geom_point(alpha = 0.25, size = 1.2, color = "#2c3e50") +
        geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "#e74c3c", linewidth = 0.8) +
        facet_wrap(~method, nrow = 2) +
        coord_cartesian(xlim = c(0, max_val), ylim = c(0, max_val)) +
        labs(
            title = sprintf("%s: Simplified Method vs DIIM Loss Reduction", sc),
            subtitle = "Points above the 45° line = simplified method outperforms DIIM",
            x = "DIIM Loss Reduction",
            y = "Simplified Method Loss Reduction"
        ) +
        theme_minimal() +
        theme(
            plot.title = element_text(face = "bold", size = 14),
            strip.text = element_text(face = "bold", size = 11),
            panel.grid.minor = element_blank()
        )

    sc_label <- tolower(gsub("[- ]", "_", sc))
    fname <- sprintf("scatter_%s_k5.png", sc_label)
    ggsave(file.path(results_dir, fname), p, width = 12, height = 8)
    cat(sprintf("  Saved %s\n", fname))
}

cat("Done!\n")
