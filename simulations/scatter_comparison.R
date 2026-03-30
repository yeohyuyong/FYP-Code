# Scatter plots of simplified vs DIIM loss reduction.
# Reuses decision_rules_mc_data.csv if available, otherwise runs 500 MC trials.

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

# --- Load Monte Carlo data ---
cat("Loading MC data from", mc_file, "\n")
mc_data <- read.csv(mc_file, stringsAsFactors = FALSE)

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

    plot_rows <- list()
    for (i in seq_along(method_prefixes)) {
        red_col <- paste0(method_prefixes[i], "_reduction")
        if (!red_col %in% names(sc_data)) next

        plot_rows[[i]] <- data.frame(
            method = names(method_prefixes)[i],
            diim_reduction = sc_data$diim_reduction,
            method_reduction = sc_data[[red_col]],
            stringsAsFactors = FALSE
        )
    }
    plot_df <- bind_rows(plot_rows)

    if (nrow(plot_df) == 0) next

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
    ggsave(file.path(results_dir, fname), p, width = 12, height = 8, bg = "white")
    cat(sprintf("  Saved %s\n", fname))
}

cat("Done!\n")
