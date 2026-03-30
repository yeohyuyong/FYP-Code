if (!file.exists("functions.R")) setwd("..")

library(openxlsx)
library(igraph)
library(ggplot2)
library(reshape2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

simplified_methods <- c("Total Output", "PCA x xi", "PageRank x xi", "BL x xi", "FL x xi")

diim_rank_sectors <- function(q0, A_star, c_star, x, lockdown_duration, total_duration, days_in_year = 366) {
    model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, days_in_year = days_in_year)
    max_el <- apply(model$EL_evolution, 1, max)
    ranked_sectors <- order(max_el, decreasing = TRUE)
    return(list(ranked_sectors = ranked_sectors, max_el = max_el))
}

evaluate_topk <- function(ranked_sectors, k, q0, A_star, c_star, x, lockdown_duration, total_duration, base_loss, days_in_year = 366) {
    key_sectors <- ranked_sectors[1:k]
    model_int <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = key_sectors, days_in_year = days_in_year)
    reduction <- base_loss - model_int$total_economic_loss
    pct_reduction <- reduction / base_loss * 100
    return(list(
        k = k, sectors = key_sectors,
        loss = model_int$total_economic_loss,
        reduction = reduction, pct_reduction = pct_reduction
    ))
}

run_scenario <- function(scenario_name, data_loader, lockdown_duration = 55, total_duration = 751, days_in_year = 366, target_k_values = c(3, 5, 7, 10, 15, 20)) {
    cat(sprintf("\n%s Scenario\n", scenario_name))

    data <- data_loader()
    A <- data$A
    x <- data$x
    q0 <- data$q0
    q0[q0 == 0] <- 1e-8 # avoid divide-by-zero
    c_star <- data$c_star
    A_star <- data$A_star
    num_sectors <- length(q0)

    target_k_values <- target_k_values[target_k_values < num_sectors]
    cat(sprintf("  Sectors: %d, k values: %s\n", num_sectors, paste(target_k_values, collapse = ", ")))

    # Baseline (no intervention)
    base_model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, days_in_year = days_in_year)
    base_loss <- base_model$total_economic_loss
    cat(sprintf("  Baseline loss: %.2f\n\n", base_loss))

    # --- Sector rankings ---

    cat("  Computing simplified rankings (xi-weighted)...\n")
    simplified_rankings <- compute_simplified_rankings(A, A_star, x)
    for (method_name in names(simplified_rankings)) { cat(sprintf("    %-20s top-5: [%s]\n", method_name, paste(simplified_rankings[[method_name]][1:min(5, num_sectors)], collapse = ", "))) }

    # DIIM ranking (gold standard — uses full temporal simulation)
    cat("  Computing DIIM ranking...\n")
    diim_results <- diim_rank_sectors(q0, A_star, c_star, x, lockdown_duration, total_duration, days_in_year)
    cat(sprintf("    DIIM top-5: [%s]\n", paste(diim_results$ranked_sectors[1:5], collapse = ", ")))

    cat("\n  Evaluating all methods at each k...\n")

    all_methods <- simplified_rankings
    all_methods[["DIIM"]] <- diim_results$ranked_sectors

    results_df <- data.frame(scenario = character(), method = character(), k = integer(), pct_reduction = numeric(), sectors = character(), stringsAsFactors = FALSE)

    for (method_name in names(all_methods)) {
        for (kval in target_k_values) {
            evaluation <- evaluate_topk(all_methods[[method_name]], kval, q0, A_star, c_star, x, lockdown_duration, total_duration, base_loss, days_in_year)

            results_df <- rbind(results_df, data.frame(scenario = scenario_name, method = method_name, k = kval, pct_reduction = evaluation$pct_reduction, sectors = paste(evaluation$sectors, collapse = ","), stringsAsFactors = FALSE))
        }
    }

    cat("\n  Results (% loss reduction):\n")
    summary_table <- reshape2::dcast(results_df, method ~ k, value.var = "pct_reduction")
    print(summary_table)

    return(results_df)
}

# COVID-19 (15 sectors)
covid_results <- run_scenario("COVID-19", download_data, lockdown_duration = 55, total_duration = 751, days_in_year = 366, target_k_values = c(3, 5, 7, 10, 12))

# Manpower (15 sectors)
manpower_results <- run_scenario("Manpower", download_manpower_data, lockdown_duration = 55, total_duration = 751, days_in_year = 365, target_k_values = c(3, 5, 7, 10, 12))

combined <- rbind(covid_results, manpower_results)
write.csv(combined, file.path(results_dir, "enhanced_comparison.csv"), row.names = FALSE)
cat("\nSaved enhanced_comparison.csv\n")

plot_topk_comparison <- function(data, scenario_name, filename) {
    plot_data <- data[data$scenario == scenario_name, ]
    plot_data$method_type <- ifelse(plot_data$method %in% simplified_methods, "Simplified Method", ifelse(plot_data$method == "DIIM", "DIIM", "Other"))

    p <- ggplot(plot_data, aes(x = k, y = pct_reduction, color = method, linetype = method_type)) + geom_line(linewidth = 1.1) + geom_point(size = 2.5) + scale_color_brewer(palette = "Set2", name = "Method") + scale_linetype_manual(values = c("Simplified Method" = "solid", "DIIM" = "dashed", "Other" = "dotted"), name = "Type") + labs(title = sprintf("%s: Loss Reduction vs k", scenario_name), subtitle = "Comparing xi-weighted simplified methods, DIIM, and Sensitivity", x = "Number of Sectors Intervened (k)", y = "% Loss Reduction") + theme_minimal() + theme(plot.title = element_text(face = "bold", size = 14), legend.position = "bottom", legend.box = "vertical") + guides(color = guide_legend(nrow = 3), linetype = guide_legend(nrow = 1))

    ggsave(file.path(results_dir, filename), p, width = 12, height = 7, bg = "white"); cat(sprintf("Saved %s\n", filename))
}

plot_topk_comparison(combined, "COVID-19", "enhanced_covid_topk.png")
plot_topk_comparison(combined, "Manpower", "enhanced_manpower_topk.png")

simplified_data <- combined[combined$method %in% simplified_methods, ]
if (nrow(simplified_data) > 0) {
    p_simplified <- ggplot(simplified_data, aes(x = k, y = pct_reduction, color = method, shape = method)) + geom_line(linewidth = 1.2) + geom_point(size = 3) + facet_wrap(~scenario, scales = "free_y") + scale_color_brewer(palette = "Dark2", name = "Method") + labs(title = "Xi-Weighted Simplified Methods Comparison", x = "Top-k Sectors", y = "% Loss Reduction") + theme_minimal() + theme(plot.title = element_text(face = "bold", size = 14), legend.position = "bottom")
    ggsave(file.path(results_dir, "enhanced_simplified_methods.png"), p_simplified, width = 12, height = 6, bg = "white"); cat("Saved enhanced_simplified_methods.png\n")
}

winners <- combined %>%
    group_by(scenario, k) %>%
    filter(pct_reduction == max(pct_reduction)) %>%
    ungroup()

p_winner <- ggplot(winners, aes(x = factor(k), y = pct_reduction, fill = method)) + geom_bar(stat = "identity", position = "dodge", alpha = 0.85) + facet_wrap(~scenario, scales = "free_y") + scale_fill_brewer(palette = "Set2", name = "Best Method") + labs(title = "Best Method at Each k", x = "Number of Sectors (k)", y = "% Loss Reduction") + theme_minimal() + theme(plot.title = element_text(face = "bold", size = 14), legend.position = "bottom")

ggsave(file.path(results_dir, "enhanced_best_at_k.png"), p_winner, width = 12, height = 6, bg = "white")
cat("Saved enhanced_best_at_k.png\n")

cat("ENHANCED COMPARISON SUMMARY\n\n")

for (scenario_name in unique(combined$scenario)) {
    cat(sprintf("--- %s ---\n", scenario_name))
    scenario_data <- combined[combined$scenario == scenario_name, ]

    for (kval in sort(unique(scenario_data$k))) {
        k_data <- scenario_data[scenario_data$k == kval, ]
        best_method <- k_data[which.max(k_data$pct_reduction), ]
        cat(sprintf("  k=%2d: Best = %-25s (%.2f%%)", kval, best_method$method, best_method$pct_reduction))

        simp_vals <- k_data[k_data$method %in% simplified_methods, ]
        if (nrow(simp_vals) > 0) {
            best_simp <- simp_vals[which.max(simp_vals$pct_reduction), ]
            cat(sprintf("  |  Best Simplified = %-20s (%.2f%%)", best_simp$method, best_simp$pct_reduction))
        }
        cat("\n")
    }
    cat("\n")
}

cat("\nDone!\n")
