setwd("..")

library(openxlsx)
library(igraph)
library(ggplot2)
library(reshape2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

diim_rank_sectors <- function(q0, A_star, c_star, x,
                              lockdown_duration, total_duration,
                              days_in_year = 366) {
    model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
                  days_in_year = days_in_year)
    max_el <- apply(model$EL_evolution, 1, max)
    ranked_sectors <- order(max_el, decreasing = TRUE)
    return(list(ranked_sectors = ranked_sectors, max_el = max_el))
}

sensitivity_rank_sectors <- function(q0, A_star, c_star, x,
                                     lockdown_duration, total_duration,
                                     days_in_year = 366) {
    n <- length(q0)
    base_model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
                       days_in_year = days_in_year)
    base_loss <- base_model$total_economic_loss

    marginal_reductions <- numeric(n)
    for (i in 1:n) {
        q0_test    <- q0
        q0_test[i] <- q0_test[i] * 0.9
        model_test <- DIIM(q0_test, A_star, c_star, x, lockdown_duration,
                           total_duration, days_in_year = days_in_year)
        marginal_reductions[i] <- base_loss - model_test$total_economic_loss
    }

    ranked_sectors <- order(marginal_reductions, decreasing = TRUE)
    return(list(ranked_sectors = ranked_sectors,
               marginal_reductions = marginal_reductions))
}

evaluate_topk <- function(ranked_sectors, k, q0, A_star, c_star, x,
                          lockdown_duration, total_duration,
                          base_loss, days_in_year = 366) {
    key_sectors <- ranked_sectors[1:k]
    model_int <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
                      key_sectors = key_sectors, days_in_year = days_in_year)
    reduction <- base_loss - model_int$total_economic_loss
    pct_reduction <- reduction / base_loss * 100
    return(list(k = k, sectors = key_sectors,
               loss = model_int$total_economic_loss,
               reduction = reduction, pct_reduction = pct_reduction))
}

run_scenario <- function(scenario_name, data_loader,
                         lockdown_duration = 55, total_duration = 751,
                         days_in_year = 366,
                         target_k_values = c(3, 5, 7, 10, 15, 20)) {
    cat(sprintf("\n%s Scenario\n", scenario_name))

    # 1. Setup the scenario data
    data   <- data_loader()
    A      <- data$A
    x      <- data$x
    
    # Add a tiny epsilon to 0 bounds to prevent divide-by-zero or static evaluation errors
    q0     <- data$q0
    q0[q0 == 0] <- 1e-8
    
    c_star <- data$c_star
    A_star <- data$A_star
    num_sectors <- length(q0)

    # Filter target k values so we don't try to intervene on more sectors than exist
    target_k_values <- target_k_values[target_k_values < num_sectors]
    cat(sprintf("  Sectors: %d, k values: %s\n", num_sectors, paste(target_k_values, collapse = ", ")))

    # 2. Compute Baseline Model (No intervention)
    base_model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
                       days_in_year = days_in_year)
    base_loss <- base_model$total_economic_loss
    cat(sprintf("  Baseline loss: %.2f\n\n", base_loss))

    # --- Step 1: Compute sector rankings using all methods ---

    # Cheap methods (xi-weighted, no simulation needed)
    cat("  Computing cheap rankings (xi-weighted)...\n")
    cheap_rankings <- compute_cheap_rankings(A, A_star, x)
    for (method_name in names(cheap_rankings)) {
        cat(sprintf("    %-20s top-5: [%s]\n", method_name,
            paste(cheap_rankings[[method_name]][1:min(5, num_sectors)], collapse = ", ")))
    }

    # DIIM - gold standard (runs the simulation to observe temporal progression)
    cat("  Computing DIIM ranking...\n")
    diim_results <- diim_rank_sectors(q0, A_star, c_star, x,
                                      lockdown_duration, total_duration, days_in_year)
    cat(sprintf("    DIIM top-5: [%s]\n",
        paste(diim_results$ranked_sectors[1:5], collapse = ", ")))

    # Sensitivity - measures the marginal drop in total loss for each sector's isolated intervention
    cat("  Computing Sensitivity ranking...\n")
    sensitivity_results <- sensitivity_rank_sectors(q0, A_star, c_star, x,
                                                    lockdown_duration, total_duration, days_in_year)
    cat(sprintf("    Sensitivity top-5: [%s]\n",
        paste(sensitivity_results$ranked_sectors[1:5], collapse = ", ")))


    # --- Step 2: Evaluate the performance of all methods across each k ---
    cat("\n  Evaluating all methods at each k...\n")

    # Combine all sector rankings into a single named list
    all_methods <- cheap_rankings  # already contains all 5 cheap methods
    all_methods[["DIIM"]]          <- diim_results$ranked_sectors
    all_methods[["Sensitivity"]]    <- sensitivity_results$ranked_sectors

    # Prepare an empty dataframe to hold results
    results_dataframe <- data.frame(
        scenario = character(), method = character(),
        k = integer(), pct_reduction = numeric(),
        sectors = character(), stringsAsFactors = FALSE)

    # Loop through each ranking method and test its top k sectors
    for (method_name in names(all_methods)) {
        for (current_k in target_k_values) {
            
            # Evaluate the percentage loss reduction of intervening on these specific 'k' sectors
            evaluation <- evaluate_topk(all_methods[[method_name]], current_k, q0, A_star, c_star, x,
                                        lockdown_duration, total_duration, base_loss, days_in_year)
                                        
            results_dataframe <- rbind(results_dataframe, data.frame(
                scenario = scenario_name, method = method_name, k = current_k,
                pct_reduction = evaluation$pct_reduction,
                sectors = paste(evaluation$sectors, collapse = ","),
                stringsAsFactors = FALSE))
        }
    }

    # Print summary format via data shaping
    cat("\n  Results (% loss reduction):\n")
    summary_table <- reshape2::dcast(results_dataframe, method ~ k, value.var = "pct_reduction")
    print(summary_table)

    return(results_dataframe)
}

# COVID-19 (15 sectors)
covid_results <- run_scenario("COVID-19", download_data,
    lockdown_duration = 55, total_duration = 751,
    days_in_year = 366, target_k_values = c(3, 5, 7, 10, 12))

# Manpower (15 sectors)
manpower_results <- run_scenario("Manpower", download_manpower_data,
    lockdown_duration = 55, total_duration = 751,
    days_in_year = 365, target_k_values = c(3, 5, 7, 10, 12))

combined <- rbind(covid_results, manpower_results)
write.csv(combined, file.path(results_dir, "enhanced_comparison.csv"), row.names = FALSE)
cat("\nSaved enhanced_comparison.csv\n")

plot_topk_comparison <- function(data, scenario_name, filename) {
    plot_data <- data[data$scenario == scenario_name, ]
    plot_data$method_type <- ifelse(plot_data$method %in% c("Total Output", "PCA x xi", "PageRank x xi", "BL x xi", "FL x xi"), "Cheap Method",
        ifelse(plot_data$method == "DIIM", "DIIM",
            ifelse(plot_data$method == "Sensitivity", "Sensitivity", "Other")))

    p <- ggplot(plot_data, aes(x = k, y = pct_reduction,
                               color = method, linetype = method_type)) +
        geom_line(linewidth = 1.1) +
        geom_point(size = 2.5) +
        scale_color_brewer(palette = "Set2", name = "Method") +
        scale_linetype_manual(
            values = c("Cheap Method" = "solid", "DIIM" = "dashed",
                       "Sensitivity" = "dotdash", "Other" = "dotted"),
            name = "Type") +
        labs(title = sprintf("%s: Loss Reduction vs k", scenario_name),
             subtitle = "Comparing xi-weighted cheap methods, DIIM, and Sensitivity",
             x = "Number of Sectors Intervened (k)",
             y = "% Loss Reduction") +
        theme_minimal() +
        theme(plot.title = element_text(face = "bold", size = 14),
              legend.position = "bottom", legend.box = "vertical") +
        guides(color = guide_legend(nrow = 3), linetype = guide_legend(nrow = 1))

    ggsave(file.path(results_dir, filename), p, width = 12, height = 7)
    cat(sprintf("Saved %s\n", filename))
}

plot_topk_comparison(combined, "COVID-19", "enhanced_covid_topk.png")
plot_topk_comparison(combined, "Manpower", "enhanced_manpower_topk.png")

# Compare all cheap methods side-by-side
cheap_methods <- c("Total Output", "PCA x xi", "PageRank x xi", "BL x xi", "FL x xi")
cheap_data <- combined[combined$method %in% cheap_methods, ]
if (nrow(cheap_data) > 0) {
    p_cheap <- ggplot(cheap_data, aes(x = k, y = pct_reduction,
                                   color = method, shape = method)) +
        geom_line(linewidth = 1.2) +
        geom_point(size = 3) +
        facet_wrap(~scenario, scales = "free_y") +
        scale_color_brewer(palette = "Dark2", name = "Method") +
        labs(title = "Xi-Weighted Cheap Methods Comparison",
             x = "Top-k Sectors", y = "% Loss Reduction") +
        theme_minimal() +
        theme(plot.title = element_text(face = "bold", size = 14),
              legend.position = "bottom")
    ggsave(file.path(results_dir, "enhanced_cheap_methods.png"), p_cheap,
           width = 12, height = 6)
    cat("Saved enhanced_cheap_methods.png\n")
}

# Find the method that achieved the maximum percentage reduction for each 'k' value in each scenario
winners <- combined %>%
    group_by(scenario, k) %>%
    filter(pct_reduction == max(pct_reduction)) %>%
    ungroup()

p_winner <- ggplot(winners, aes(x = factor(k), y = pct_reduction, fill = method)) +
    geom_bar(stat = "identity", position = "dodge", alpha = 0.85) +
    facet_wrap(~scenario, scales = "free_y") +
    scale_fill_brewer(palette = "Set2", name = "Best Method") +
    labs(title = "Best Method at Each k",
         x = "Number of Sectors (k)", y = "% Loss Reduction") +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14),
          legend.position = "bottom")

ggsave(file.path(results_dir, "enhanced_best_at_k.png"), p_winner,
       width = 12, height = 6)
cat("Saved enhanced_best_at_k.png\n")

cat("ENHANCED COMPARISON SUMMARY\n\n")

for (scenario_name in unique(combined$scenario)) {
    cat(sprintf("--- %s ---\n", scenario_name))
    scenario_data <- combined[combined$scenario == scenario_name, ]

    for (current_k in sort(unique(scenario_data$k))) {
        k_data <- scenario_data[scenario_data$k == current_k, ]
        best_method <- k_data[which.max(k_data$pct_reduction), ]
        cat(sprintf("  k=%2d: Best = %-25s (%.2f%%)", current_k, best_method$method, best_method$pct_reduction))

        # Also show the best PCA variant alongside the overall winner for easy comparison
        pca_vals <- k_data[k_data$method %in% c("Total Output", "PCA x xi", "PageRank x xi", "BL x xi", "FL x xi"), ]
        if (nrow(pca_vals) > 0) {
            best_pca <- pca_vals[which.max(pca_vals$pct_reduction), ]
            cat(sprintf("  |  Best PCA = %-20s (%.2f%%)", best_pca$method, best_pca$pct_reduction))
        }
        cat("\n")
    }
    cat("\n")
}

# Compare Total Output baseline vs best cheap structural method (at k=5)
cat("\nComparison: Structural Methods vs Total Output (at k=5):\n")
for (scenario_name in unique(combined$scenario)) {
    scenario_data <- combined[combined$scenario == scenario_name & combined$k == 5, ]
    
    total_output_red <- scenario_data[scenario_data$method == "Total Output", "pct_reduction"]
    structural <- scenario_data[scenario_data$method %in% c("PCA", "PageRank", "BL", "FL"), ]
    
    if (length(total_output_red) > 0 && nrow(structural) > 0) {
        best_struct <- structural[which.max(structural$pct_reduction), ]
        improvement <- best_struct$pct_reduction - total_output_red
        
        cat(sprintf("  %s: Total Output=%.3f%% -> %s=%.3f%% (delta=%.3f%%)\n",
            scenario_name, total_output_red, best_struct$method, best_struct$pct_reduction, improvement))
    }
}

cat("\nDone!\n")
