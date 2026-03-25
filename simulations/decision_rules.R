if (!file.exists("functions.R")) setwd("..")

library(openxlsx)
library(igraph)
library(ggplot2)
library(reshape2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

# full method names (used for display and as keys)
method_names <- c("Total Output", "PCA x xi", "PageRank x xi",
                  "BL x xi", "FL x xi")

# short column prefixes (letters only, no spaces/hyphens)
method_prefixes <- c(
    "Total Output"  = "TotalOutput",
    "PCA x xi"      = "PCAxi",
    "PageRank x xi" = "PageRankxi",
    "BL x xi"       = "BLxi",
    "FL x xi"       = "FLxi"
)

# short labels for plots
method_labels <- c("Total Output", "PCA x xi", "PageRank x xi",
                   "BL x xi", "FL x xi")

compute_q0_features <- function(q0, simplified_rankings, k = 5) {
    num_sectors <- length(q0)
    total_q0 <- sum(q0)
    
    # Store the calculated features in a list
    calculated_features <- list()

    # Iterate over every simplified method to calculate how well its top-k sectors align with q0
    for (method_name in names(simplified_rankings)) {
        top_k_sectors <- simplified_rankings[[method_name]][1:k]
        prefix <- method_prefixes[method_name]

        # Metric 1: Share of total q0 captured by this method's top-k sectors
        calculated_features[[paste0(prefix, "_share")]] <- sum(q0[top_k_sectors]) / total_q0

        # Metric 2: Average rank of the method's sectors in the actual q0 ordering (normalised 0-1)
        q0_ranks <- rank(-q0)
        calculated_features[[paste0(prefix, "_avgrank")]] <- mean(q0_ranks[top_k_sectors]) / num_sectors

        # Metric 3: Overlap - How many of the top-k highest q0 sectors match the method's top-k
        q0_topk <- order(q0, decreasing = TRUE)[1:k]
        calculated_features[[paste0(prefix, "_overlap")]] <- length(intersect(top_k_sectors, q0_topk)) / k
    }

    # Add general statistical distribution features of the q0 array
    calculated_features[["q0_gini"]] <- 1 - 2 * sum((1:num_sectors) * sort(q0)) / (num_sectors * total_q0) + (num_sectors + 1) / num_sectors
    calculated_features[["q0_cv"]]   <- sd(q0) / mean(q0)
    calculated_features[["q0_max_ratio"]] <- max(q0) / mean(q0)
    
    # Calculate normalized entropy of q0
    probability_distribution <- q0 / total_q0
    probability_distribution <- probability_distribution[probability_distribution > 0]
    normalized_entropy <- -sum(probability_distribution * log(probability_distribution)) / log(num_sectors)
    calculated_features[["q0_entropy"]] <- normalized_entropy

    return(as.data.frame(calculated_features))
}

run_monte_carlo <- function(scenario_name, data_loader,
                            lockdown_duration = 55, total_duration = 751,
                            days_in_year = 366, n_mc = 2000, k = 5) {
    cat(sprintf("\n%s: %d Monte Carlo trials (k=%d)\n", scenario_name, n_mc, k))

    # 1. Load context data
    data <- data_loader()
    A      <- data$A
    x      <- data$x
    c_star <- data$c_star
    A_star <- data$A_star
    
    q0_base <- data$q0
    q0_base[q0_base == 0] <- 1e-8
    num_sectors <- length(q0_base)
    max_q0 <- max(q0_base) * 2

    # 2. Compute the static rankings for the simplified methods upfront
    cat("  Computing simplified rankings...\n")
    simplified_rankings <- compute_simplified_rankings(A, A_star, x)
    for (method_name in names(simplified_rankings)) {
        cat(sprintf("    %-25s top-%d: [%s]\n", method_name, k,
            paste(simplified_rankings[[method_name]][1:k], collapse = ", ")))
    }

    # 3. Begin Monte Carlo iterations
    cat("  Running Monte Carlo...\n")
    set.seed(42)
    all_results <- list()

    for (trial in 1:n_mc) {
        # Perturb the real q0: multiply each sector by a log-normal noise factor
        # This preserves the relative structure of the real scenario while varying intensity
        noise <- exp(rnorm(num_sectors, mean = 0, sd = 0.5))
        random_q0 <- pmax(q0_base * noise, 1e-6)
        random_q0 <- pmin(random_q0, 1)  # cap at 1 (100% inoperability)

        # Evaluate the DIIM baseline with no key sectors protected
        base_model <- DIIM(random_q0, A_star, c_star, x,
            lockdown_duration, total_duration, days_in_year = days_in_year)
        base_loss <- base_model$total_economic_loss
        
        if (base_loss < 1e-6) next # Skip degenerate trials

        # Evaluate the Gold Standard (DIIM using its own internal top-k logic)
        max_el <- apply(base_model$EL_evolution, 1, max)
        diim_topk <- order(max_el, decreasing = TRUE)[1:k]
        
        diim_model <- DIIM(random_q0, A_star, c_star, x,
            lockdown_duration, total_duration,
            key_sectors = diim_topk, days_in_year = days_in_year)
            
        diim_reduction <- base_loss - diim_model$total_economic_loss
        if (diim_reduction < 1e-10) next

        # Start building the results row for this trial
        trial_row <- compute_q0_features(random_q0, simplified_rankings, k)
        trial_row$trial     <- trial
        trial_row$scenario  <- scenario_name
        trial_row$base_loss <- base_loss
        trial_row$diim_reduction <- diim_reduction
        trial_row$diim_pct  <- diim_reduction / base_loss * 100

        # Evaluate each computed simplified method
        for (method_name in names(simplified_rankings)) {
            topk_sectors <- simplified_rankings[[method_name]][1:k]
            
            method_model <- DIIM(random_q0, A_star, c_star, x,
                lockdown_duration, total_duration,
                key_sectors = topk_sectors, days_in_year = days_in_year)
                
            method_reduction <- base_loss - method_model$total_economic_loss

            # Store comparative states vs gold standard
            prefix <- method_prefixes[method_name]
            trial_row[[paste0(prefix, "_reduction")]] <- method_reduction
            trial_row[[paste0(prefix, "_ratio")]]     <- method_reduction / diim_reduction
            trial_row[[paste0(prefix, "_wins")]]      <- method_reduction >= diim_reduction
            trial_row[[paste0(prefix, "_close")]]     <- method_reduction / diim_reduction >= 0.80
        }

        all_results[[length(all_results) + 1]] <- trial_row

        if (trial %% 200 == 0) {
            pca_close_rate <- mean(sapply(all_results, function(r) r$TotalOutput_close), na.rm = TRUE)
            cat(sprintf("    %d/%d done (TotalOutput close rate: %.1f%%)\n",
                trial, n_mc, pca_close_rate * 100))
        }
    }

    monte_carlo_df <- bind_rows(all_results)
    cat(sprintf("  Valid trials: %d\n", nrow(monte_carlo_df)))
    return(list(mc_df = monte_carlo_df, simplified_rankings = simplified_rankings))
}

build_decision_rules <- function(mc_df, scenario_name, methods) {
    cat(sprintf("\nDecision Rules for %s\n", scenario_name))
    rules <- list()

    for (m in methods) {
        prefix      <- method_prefixes[m]
        close_col   <- paste0(prefix, "_close")
        ratio_col   <- paste0(prefix, "_ratio")
        share_col   <- paste0(prefix, "_share")
        avgrank_col <- paste0(prefix, "_avgrank")
        overlap_col <- paste0(prefix, "_overlap")

        if (!close_col %in% names(mc_df)) next

        close_rate <- mean(mc_df[[close_col]], na.rm = TRUE)
        win_rate   <- mean(mc_df[[paste0(prefix, "_wins")]], na.rm = TRUE)
        mean_ratio <- mean(mc_df[[ratio_col]], na.rm = TRUE)

        cat(sprintf("\n  %-25s close=%.1f%% wins=%.1f%% ratio=%.3f\n",
            m, close_rate * 100, win_rate * 100, mean_ratio))

        # feature columns for logistic regression
        feature_cols <- c(share_col, avgrank_col, overlap_col,
                          "q0_gini", "q0_cv", "q0_max_ratio", "q0_entropy")
        feature_cols <- feature_cols[feature_cols %in% names(mc_df)]

        if (length(feature_cols) == 0 || close_rate == 0 || close_rate == 1) {
            cat("    Skipping logistic regression (degenerate)\n")
            rules[[m]] <- list(method = m, close_rate = close_rate,
                win_rate = win_rate, mean_ratio = mean_ratio,
                model = NULL, threshold = NA, best_predictor = NA)
            next
        }

        formula_str <- paste(close_col, "~", paste(feature_cols, collapse = " + "))
        tryCatch({
            lr_model <- glm(as.formula(formula_str), data = mc_df, family = binomial)

            # find the most significant predictor (lowest p-value)
            coeffs <- summary(lr_model)$coefficients
            if (nrow(coeffs) > 1) {
                p_vals    <- coeffs[2:nrow(coeffs), 4]
                best_pred <- names(which.min(p_vals))
            } else {
                best_pred <- NA
            }

            cat(sprintf("    Best predictor: %s (p=%.4f)\n",
                best_pred, min(coeffs[2:nrow(coeffs), 4])))

            # sweep thresholds on the best predictor to find Youden's J optimal
            if (!is.na(best_pred) && best_pred %in% names(mc_df)) {
                pred_vals  <- mc_df[[best_pred]]
                close_vals <- mc_df[[close_col]]
                thresholds <- quantile(pred_vals, probs = seq(0.05, 0.95, 0.05))
                best_j     <- -Inf
                best_thresh <- NA

                for (thresh in thresholds) {
                    if (grepl("avgrank", best_pred)) {
                        predicted <- pred_vals < thresh  # lower rank = better
                    } else {
                        predicted <- pred_vals > thresh
                    }
                    tp <- sum(predicted & close_vals)
                    tn <- sum(!predicted & !close_vals)
                    fp <- sum(predicted & !close_vals)
                    fn <- sum(!predicted & close_vals)

                    sens <- tp / max(tp + fn, 1)
                    spec <- tn / max(tn + fp, 1)
                    j    <- sens + spec - 1

                    if (j > best_j) {
                        best_j      <- j
                        best_thresh <- thresh
                    }
                }

                direction <- if (grepl("avgrank", best_pred)) "<" else ">"
                cat(sprintf("    Rule: Use %s if %s %s %.3f (Youden J=%.3f)\n",
                    m, best_pred, direction, best_thresh, best_j))

                rules[[m]] <- list(method = m, close_rate = close_rate,
                    win_rate = win_rate, mean_ratio = mean_ratio,
                    model = lr_model, best_predictor = best_pred,
                    threshold = best_thresh, direction = direction,
                    youden_j = best_j)
            } else {
                rules[[m]] <- list(method = m, close_rate = close_rate,
                    win_rate = win_rate, mean_ratio = mean_ratio,
                    model = lr_model, best_predictor = best_pred,
                    threshold = NA)
            }
        }, error = function(e) {
            cat(sprintf("    Logistic regression failed: %s\n", e$message))
            rules[[m]] <<- list(method = m, close_rate = close_rate,
                win_rate = win_rate, mean_ratio = mean_ratio, model = NULL)
        })
    }

    return(rules)
}

generate_plots <- function(mc_df, rules, scenario_name, prefix) {
    clean_methods <- c("TotalOutput", "PCAxi", "PageRankxi",
                       "BLxi", "FLxi")

    # --- Plot 1: Performance ratio distributions ---
    ratio_data <- data.frame()
    for (i in seq_along(clean_methods)) {
        col <- paste0(clean_methods[i], "_ratio")
        if (col %in% names(mc_df)) {
            ratio_data <- rbind(ratio_data, data.frame(
                method = method_labels[i],
                ratio  = mc_df[[col]],
                stringsAsFactors = FALSE))
        }
    }

    if (nrow(ratio_data) > 0) {
        p1 <- ggplot(ratio_data, aes(x = ratio, fill = method)) +
            geom_density(alpha = 0.5) +
            geom_vline(xintercept = 0.80, linetype = "dashed", color = "red", linewidth = 0.8) +
            geom_vline(xintercept = 1.0, linetype = "dotted", color = "darkgreen", linewidth = 0.8) +
            annotate("text", x = 0.80, y = Inf, label = "80% threshold",
                     vjust = 2, hjust = 1.1, color = "red", size = 3) +
            scale_fill_brewer(palette = "Set2") +
            labs(title = sprintf("%s: Performance Ratio Distributions", scenario_name),
                 subtitle = "Ratio = simplified method reduction / DIIM reduction",
                 x = "Performance Ratio", y = "Density", fill = "Method") +
            theme_minimal() +
            theme(plot.title = element_text(face = "bold", size = 14),
                  legend.position = "bottom")
        ggsave(file.path(results_dir, paste0(prefix, "_ratio_distributions.png")),
               p1, width = 12, height = 7)
        cat(sprintf("  Saved %s_ratio_distributions.png\n", prefix))
    }

    # --- Plot 2: Close rates bar chart ---
    close_data <- data.frame()
    for (i in seq_along(clean_methods)) {
        col <- paste0(clean_methods[i], "_close")
        if (col %in% names(mc_df)) {
            close_data <- rbind(close_data, data.frame(
                method     = method_labels[i],
                close_rate = mean(mc_df[[col]], na.rm = TRUE),
                win_rate   = mean(mc_df[[paste0(clean_methods[i], "_wins")]], na.rm = TRUE),
                stringsAsFactors = FALSE))
        }
    }

    if (nrow(close_data) > 0) {
        close_long <- melt(close_data, id.vars = "method",
                           variable.name = "metric", value.name = "rate")
        close_long$metric <- ifelse(close_long$metric == "close_rate",
                                    ">=80% of DIIM", "Beats DIIM")

        p2 <- ggplot(close_long, aes(x = reorder(method, rate), y = rate, fill = metric)) +
            geom_bar(stat = "identity", position = "dodge", alpha = 0.85) +
            coord_flip() +
            scale_y_continuous(labels = scales::percent) +
            scale_fill_manual(values = c(">=80% of DIIM" = "#3498DB",
                                         "Beats DIIM" = "#2ECC71")) +
            labs(title = sprintf("%s: How Often Each Simplified Method Matches DIIM", scenario_name),
                 x = "", y = "Rate", fill = "") +
            theme_minimal() +
            theme(plot.title = element_text(face = "bold", size = 14),
                  legend.position = "bottom")
        ggsave(file.path(results_dir, paste0(prefix, "_close_rates.png")),
               p2, width = 10, height = 6)
        cat(sprintf("  Saved %s_close_rates.png\n", prefix))
    }

    # --- Plot 3: Decision boundary for each method with a rule ---
    for (rule in rules) {
        if (is.null(rule$best_predictor) || is.na(rule$best_predictor)) next
        if (is.na(rule$threshold)) next

        prefix_m  <- method_prefixes[rule$method]
        close_col <- paste0(prefix_m, "_close")
        pred_col  <- rule$best_predictor

        if (!pred_col %in% names(mc_df) || !close_col %in% names(mc_df)) next

        p3 <- ggplot(mc_df, aes_string(x = pred_col, y = paste0(prefix_m, "_ratio"))) +
            geom_point(aes_string(color = close_col), alpha = 0.3, size = 1) +
            geom_vline(xintercept = rule$threshold, linetype = "dashed",
                       color = "red", linewidth = 1) +
            geom_hline(yintercept = 0.80, linetype = "dotted", color = "blue") +
            scale_color_manual(
                values = c("TRUE" = "#2ECC71", "FALSE" = "#E74C3C"),
                labels = c("TRUE" = ">=80%", "FALSE" = "<80%"),
                name = "Close to DIIM") +
            labs(title = sprintf("%s: %s Decision Boundary", scenario_name, rule$method),
                 subtitle = sprintf("Rule: Use %s if %s %s %.3f",
                     rule$method, pred_col, rule$direction, rule$threshold),
                 x = pred_col, y = "Performance Ratio") +
            theme_minimal() +
            theme(plot.title = element_text(face = "bold", size = 13))

        fname <- paste0(prefix, "_boundary_", tolower(prefix_m), ".png")
        ggsave(file.path(results_dir, fname), p3, width = 10, height = 7)
        cat(sprintf("  Saved %s\n", fname))
    }
}

k_values <- c(3, 5, 7, 10, 12)
all_mc_results <- list()

for (k_val in k_values) {
    cat(sprintf("\n========== k = %d ==========\n", k_val))

    covid_mc <- run_monte_carlo("COVID-19", download_data,
        lockdown_duration = 55, total_duration = 751,
        days_in_year = 366, n_mc = 2000, k = k_val)
    covid_mc$mc_df$k <- k_val
    all_mc_results[[paste0("covid_k", k_val)]] <- covid_mc$mc_df

    manpower_mc <- run_monte_carlo("Manpower", download_manpower_data,
        lockdown_duration = 55, total_duration = 751,
        days_in_year = 365, n_mc = 2000, k = k_val)
    manpower_mc$mc_df$k <- k_val
    all_mc_results[[paste0("manpower_k", k_val)]] <- manpower_mc$mc_df

    # Build decision rules and generate plots for this k value
    covid_rules <- build_decision_rules(covid_mc$mc_df, sprintf("COVID-19 (k=%d)", k_val), method_labels)
    generate_plots(covid_mc$mc_df, covid_rules, sprintf("COVID-19 (k=%d)", k_val), sprintf("covid_k%d", k_val))

    manpower_rules <- build_decision_rules(manpower_mc$mc_df, sprintf("Manpower (k=%d)", k_val), method_labels)
    generate_plots(manpower_mc$mc_df, manpower_rules, sprintf("Manpower (k=%d)", k_val), sprintf("manpower_k%d", k_val))
}

# Combine all results
combined_mc <- bind_rows(all_mc_results)
write.csv(combined_mc, file.path(results_dir, "decision_rules_mc_data.csv"), row.names = FALSE)
cat("Saved decision_rules_mc_data.csv\n\n")

# Build summary table: mean ratio by method, scenario, and k
cat("SUMMARY: Mean Performance Ratio (simplified / DIIM) by k\n\n")

clean_methods <- c("TotalOutput", "PCAxi", "PageRankxi", "BLxi", "FLxi")
summary_lines <- character()

for (sc_name in c("COVID-19", "Manpower")) {
    cat(sprintf("--- %s ---\n", sc_name))
    summary_lines <- c(summary_lines, sprintf("--- %s ---", sc_name))

    sc_data <- combined_mc[combined_mc$scenario == sc_name, ]

    # Header
    header <- sprintf("  %-20s %s", "Method", paste(sprintf("k=%-4d", k_values), collapse = "  "))
    cat(header, "\n")
    summary_lines <- c(summary_lines, header)

    for (i in seq_along(clean_methods)) {
        prefix <- clean_methods[i]
        ratio_col <- paste0(prefix, "_ratio")
        close_col <- paste0(prefix, "_close")

        ratios <- c()
        for (k_val in k_values) {
            k_data <- sc_data[sc_data$k == k_val, ]
            if (ratio_col %in% names(k_data)) {
                ratios <- c(ratios, mean(k_data[[ratio_col]], na.rm = TRUE))
            } else {
                ratios <- c(ratios, NA)
            }
        }

        line <- sprintf("  %-20s %s", method_labels[i],
            paste(sprintf("%.3f ", ratios), collapse = "  "))
        cat(line, "\n")
        summary_lines <- c(summary_lines, line)
    }

    # Also print close rates
    cat("\n  Close rates (>=80% of DIIM):\n")
    summary_lines <- c(summary_lines, "  Close rates (>=80% of DIIM):")

    for (i in seq_along(clean_methods)) {
        prefix <- clean_methods[i]
        close_col <- paste0(prefix, "_close")

        close_rates <- c()
        for (k_val in k_values) {
            k_data <- sc_data[sc_data$k == k_val, ]
            if (close_col %in% names(k_data)) {
                close_rates <- c(close_rates, mean(k_data[[close_col]], na.rm = TRUE) * 100)
            } else {
                close_rates <- c(close_rates, NA)
            }
        }

        line <- sprintf("  %-20s %s", method_labels[i],
            paste(sprintf("%5.1f%%", close_rates), collapse = "  "))
        cat(line, "\n")
        summary_lines <- c(summary_lines, line)
    }
    cat("\n")
    summary_lines <- c(summary_lines, "")
}

# save summary
writeLines(summary_lines, file.path(results_dir, "decision_rules_summary.txt"))
cat("Saved decision_rules_summary.txt\n")
cat("Done!\n")
