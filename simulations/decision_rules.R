###############################################################################
# decision_rules.R
#
# Core question: When can a CHEAP structural method substitute for the
# expensive DIIM simulation for identifying key economic sectors?
#
# Cheap methods (fixed ranking, no simulation needed):
#   1. PCA (multi-PC distance on H = A*L matrix)
#   2. Network Centrality (eigenvector + PageRank + betweenness on A_star)
#   3. Backward Linkage (column sums of Leontief inverse)
#   4. Forward Linkage (row sums of Leontief inverse)
#   5. Output-Weighted Linkage (total linkage × sector output)
#
# Expensive method (benchmark):
#   DIIM: run full simulation, rank by cumulative economic loss
#
# Approach:
#   - 2000 Monte Carlo random q0 vectors
#   - For each trial, compute performance ratio = cheap_reduction / DIIM_reduction
#   - Logistic regression to predict when ratio >= 0.95 ("close enough")
#   - Derive decision rules with ROC-optimal thresholds
#
# Run from project root:
#   Rscript simulations/decision_rules.R
###############################################################################

library(openxlsx)
library(igraph)
library(ggplot2)
library(reshape2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)


###############################################################################
# CHEAP METHOD RANKINGS (computed once from structure only)
###############################################################################

compute_cheap_rankings <- function(A, A_star, x) {
    n <- nrow(A)

    rankings <- list()

    # ── 1. PCA (2-PC distance on H matrix) ──
    I_minus_A <- diag(n) - A
    L <- solve(I_minus_A)
    H <- A %*% L
    eig <- eigen(H)
    pc_loadings <- Re(eig$vectors[, 1:min(2, n)])
    pca_dist <- sqrt(rowSums(pc_loadings^2))
    rankings[["PCA"]] <- order(pca_dist, decreasing = TRUE)

    # ── 2. Network Centrality ──
    g <- graph_from_adjacency_matrix(A_star,
        mode = "directed", weighted = TRUE,
        diag = FALSE
    )
    eig_cent <- eigen_centrality(as_undirected(g, mode = "collapse"))$vector
    pr <- page_rank(g)$vector
    bt <- betweenness(g, directed = TRUE, weights = 1 / E(g)$weight)
    norm_fn <- function(v) (v - min(v)) / max(max(v) - min(v), 1e-10)
    composite <- norm_fn(eig_cent) + norm_fn(pr) + norm_fn(bt)
    rankings[["Network Centrality"]] <- order(composite, decreasing = TRUE)

    # ── 3. Backward Linkage (column sums of L) ──
    BL <- colSums(L)
    rankings[["Backward Linkage"]] <- order(BL, decreasing = TRUE)

    # ── 4. Forward Linkage (row sums of L) ──
    FL <- rowSums(L)
    rankings[["Forward Linkage"]] <- order(FL, decreasing = TRUE)

    # ── 5. Output-Weighted Linkage ──
    total_linkage <- (BL / mean(BL) + FL / mean(FL)) / 2
    weighted <- total_linkage * as.vector(x)
    rankings[["Output-Weighted Linkage"]] <- order(weighted, decreasing = TRUE)

    return(rankings)
}


###############################################################################
# COMPUTE Q0 FEATURES (for logistic regression)
###############################################################################

compute_q0_features <- function(q0, cheap_rankings, k = 5) {
    n <- length(q0)
    total_q0 <- sum(q0)

    features <- list()

    for (method_name in names(cheap_rankings)) {
        top_k <- cheap_rankings[[method_name]][1:k]
        prefix <- gsub("[^a-zA-Z]", "", method_name) # clean name for column

        # Share of total q0 captured by this method's top-k sectors
        share <- sum(q0[top_k]) / total_q0
        features[[paste0(prefix, "_share")]] <- share

        # Average rank of this method's sectors in q0 (lower rank = higher q0)
        q0_ranks <- rank(-q0)
        avg_rank <- mean(q0_ranks[top_k]) / n # normalised 0-1
        features[[paste0(prefix, "_avgrank")]] <- avg_rank

        # Overlap: how many of top-k by q0 are in this method's top-k
        q0_topk <- order(q0, decreasing = TRUE)[1:k]
        overlap <- length(intersect(top_k, q0_topk)) / k
        features[[paste0(prefix, "_overlap")]] <- overlap
    }

    # General q0 features
    features[["q0_gini"]] <- 1 - 2 * sum((1:n) * sort(q0)) / (n * total_q0) +
        (n + 1) / n
    features[["q0_cv"]] <- sd(q0) / mean(q0)
    features[["q0_max_ratio"]] <- max(q0) / mean(q0)
    features[["q0_entropy"]] <- {
        p <- q0 / total_q0
        p <- p[p > 0]
        -sum(p * log(p)) / log(n) # normalised entropy
    }

    return(as.data.frame(features))
}


###############################################################################
# MONTE CARLO: Run trials for a given scenario
###############################################################################

run_monte_carlo <- function(scenario_name, data_loader,
                            lockdown_duration = 55, total_duration = 751,
                            days_in_year = 366, n_mc = 2000, k = 5) {
    cat(sprintf(
        "\n═══ %s: %d Monte Carlo trials (k=%d) ═══\n",
        scenario_name, n_mc, k
    ))

    data <- data_loader()
    A <- data$A
    x <- data$x
    q0_base <- data$q0
    q0_base[q0_base == 0] <- 1e-8
    c_star <- data$c_star
    A_star <- data$A_star
    n <- length(q0_base)
    max_q0 <- max(q0_base) * 2

    # Step 1: Compute cheap rankings (ONCE)
    cat("  Computing cheap rankings...\n")
    cheap_rankings <- compute_cheap_rankings(A, A_star, x)
    for (m in names(cheap_rankings)) {
        cat(sprintf(
            "    %-25s top-%d: [%s]\n", m, k,
            paste(cheap_rankings[[m]][1:k], collapse = ", ")
        ))
    }

    # Step 2: Monte Carlo
    cat("  Running Monte Carlo...\n")
    set.seed(42)

    all_results <- list()

    for (trial in 1:n_mc) {
        # Random q0
        q0_rand <- runif(n, min = 1e-6, max = max_q0)

        # DIIM baseline
        base_model <- DIIM(q0_rand, A_star, c_star, x, lockdown_duration,
            total_duration,
            days_in_year = days_in_year
        )
        base_loss <- base_model$total_economic_loss

        if (base_loss < 1e-6) next

        # DIIM: top-k sectors by economic loss
        max_el <- apply(base_model$EL_evolution, 1, max)
        diim_topk <- order(max_el, decreasing = TRUE)[1:k]
        diim_model <- DIIM(q0_rand, A_star, c_star, x, lockdown_duration,
            total_duration,
            key_sectors = diim_topk,
            days_in_year = days_in_year
        )
        diim_reduction <- base_loss - diim_model$total_economic_loss

        if (diim_reduction < 1e-10) next # skip if DIIM gives ~0 reduction

        # Each cheap method
        trial_row <- compute_q0_features(q0_rand, cheap_rankings, k)
        trial_row$trial <- trial
        trial_row$scenario <- scenario_name
        trial_row$base_loss <- base_loss
        trial_row$diim_reduction <- diim_reduction
        trial_row$diim_pct <- diim_reduction / base_loss * 100

        for (m in names(cheap_rankings)) {
            topk <- cheap_rankings[[m]][1:k]
            m_model <- DIIM(q0_rand, A_star, c_star, x, lockdown_duration,
                total_duration,
                key_sectors = topk,
                days_in_year = days_in_year
            )
            m_reduction <- base_loss - m_model$total_economic_loss
            col_red <- paste0(gsub("[^a-zA-Z]", "", m), "_reduction")
            col_ratio <- paste0(gsub("[^a-zA-Z]", "", m), "_ratio")
            col_wins <- paste0(gsub("[^a-zA-Z]", "", m), "_wins")
            col_close <- paste0(gsub("[^a-zA-Z]", "", m), "_close")

            trial_row[[col_red]] <- m_reduction
            trial_row[[col_ratio]] <- m_reduction / diim_reduction
            trial_row[[col_wins]] <- m_reduction >= diim_reduction
            trial_row[[col_close]] <- m_reduction / diim_reduction >= 0.95
        }

        all_results[[length(all_results) + 1]] <- trial_row

        if (trial %% 200 == 0) {
            # Quick progress
            pca_close_rate <- mean(sapply(all_results, function(r) r$PCA_close), na.rm = TRUE)
            cat(sprintf(
                "    %d/%d done (PCA close rate: %.1f%%)\n",
                trial, n_mc, pca_close_rate * 100
            ))
        }
    }

    mc_df <- bind_rows(all_results)
    cat(sprintf("  Valid trials: %d\n", nrow(mc_df)))

    return(list(mc_df = mc_df, cheap_rankings = cheap_rankings))
}


###############################################################################
# LOGISTIC REGRESSION + DECISION RULES
###############################################################################

build_decision_rules <- function(mc_df, scenario_name, methods) {
    cat(sprintf("\n--- Decision Rules for %s ---\n", scenario_name))

    rules <- list()

    for (m in methods) {
        clean_m <- gsub("[^a-zA-Z]", "", m)
        close_col <- paste0(clean_m, "_close")
        ratio_col <- paste0(clean_m, "_ratio")
        share_col <- paste0(clean_m, "_share")
        avgrank_col <- paste0(clean_m, "_avgrank")
        overlap_col <- paste0(clean_m, "_overlap")

        if (!close_col %in% names(mc_df)) next

        close_rate <- mean(mc_df[[close_col]], na.rm = TRUE)
        win_rate <- mean(mc_df[[paste0(clean_m, "_wins")]], na.rm = TRUE)
        mean_ratio <- mean(mc_df[[ratio_col]], na.rm = TRUE)

        cat(sprintf(
            "\n  %-25s close_rate=%.1f%% win_rate=%.1f%% mean_ratio=%.3f\n",
            m, close_rate * 100, win_rate * 100, mean_ratio
        ))

        # Logistic regression: predict "close" from q0 features
        feature_cols <- c(
            share_col, avgrank_col, overlap_col,
            "q0_gini", "q0_cv", "q0_max_ratio", "q0_entropy"
        )
        feature_cols <- feature_cols[feature_cols %in% names(mc_df)]

        if (length(feature_cols) == 0 || close_rate == 0 || close_rate == 1) {
            cat("    Skipping logistic regression (degenerate)\n")
            rules[[m]] <- list(
                method = m, close_rate = close_rate, win_rate = win_rate,
                mean_ratio = mean_ratio, model = NULL, threshold = NA,
                best_predictor = NA
            )
            next
        }

        formula_str <- paste(close_col, "~", paste(feature_cols, collapse = " + "))
        tryCatch(
            {
                lr_model <- glm(as.formula(formula_str), data = mc_df, family = binomial)

                # Find best single predictor
                coeffs <- summary(lr_model)$coefficients
                # Exclude intercept
                if (nrow(coeffs) > 1) {
                    p_vals <- coeffs[2:nrow(coeffs), 4]
                    best_pred <- names(which.min(p_vals))
                } else {
                    best_pred <- NA
                }

                cat(sprintf(
                    "    Best predictor: %s (p=%.4f)\n", best_pred,
                    min(coeffs[2:nrow(coeffs), 4])
                ))

                # Decision boundary on best predictor
                if (!is.na(best_pred) && best_pred %in% names(mc_df)) {
                    pred_vals <- mc_df[[best_pred]]
                    close_vals <- mc_df[[close_col]]

                    # Sweep thresholds to find Youden's J optimal
                    thresholds <- quantile(pred_vals, probs = seq(0.05, 0.95, 0.05))
                    best_j <- -Inf
                    best_thresh <- NA

                    for (thresh in thresholds) {
                        if (grepl("avgrank", best_pred)) {
                            predicted <- pred_vals < thresh # lower rank = better
                        } else {
                            predicted <- pred_vals > thresh # higher = better
                        }
                        tp <- sum(predicted & close_vals)
                        tn <- sum(!predicted & !close_vals)
                        fp <- sum(predicted & !close_vals)
                        fn <- sum(!predicted & close_vals)

                        sens <- tp / max(tp + fn, 1)
                        spec <- tn / max(tn + fp, 1)
                        j <- sens + spec - 1

                        if (j > best_j) {
                            best_j <- j
                            best_thresh <- thresh
                        }
                    }

                    direction <- if (grepl("avgrank", best_pred)) "<" else ">"
                    cat(sprintf(
                        "    Decision rule: Use %s if %s %s %.3f (Youden J=%.3f)\n",
                        m, best_pred, direction, best_thresh, best_j
                    ))

                    rules[[m]] <- list(
                        method = m, close_rate = close_rate, win_rate = win_rate,
                        mean_ratio = mean_ratio, model = lr_model,
                        best_predictor = best_pred, threshold = best_thresh,
                        direction = direction, youden_j = best_j
                    )
                } else {
                    rules[[m]] <- list(
                        method = m, close_rate = close_rate, win_rate = win_rate,
                        mean_ratio = mean_ratio, model = lr_model,
                        best_predictor = best_pred, threshold = NA
                    )
                }
            },
            error = function(e) {
                cat(sprintf("    Logistic regression failed: %s\n", e$message))
                rules[[m]] <<- list(
                    method = m, close_rate = close_rate, win_rate = win_rate,
                    mean_ratio = mean_ratio, model = NULL
                )
            }
        )
    }

    return(rules)
}


###############################################################################
# PLOTS
###############################################################################

generate_plots <- function(mc_df, rules, scenario_name, prefix) {
    methods <- c(
        "PCA", "NetworkCentrality", "BackwardLinkage",
        "ForwardLinkage", "OutputWeightedLinkage"
    )
    method_labels <- c(
        "PCA", "Network Centrality", "Backward Linkage",
        "Forward Linkage", "Output-Weighted Linkage"
    )

    # ── Plot 1: Performance ratio distributions ──
    ratio_data <- data.frame()
    for (i in seq_along(methods)) {
        col <- paste0(methods[i], "_ratio")
        if (col %in% names(mc_df)) {
            ratio_data <- rbind(ratio_data, data.frame(
                method = method_labels[i],
                ratio = mc_df[[col]],
                stringsAsFactors = FALSE
            ))
        }
    }

    if (nrow(ratio_data) > 0) {
        p1 <- ggplot(ratio_data, aes(x = ratio, fill = method)) +
            geom_density(alpha = 0.5) +
            geom_vline(xintercept = 0.95, linetype = "dashed", color = "red", linewidth = 0.8) +
            geom_vline(xintercept = 1.0, linetype = "dotted", color = "darkgreen", linewidth = 0.8) +
            annotate("text",
                x = 0.95, y = Inf, label = "95% threshold",
                vjust = 2, hjust = 1.1, color = "red", size = 3
            ) +
            scale_fill_brewer(palette = "Set2") +
            labs(
                title = sprintf("%s: Performance Ratio Distributions", scenario_name),
                subtitle = "Ratio = cheap method reduction / DIIM reduction",
                x = "Performance Ratio", y = "Density", fill = "Method"
            ) +
            theme_minimal() +
            theme(
                plot.title = element_text(face = "bold", size = 14),
                legend.position = "bottom"
            )

        ggsave(file.path(results_dir, paste0(prefix, "_ratio_distributions.png")),
            p1,
            width = 12, height = 7
        )
        cat(sprintf("  -> Saved %s_ratio_distributions.png\n", prefix))
    }

    # ── Plot 2: Close rates bar chart ──
    close_data <- data.frame()
    for (i in seq_along(methods)) {
        col <- paste0(methods[i], "_close")
        if (col %in% names(mc_df)) {
            close_data <- rbind(close_data, data.frame(
                method = method_labels[i],
                close_rate = mean(mc_df[[col]], na.rm = TRUE),
                win_rate = mean(mc_df[[paste0(methods[i], "_wins")]], na.rm = TRUE),
                stringsAsFactors = FALSE
            ))
        }
    }

    if (nrow(close_data) > 0) {
        close_long <- melt(close_data,
            id.vars = "method",
            variable.name = "metric", value.name = "rate"
        )
        close_long$metric <- ifelse(close_long$metric == "close_rate",
            "≥95% of DIIM", "Beats DIIM"
        )

        p2 <- ggplot(close_long, aes(x = reorder(method, rate), y = rate, fill = metric)) +
            geom_bar(stat = "identity", position = "dodge", alpha = 0.85) +
            coord_flip() +
            scale_y_continuous(labels = scales::percent) +
            scale_fill_manual(values = c(
                "≥95% of DIIM" = "#3498DB",
                "Beats DIIM" = "#2ECC71"
            )) +
            labs(
                title = sprintf("%s: How Often Each Cheap Method Matches DIIM", scenario_name),
                x = "", y = "Rate", fill = ""
            ) +
            theme_minimal() +
            theme(
                plot.title = element_text(face = "bold", size = 14),
                legend.position = "bottom"
            )

        ggsave(file.path(results_dir, paste0(prefix, "_close_rates.png")),
            p2,
            width = 10, height = 6
        )
        cat(sprintf("  -> Saved %s_close_rates.png\n", prefix))
    }

    # ── Plot 3: Decision boundary for best method ──
    for (rule in rules) {
        if (is.null(rule$best_predictor) || is.na(rule$best_predictor)) next
        if (is.na(rule$threshold)) next

        clean_m <- gsub("[^a-zA-Z]", "", rule$method)
        close_col <- paste0(clean_m, "_close")
        pred_col <- rule$best_predictor

        if (!pred_col %in% names(mc_df) || !close_col %in% names(mc_df)) next

        p3 <- ggplot(mc_df, aes_string(x = pred_col, y = paste0(clean_m, "_ratio"))) +
            geom_point(aes_string(color = close_col), alpha = 0.3, size = 1) +
            geom_vline(
                xintercept = rule$threshold, linetype = "dashed",
                color = "red", linewidth = 1
            ) +
            geom_hline(yintercept = 0.95, linetype = "dotted", color = "blue") +
            scale_color_manual(
                values = c("TRUE" = "#2ECC71", "FALSE" = "#E74C3C"),
                labels = c("TRUE" = "≥95%", "FALSE" = "<95%"),
                name = "Close to DIIM"
            ) +
            labs(
                title = sprintf("%s: %s Decision Boundary", scenario_name, rule$method),
                subtitle = sprintf(
                    "Rule: Use %s if %s %s %.3f",
                    rule$method, pred_col, rule$direction, rule$threshold
                ),
                x = pred_col, y = "Performance Ratio"
            ) +
            theme_minimal() +
            theme(plot.title = element_text(face = "bold", size = 13))

        fname <- paste0(prefix, "_boundary_", tolower(gsub("[^a-zA-Z]", "", rule$method)), ".png")
        ggsave(file.path(results_dir, fname), p3, width = 10, height = 7)
        cat(sprintf("  -> Saved %s\n", fname))
    }
}


###############################################################################
# MAIN
###############################################################################
cat("=== Decision Rules: Cheap Methods vs DIIM ===\n")

method_names <- c(
    "PCA", "Network Centrality", "Backward Linkage",
    "Forward Linkage", "Output-Weighted Linkage"
)

# ── COVID-19 ──
covid_mc <- run_monte_carlo("COVID-19", download_data,
    lockdown_duration = 55, total_duration = 751,
    days_in_year = 366, n_mc = 2000, k = 5
)
covid_rules <- build_decision_rules(covid_mc$mc_df, "COVID-19", method_names)

cat("\n--- COVID-19 Plots ---\n")
generate_plots(covid_mc$mc_df, covid_rules, "COVID-19", "covid_dr")

# ── Manpower ──
manpower_mc <- run_monte_carlo("Manpower", download_manpower_data,
    lockdown_duration = 55, total_duration = 751,
    days_in_year = 365, n_mc = 2000, k = 5
)
manpower_rules <- build_decision_rules(manpower_mc$mc_df, "Manpower", method_names)

cat("\n--- Manpower Plots ---\n")
generate_plots(manpower_mc$mc_df, manpower_rules, "Manpower", "manpower_dr")

# ── Save combined data ──
combined_mc <- bind_rows(covid_mc$mc_df, manpower_mc$mc_df)
write.csv(combined_mc, file.path(results_dir, "decision_rules_mc_data.csv"),
    row.names = FALSE
)
cat("\n  -> Saved decision_rules_mc_data.csv\n")


###############################################################################
# SUMMARY REPORT
###############################################################################
cat("\n══════════════════════════════════════════════════════\n")
cat("  DECISION RULES SUMMARY\n")
cat("══════════════════════════════════════════════════════\n\n")

summary_lines <- character()

for (sc_name in c("COVID-19", "Manpower")) {
    rules <- if (sc_name == "COVID-19") covid_rules else manpower_rules
    cat(sprintf("--- %s ---\n", sc_name))
    summary_lines <- c(summary_lines, sprintf("--- %s ---", sc_name))

    for (m in names(rules)) {
        r <- rules[[m]]
        rule_str <- if (!is.na(r$threshold) && !is.null(r$best_predictor) && !is.na(r$best_predictor)) {
            sprintf(
                "Use if %s %s %.3f (J=%.3f)", r$best_predictor, r$direction,
                r$threshold, r$youden_j
            )
        } else {
            "No clear decision boundary"
        }

        line <- sprintf(
            "  %-25s close=%.1f%% wins=%.1f%% ratio=%.3f | %s",
            m, r$close_rate * 100, r$win_rate * 100, r$mean_ratio, rule_str
        )
        cat(line, "\n")
        summary_lines <- c(summary_lines, line)
    }
    cat("\n")
    summary_lines <- c(summary_lines, "")
}

# Save summary
writeLines(summary_lines, file.path(results_dir, "decision_rules_summary.txt"))
cat("  -> Saved decision_rules_summary.txt\n")

cat("\n══════════════════════════════════════════════════════\n")
cat("All decision rules analysis complete!\n")
cat(sprintf(
    "Output: %d files in simulations/results/\n",
    length(list.files(results_dir))
))
cat("══════════════════════════════════════════════════════\n")
