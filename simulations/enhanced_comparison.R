###############################################################################
# enhanced_comparison.R
#
# Enhanced comparison with two key improvements:
#   1. Dynamic PCA sector ranking using Euclidean distance in multi-PC space
#      (PC1-only, PC1+PC2, PC1+PC2+PC3, PC1..PC5)
#   2. Variable top-k sweep (k = 3, 5, 7, 10, 15, 20)
#
# Run from project root:
#   Rscript simulations/enhanced_comparison.R
###############################################################################

library(openxlsx)
library(igraph)
library(ggplot2)
library(reshape2)
library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)


# ─────────────────────────────────────────────────────────────────────────────
# Dynamic PCA: rank sectors by distance from origin in multi-PC space
# ─────────────────────────────────────────────────────────────────────────────
pca_rank_sectors <- function(A_matrix, n_pcs = 2) {
    # Construct H = A %*% L  (influence matrix)
    n <- nrow(A_matrix)
    I_minus_A <- diag(n) - A_matrix
    L <- solve(I_minus_A)
    H <- A_matrix %*% L

    # Eigendecomposition of H
    eig <- eigen(H)
    eigenvalues <- Re(eig$values)
    eigenvectors <- Re(eig$vectors)

    # Limit n_pcs to available components
    n_pcs <- min(n_pcs, ncol(eigenvectors))

    # Compute loadings on the first n_pcs principal components
    loadings <- eigenvectors[, 1:n_pcs, drop = FALSE]

    # Euclidean distance from origin in PC space
    distances <- sqrt(rowSums(loadings^2))

    # Rank sectors by distance (highest = most important)
    ranked_sectors <- order(distances, decreasing = TRUE)

    # Variance explained by each PC used
    var_explained <- eigenvalues / sum(abs(eigenvalues))
    cumulative_var <- cumsum(var_explained[1:n_pcs])

    return(list(
        ranked_sectors = ranked_sectors,
        distances = distances,
        loadings = loadings,
        var_explained = var_explained[1:n_pcs],
        cumulative_var = tail(cumulative_var, 1),
        n_pcs = n_pcs
    ))
}


# ─────────────────────────────────────────────────────────────────────────────
# DIIM sector ranking: top-k by economic loss
# ─────────────────────────────────────────────────────────────────────────────
diim_rank_sectors <- function(q0, A_star, c_star, x,
                              lockdown_duration, total_duration,
                              days_in_year = 366) {
    model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
        days_in_year = days_in_year
    )
    max_el <- apply(model$EL_evolution, 1, max)
    ranked_sectors <- order(max_el, decreasing = TRUE)
    return(list(ranked_sectors = ranked_sectors, max_el = max_el))
}


# ─────────────────────────────────────────────────────────────────────────────
# Sensitivity sector ranking: marginal loss reduction per sector
# ─────────────────────────────────────────────────────────────────────────────
sensitivity_rank_sectors <- function(q0, A_star, c_star, x,
                                     lockdown_duration, total_duration,
                                     days_in_year = 366) {
    n <- length(q0)
    base_model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
        days_in_year = days_in_year
    )
    base_loss <- base_model$total_economic_loss

    marginal_reductions <- numeric(n)
    for (i in 1:n) {
        q0_test <- q0
        q0_test[i] <- q0_test[i] * 0.9
        model_test <- DIIM(q0_test, A_star, c_star, x, lockdown_duration,
            total_duration,
            days_in_year = days_in_year
        )
        marginal_reductions[i] <- base_loss - model_test$total_economic_loss
    }

    ranked_sectors <- order(marginal_reductions, decreasing = TRUE)
    return(list(
        ranked_sectors = ranked_sectors,
        marginal_reductions = marginal_reductions
    ))
}


# ─────────────────────────────────────────────────────────────────────────────
# Network Centrality sector ranking
# ─────────────────────────────────────────────────────────────────────────────
centrality_rank_sectors <- function(A_star) {
    g <- graph_from_adjacency_matrix(A_star,
        mode = "directed", weighted = TRUE,
        diag = FALSE
    )
    eig_cent <- eigen_centrality(as_undirected(g, mode = "collapse"))$vector
    pr <- page_rank(g)$vector
    bt <- betweenness(g, directed = TRUE, weights = 1 / E(g)$weight)

    norm_fn <- function(x) (x - min(x)) / max(max(x) - min(x), 1e-10)
    composite <- norm_fn(eig_cent) + norm_fn(pr) + norm_fn(bt)

    ranked_sectors <- order(composite, decreasing = TRUE)
    return(list(ranked_sectors = ranked_sectors, composite = composite))
}


# ─────────────────────────────────────────────────────────────────────────────
# Evaluate: given ranked sectors, pick top-k and run DIIM intervention
# ─────────────────────────────────────────────────────────────────────────────
evaluate_topk <- function(ranked_sectors, k, q0, A_star, c_star, x,
                          lockdown_duration, total_duration,
                          base_loss, days_in_year = 366) {
    key_sectors <- ranked_sectors[1:k]
    model_int <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
        key_sectors = key_sectors, days_in_year = days_in_year
    )
    reduction <- base_loss - model_int$total_economic_loss
    pct_reduction <- reduction / base_loss * 100
    return(list(
        k = k, sectors = key_sectors,
        loss = model_int$total_economic_loss,
        reduction = reduction,
        pct_reduction = pct_reduction
    ))
}


###############################################################################
# MAIN: Run comparison for a given scenario
###############################################################################
run_scenario <- function(scenario_name, data_loader,
                         lockdown_duration = 55, total_duration = 751,
                         days_in_year = 366,
                         k_values = c(3, 5, 7, 10, 15, 20)) {
    cat(sprintf("\n═══════════════════════════════════════\n"))
    cat(sprintf("  %s Scenario\n", scenario_name))
    cat(sprintf("═══════════════════════════════════════\n"))

    data <- data_loader()
    A <- data$A
    x <- data$x
    q0 <- data$q0
    q0[q0 == 0] <- 1e-8
    c_star <- data$c_star
    A_star <- data$A_star
    n <- length(q0)

    # Limit k to n-1
    k_values <- k_values[k_values < n]

    cat(sprintf("  Sectors: %d, k values: %s\n", n, paste(k_values, collapse = ", ")))

    # Baseline
    base_model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
        days_in_year = days_in_year
    )
    base_loss <- base_model$total_economic_loss
    cat(sprintf("  Baseline loss: %.2f\n\n", base_loss))

    # ─── Step 1: Compute all rankings ───
    cat("  Computing PCA rankings...\n")
    pca_variants <- list()
    pc_counts <- c(1, 2, 3, 5)
    pc_counts <- pc_counts[pc_counts <= n]

    for (npc in pc_counts) {
        pca_res <- pca_rank_sectors(A, n_pcs = npc)
        label <- if (npc == 1) "PCA (PC1 only)" else sprintf("PCA (%d-PC distance)", npc)
        pca_variants[[label]] <- pca_res
        cat(sprintf(
            "    %s: cumulative variance = %.1f%%, top-5 = [%s]\n",
            label, pca_res$cumulative_var * 100,
            paste(pca_res$ranked_sectors[1:min(5, n)], collapse = ", ")
        ))
    }

    cat("  Computing DIIM ranking...\n")
    diim_res <- diim_rank_sectors(
        q0, A_star, c_star, x,
        lockdown_duration, total_duration, days_in_year
    )
    cat(sprintf(
        "    DIIM top-5: [%s]\n",
        paste(diim_res$ranked_sectors[1:5], collapse = ", ")
    ))

    cat("  Computing Sensitivity ranking...\n")
    sens_res <- sensitivity_rank_sectors(
        q0, A_star, c_star, x,
        lockdown_duration, total_duration, days_in_year
    )
    cat(sprintf(
        "    Sensitivity top-5: [%s]\n",
        paste(sens_res$ranked_sectors[1:5], collapse = ", ")
    ))

    cat("  Computing Network Centrality ranking...\n")
    cent_res <- centrality_rank_sectors(A_star)
    cat(sprintf(
        "    Centrality top-5: [%s]\n",
        paste(cent_res$ranked_sectors[1:5], collapse = ", ")
    ))

    # ─── Step 2: Evaluate all methods at all k values ───
    cat("\n  Evaluating all methods at each k...\n")

    all_methods <- list()
    for (label in names(pca_variants)) {
        all_methods[[label]] <- pca_variants[[label]]$ranked_sectors
    }
    all_methods[["DIIM"]] <- diim_res$ranked_sectors
    all_methods[["Sensitivity"]] <- sens_res$ranked_sectors
    all_methods[["Network Centrality"]] <- cent_res$ranked_sectors

    results <- data.frame(
        scenario = character(), method = character(),
        k = integer(), pct_reduction = numeric(),
        sectors = character(), stringsAsFactors = FALSE
    )

    for (mname in names(all_methods)) {
        for (k in k_values) {
            ev <- evaluate_topk(
                all_methods[[mname]], k, q0, A_star, c_star, x,
                lockdown_duration, total_duration, base_loss, days_in_year
            )
            results <- rbind(results, data.frame(
                scenario = scenario_name, method = mname, k = k,
                pct_reduction = ev$pct_reduction,
                sectors = paste(ev$sectors, collapse = ","),
                stringsAsFactors = FALSE
            ))
        }
    }

    # Print summary table
    cat("\n  Results Table (% loss reduction):\n")
    summary_table <- reshape2::dcast(results, method ~ k, value.var = "pct_reduction")
    print(summary_table)

    return(results)
}


###############################################################################
# Run both scenarios
###############################################################################
cat("=== Enhanced Comparison: Multi-PC Distance + Variable Top-K ===\n")

# COVID-19 (15 sectors): k up to 12
covid_results <- run_scenario(
    "COVID-19", download_data,
    lockdown_duration = 55, total_duration = 751,
    days_in_year = 366,
    k_values = c(3, 5, 7, 10, 12)
)

# Manpower (107 sectors): k up to 30
manpower_results <- run_scenario(
    "Manpower", download_manpower_data,
    lockdown_duration = 55, total_duration = 751,
    days_in_year = 365,
    k_values = c(3, 5, 7, 10, 15, 20, 30)
)

# Combine and save
combined <- rbind(covid_results, manpower_results)
write.csv(combined, file.path(results_dir, "enhanced_comparison.csv"),
    row.names = FALSE
)
cat("\n  -> Saved enhanced_comparison.csv\n")


###############################################################################
# PLOTS
###############################################################################
cat("\n--- Generating Plots ---\n")

# Helper: generate a line chart for a given scenario
plot_topk_comparison <- function(data, scenario_name, filename) {
    plot_data <- data[data$scenario == scenario_name, ]

    # Categorize PCA variants
    plot_data$method_type <- ifelse(grepl("PCA", plot_data$method), "PCA Variant",
        ifelse(plot_data$method == "DIIM", "DIIM",
            ifelse(plot_data$method == "Sensitivity", "Sensitivity",
                "Other"
            )
        )
    )

    p <- ggplot(plot_data, aes(
        x = k, y = pct_reduction,
        color = method, linetype = method_type
    )) +
        geom_line(linewidth = 1.1) +
        geom_point(size = 2.5) +
        scale_color_brewer(palette = "Set2", name = "Method") +
        scale_linetype_manual(
            values = c(
                "PCA Variant" = "solid",
                "DIIM" = "dashed",
                "Sensitivity" = "dotdash",
                "Other" = "dotted"
            ),
            name = "Type"
        ) +
        labs(
            title = sprintf(
                "%s: Loss Reduction vs Number of Sectors Intervened (k)",
                scenario_name
            ),
            subtitle = "Comparing PCA variants, DIIM, Sensitivity, and Network Centrality",
            x = "Number of Sectors Intervened (k)",
            y = "% Loss Reduction"
        ) +
        theme_minimal() +
        theme(
            plot.title = element_text(face = "bold", size = 14),
            legend.position = "bottom",
            legend.box = "vertical"
        ) +
        guides(
            color = guide_legend(nrow = 3),
            linetype = guide_legend(nrow = 1)
        )

    ggsave(file.path(results_dir, filename), p, width = 12, height = 7)
    cat(sprintf("  -> Saved %s\n", filename))
}

# Plot for each scenario
plot_topk_comparison(combined, "COVID-19", "enhanced_covid_topk.png")
plot_topk_comparison(combined, "Manpower", "enhanced_manpower_topk.png")


# Plot: PCA variant comparison (which PC distance works best?)
pca_data <- combined[grepl("PCA", combined$method), ]
if (nrow(pca_data) > 0) {
    p_pca <- ggplot(pca_data, aes(
        x = k, y = pct_reduction,
        color = method, shape = method
    )) +
        geom_line(linewidth = 1.2) +
        geom_point(size = 3) +
        facet_wrap(~scenario, scales = "free_y") +
        scale_color_brewer(palette = "Dark2", name = "PCA Variant") +
        labs(
            title = "PCA Variant Comparison: PC1 vs Multi-PC Distance",
            x = "Top-k Sectors", y = "% Loss Reduction"
        ) +
        theme_minimal() +
        theme(
            plot.title = element_text(face = "bold", size = 14),
            legend.position = "bottom"
        )

    ggsave(file.path(results_dir, "enhanced_pca_variants.png"), p_pca,
        width = 12, height = 6
    )
    cat("  -> Saved enhanced_pca_variants.png\n")
}


# Plot: Winner at each k (bar chart)
winners <- combined %>%
    group_by(scenario, k) %>%
    filter(pct_reduction == max(pct_reduction)) %>%
    ungroup()

p_winner <- ggplot(winners, aes(
    x = factor(k), y = pct_reduction,
    fill = method
)) +
    geom_bar(stat = "identity", position = "dodge", alpha = 0.85) +
    facet_wrap(~scenario, scales = "free_y") +
    scale_fill_brewer(palette = "Set2", name = "Best Method") +
    labs(
        title = "Best Method at Each k",
        x = "Number of Sectors (k)", y = "% Loss Reduction"
    ) +
    theme_minimal() +
    theme(
        plot.title = element_text(face = "bold", size = 14),
        legend.position = "bottom"
    )

ggsave(file.path(results_dir, "enhanced_best_at_k.png"), p_winner,
    width = 12, height = 6
)
cat("  -> Saved enhanced_best_at_k.png\n")


# ─────────────────────────────────────────────────────────────────────────────
# Summary
# ─────────────────────────────────────────────────────────────────────────────
cat("\n══════════════════════════════════════════════════════\n")
cat("  ENHANCED COMPARISON SUMMARY\n")
cat("══════════════════════════════════════════════════════\n\n")

for (sc in unique(combined$scenario)) {
    cat(sprintf("--- %s ---\n", sc))
    sc_data <- combined[combined$scenario == sc, ]

    for (k_val in sort(unique(sc_data$k))) {
        k_data <- sc_data[sc_data$k == k_val, ]
        best <- k_data[which.max(k_data$pct_reduction), ]
        cat(sprintf("  k=%2d: Best = %-25s (%.2f%%)", k_val, best$method, best$pct_reduction))

        # Also show PCA variants for comparison
        pca_vals <- k_data[grepl("PCA", k_data$method), ]
        if (nrow(pca_vals) > 0) {
            best_pca <- pca_vals[which.max(pca_vals$pct_reduction), ]
            cat(sprintf("  |  Best PCA = %-20s (%.2f%%)", best_pca$method, best_pca$pct_reduction))
        }
        cat("\n")
    }
    cat("\n")
}

# PCA improvement from multi-PC
cat("PCA Improvement from Multi-PC Distance (at k=5):\n")
for (sc in unique(combined$scenario)) {
    sc_data <- combined[combined$scenario == sc & combined$k == 5, ]
    pc1_only <- sc_data[sc_data$method == "PCA (PC1 only)", "pct_reduction"]
    multi_pc <- sc_data[grepl("2-PC|3-PC|5-PC", sc_data$method), ]
    if (length(pc1_only) > 0 && nrow(multi_pc) > 0) {
        best_multi <- multi_pc[which.max(multi_pc$pct_reduction), ]
        improvement <- best_multi$pct_reduction - pc1_only
        cat(sprintf(
            "  %s: PC1=%.3f%% -> %s=%.3f%% (delta=%.3f%%)\n",
            sc, pc1_only, best_multi$method, best_multi$pct_reduction,
            improvement
        ))
    }
}

cat("\n══════════════════════════════════════════════════════\n")
cat("Done!\n")
