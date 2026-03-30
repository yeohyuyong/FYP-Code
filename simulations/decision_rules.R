if (!file.exists("functions.R")) setwd("..")

library(dplyr)

source("functions.R")

results_dir <- "simulations/results"
dir.create(results_dir, showWarnings = FALSE, recursive = TRUE)

# short column prefixes (letters only, no spaces/hyphens)
method_prefixes <- c(
    "Total Output"  = "TotalOutput",
    "PCA x xi"      = "PCAxi",
    "PageRank x xi" = "PageRankxi",
    "BL x xi"       = "BLxi",
    "FL x xi"       = "FLxi"
)

run_monte_carlo <- function(scenario_name, data_loader, lockdown_duration = 55, total_duration = 751, days_in_year = 366, n_mc = 2000, k = 5) {
    cat(sprintf("\n%s: %d Monte Carlo trials (k=%d)\n", scenario_name, n_mc, k))

    data <- data_loader()
    A <- data$A
    x <- data$x
    c_star <- data$c_star
    A_star <- data$A_star

    q0_base <- data$q0
    q0_base[q0_base == 0] <- 1e-8
    num_sectors <- length(q0_base)

    cat("  Computing simplified rankings...\n")
    simplified_rankings <- compute_simplified_rankings(A, A_star, x)
    for (method_name in names(simplified_rankings)) {
        cat(sprintf(
            "    %-25s top-%d: [%s]\n", method_name, k,
            paste(simplified_rankings[[method_name]][1:k], collapse = ", ")
        ))
    }

    cat("  Running Monte Carlo...\n")
    set.seed(42)
    all_results <- list()

    for (trial in 1:n_mc) {
        # perturb q0 with log-normal noise (preserves relative structure)
        noise <- exp(rnorm(num_sectors, mean = 0, sd = 0.5))
        random_q0 <- pmax(q0_base * noise, 1e-6)
        random_q0 <- pmin(random_q0, 1) # cap at 1 (100% inoperability)

        # baseline: no key sectors protected
        base_model <- DIIM(random_q0, A_star, c_star, x,
            lockdown_duration, total_duration,
            days_in_year = days_in_year
        )
        base_loss <- base_model$total_economic_loss

        if (base_loss < 1e-6) next # Skip degenerate trials

        # gold standard: DIIM's own top-k
        max_el <- apply(base_model$EL_evolution, 1, max)
        diim_topk <- order(max_el, decreasing = TRUE)[1:k]

        diim_model <- DIIM(random_q0, A_star, c_star, x,
            lockdown_duration, total_duration,
            key_sectors = diim_topk, days_in_year = days_in_year
        )

        diim_reduction <- base_loss - diim_model$total_economic_loss
        if (diim_reduction < 1e-10) next

        trial_row <- data.frame(
            trial = trial,
            scenario = scenario_name,
            base_loss = base_loss,
            diim_reduction = diim_reduction,
            diim_pct = diim_reduction / base_loss * 100,
            stringsAsFactors = FALSE
        )

        for (method_name in names(simplified_rankings)) {
            topk_sectors <- simplified_rankings[[method_name]][1:k]

            method_model <- DIIM(random_q0, A_star, c_star, x,
                lockdown_duration, total_duration,
                key_sectors = topk_sectors, days_in_year = days_in_year
            )

            method_reduction <- base_loss - method_model$total_economic_loss

            prefix <- method_prefixes[method_name]
            trial_row[[paste0(prefix, "_reduction")]] <- method_reduction
            trial_row[[paste0(prefix, "_ratio")]] <- method_reduction / diim_reduction
            trial_row[[paste0(prefix, "_wins")]] <- method_reduction >= diim_reduction
            trial_row[[paste0(prefix, "_close")]] <- method_reduction / diim_reduction >= 0.80
        }

        all_results[[length(all_results) + 1]] <- trial_row
    }

    monte_carlo_df <- bind_rows(all_results)
    cat(sprintf("  Valid trials: %d\n", nrow(monte_carlo_df)))
    return(monte_carlo_df)
}

k_values <- c(3, 5, 7, 10, 12)
all_mc_results <- list()

for (k_val in k_values) {
    cat(sprintf("\n--- k = %d ---\n", k_val))

    covid_mc <- run_monte_carlo("COVID-19", download_data,
        lockdown_duration = 55, total_duration = 751,
        days_in_year = 366, n_mc = 2000, k = k_val
    )
    covid_mc$k <- k_val
    all_mc_results[[paste0("covid_k", k_val)]] <- covid_mc

    manpower_mc <- run_monte_carlo("Manpower", download_manpower_data,
        lockdown_duration = 55, total_duration = 751,
        days_in_year = 365, n_mc = 2000, k = k_val
    )
    manpower_mc$k <- k_val
    all_mc_results[[paste0("manpower_k", k_val)]] <- manpower_mc
}

combined_mc <- bind_rows(all_mc_results)
write.csv(combined_mc, file.path(results_dir, "decision_rules_mc_data.csv"), row.names = FALSE)
cat("Saved decision_rules_mc_data.csv\n")
cat("Done!\n")
