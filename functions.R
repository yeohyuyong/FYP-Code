library(openxlsx)

download_data <- function() {
  # download 2019 IOT and extract total output and final demand column
  data <- read.xlsx("dataset/io_tables/iot_2019_2021_combined.xlsx", sheet = "2019IOTable")
  iot2019 <- data[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
  x <- data[2:(2 + 15 - 1), ncol(data)]
  c <- data[2:(2 + 15 - 1), ncol(data) - 1]
  x <- as.matrix(sapply(x, as.numeric))
  c <- as.matrix(sapply(c, as.numeric))

  # extract technical coefficient matrix
  A <- read.xlsx("dataset/io_tables/iot_2019_2021_combined.xlsx", sheet = "A")
  A <- A[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
  A <- as.matrix(sapply(A, as.numeric))

  # extract sector initial inoperability
  q0_df <- read.xlsx("dataset/covid_data/unemployment_and_impact_analysis.xlsx", sheet = "sector inoperability", colNames = FALSE)
  q0 <- as.numeric(q0_df[2:16, 3])

  c_star_df <- read.xlsx("dataset/covid_data/unemployment_and_impact_analysis.xlsx", sheet = "c_star")
  c_star <- as.numeric(c_star_df[1:15, 7])

  A_star <- solve(diag(as.vector(x))) %*% A %*% diag(as.vector(x))

  return(list(
    x = x,
    c = c,
    A = A,
    q0 = q0,
    c_star = c_star,
    A_star = A_star
  ))
}

download_manpower_data <- function() {
  # Load 15-sector aggregated 2022 IOT for A matrix and x vector
  data_A <- read.xlsx("dataset/manpower_disruption_data/iot_2022_15_sectors.xlsx", sheet = "A", colNames = FALSE)
  num_sectors <- nrow(data_A) - 1
  A <- data_A[2:(num_sectors + 1), 2:(num_sectors + 1)]
  A <- as.matrix(sapply(A, as.numeric))

  # Sheet "x" contains the total output
  data_x <- read.xlsx("dataset/manpower_disruption_data/iot_2022_15_sectors.xlsx", sheet = "x")
  x <- as.numeric(data_x$Total_Output)

  # Load sector initial inoperability (q0)
  data_q0 <- read.xlsx("dataset/manpower_disruption_data/sector_initial_inoperability_15_sectors.xlsx", sheet = "Sector_initial_inoperability")
  q0 <- as.numeric(data_q0$Sector_initial_inoperability)

  # Initialize c and c_star to 0 for manpower disruption scenario
  c <- rep(0, num_sectors)
  c_star <- rep(0, num_sectors)

  # Calculate A_star
  A_star <- solve(diag(x)) %*% A %*% diag(x)

  return(list(
    x = as.matrix(x),
    c = as.matrix(c),
    A = A,
    q0 = q0,
    c_star = c_star,
    A_star = A_star
  ))
}

DIIM <- function(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = NULL, risk_management = 1, days_in_year = 366) {
  a_ii <- diag(A_star)
  T <- total_duration

  num_sectors <- length(q0)
  inoperability_evolution <- matrix(NA, nrow = num_sectors, ncol = total_duration)

  # if key_sectors is not NULL, reduce those sectors' initial q0 by 10%
  if (!is.null(key_sectors)) {
    q0[key_sectors] <- q0[key_sectors] * 0.9
  }

  # we assume after 2 years the economic activity return to 99% of pre lock down level
  q0[q0 == 0] <- 1e-8
  qT = q0 * 1/100

  k <- log(q0 / qT) / (T * (1 - a_ii))
  K <- diag(as.vector(k))

  inoperability_evolution[, 1] <- q0

  for (t in 2:total_duration) {
    if (t <= lockdown_duration) {
      inoperability_evolution[, t] <- inoperability_evolution[, t - 1] + K %*% (A_star %*% inoperability_evolution[, t - 1] + c_star - inoperability_evolution[, t - 1]) * risk_management
    } else {
      # after lockdown c* become 0 as no more external shock
      inoperability_evolution[, t] <- inoperability_evolution[, t - 1] + K %*% (A_star %*% inoperability_evolution[, t - 1] - inoperability_evolution[, t - 1]) * risk_management
    }
  }

  # convert to planned DAILY output (output is in millions)
  x_daily <- x / days_in_year
  EL_evolution <- t(apply(inoperability_evolution, 1, cumsum)) * as.vector(x_daily)
  EL_end <- EL_evolution[, ncol(EL_evolution)]
  total_economic_loss <- sum(EL_evolution[, ncol(EL_evolution)])

  return(
    list(
      inoperability_evolution = inoperability_evolution,
      EL_evolution = EL_evolution,
      EL_end = EL_end,
      total_economic_loss = total_economic_loss
    )
  )
}

simulation_ml_vs_diim <- function(q0, A_star, c_star, x, lockdown_duration, total_duration) {
  # run model to calculate the economic loss (without intervention)
  model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration)
  EL_evolution <- model$EL_evolution

  # obtain the sectors with top economic loss
  max_econ_loss <- apply(EL_evolution, 1, max)
  sorted_indices <- order(max_econ_loss, decreasing = TRUE)
  top_econ_loss_5 <- sorted_indices[1:5]

  # rerun the model with intervention for the top economic sectors
  model_diim <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = top_econ_loss_5)

  # Compute PCA x xi key sectors dynamically
  I_minus_A <- diag(nrow(A_star)) - A_star  # Use A_star since A is not passed
  L_local <- solve(I_minus_A)
  H_local <- A_star %*% L_local
  eig_local <- eigen(H_local)
  pc_loadings_local <- Re(eig_local$vectors[, 1:min(2, nrow(A_star))])
  pca_dist_local <- sqrt(rowSums(pc_loadings_local^2))
  ml_key_sectors <- order(pca_dist_local * as.vector(x), decreasing = TRUE)[1:5]
  model_ml <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = ml_key_sectors)

  model_tot_econ_loss <- model$total_economic_loss
  model_diim_tot_econ_loss <- model_diim$total_economic_loss
  model_ml_tot_econ_loss <- model_ml$total_economic_loss

  return(list(
    lockdown_duration = lockdown_duration,
    total_duration = total_duration,
    model_tot_econ_loss = model_tot_econ_loss,
    model_diim_tot_econ_loss = model_diim_tot_econ_loss,
    model_ml_tot_econ_loss = model_ml_tot_econ_loss
  ))
}

simulation_ml_vs_diim_manpower <- function(q0, A_star, c_star, x, lockdown_duration, total_duration) {
  # run model to calculate the economic loss (without intervention)
  model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration)
  EL_evolution <- model$EL_evolution

  # obtain the sectors with top economic loss
  max_econ_loss <- apply(EL_evolution, 1, max)
  sorted_indices <- order(max_econ_loss, decreasing = TRUE)
  top_econ_loss_5 <- sorted_indices[1:5]

  # rerun the model with intervention for the top economic sectors
  model_diim <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = top_econ_loss_5)

  # Compute PCA x xi key sectors dynamically
  I_minus_A <- diag(nrow(A_star)) - A_star
  L_local <- solve(I_minus_A)
  H_local <- A_star %*% L_local
  eig_local <- eigen(H_local)
  pc_loadings_local <- Re(eig_local$vectors[, 1:min(2, nrow(A_star))])
  pca_dist_local <- sqrt(rowSums(pc_loadings_local^2))
  ml_key_sectors <- order(pca_dist_local * as.vector(x), decreasing = TRUE)[1:5]
  model_ml <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = ml_key_sectors)

  model_tot_econ_loss <- model$total_economic_loss
  model_diim_tot_econ_loss <- model_diim$total_economic_loss
  model_ml_tot_econ_loss <- model_ml$total_economic_loss

  return(list(
    lockdown_duration = lockdown_duration,
    total_duration = total_duration,
    model_tot_econ_loss = model_tot_econ_loss,
    model_diim_tot_econ_loss = model_diim_tot_econ_loss,
    model_ml_tot_econ_loss = model_ml_tot_econ_loss
  ))
}



# gini: computes the Gini coefficient of a vector (measures inequality)
# Returns 0 for perfectly equal, closer to 1 for highly unequal
gini <- function(x) {
  x <- sort(abs(x))
  n <- length(x)
  if (sum(x) == 0) return(0)
  index <- 1:n
  return((2 * sum(index * x) / (n * sum(x))) - (n + 1) / n)
}

# compare_methods: runs DIIM three times (baseline, DIIM top-k intervention, PCA intervention)
# and returns losses, reductions, and whether PCA wins
compare_methods <- function(q0, A_star, c_star, x,
                            lockdown_duration, total_duration,
                            pca_sectors,
                            days_in_year = 366) {
  # 1. Baseline run (no intervention)
  model_base <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, days_in_year = days_in_year)
  EL_base <- model_base$EL_evolution

  # 2. Find DIIM top-5 sectors (by peak economic loss)
  max_econ_loss <- apply(EL_base, 1, max)
  diim_sectors <- order(max_econ_loss, decreasing = TRUE)[1:length(pca_sectors)]

  # 3. Rerun with DIIM intervention
  model_diim <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
                     key_sectors = diim_sectors, days_in_year = days_in_year)

  # 4. Rerun with PCA intervention
  model_pca <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
                    key_sectors = pca_sectors, days_in_year = days_in_year)

  loss_base <- model_base$total_economic_loss
  loss_diim <- model_diim$total_economic_loss
  loss_pca  <- model_pca$total_economic_loss

  reduction_diim <- loss_base - loss_diim
  reduction_pca  <- loss_base - loss_pca
  pca_wins <- reduction_pca > reduction_diim
  overlap <- length(intersect(pca_sectors, diim_sectors)) # number of overlapping sectors

  return(list(
    loss_base = loss_base, loss_diim = loss_diim, loss_pca = loss_pca,
    reduction_diim = reduction_diim, reduction_pca = reduction_pca,
    pca_wins = pca_wins, diim_sectors = paste(diim_sectors, collapse = ","),
    overlap = overlap
  ))
}

# pca_rank_sectors: ranks sectors by Euclidean distance in multi-PC space, weighted by total output
# Uses eigendecomposition of the influence matrix H = A * L
# Sectors farther from origin in PC space AND with higher output are ranked more important
pca_rank_sectors <- function(A_matrix, x_vector, n_pcs = 2) {
  n <- nrow(A_matrix)
  I_minus_A <- diag(n) - A_matrix
  L <- solve(I_minus_A)          # Leontief inverse
  H <- A_matrix %*% L            # influence matrix

  eig <- eigen(H) # computes both eigenvalue and eigenvector
  eigenvalues <- Re(eig$values)
  eigenvectors <- Re(eig$vectors)
  n_pcs <- min(n_pcs, ncol(eigenvectors))

  # loadings on the first n_pcs principal components
  loadings <- eigenvectors[, 1:n_pcs, drop = FALSE]

  # Euclidean distance from origin
  distances <- sqrt(rowSums(loadings^2))

  # weight by total output (xi)
  weighted_distances <- distances * as.vector(x_vector)
  ranked_sectors <- order(weighted_distances, decreasing = TRUE)

  # variance explained by each PC
  var_explained <- eigenvalues / sum(abs(eigenvalues))
  cumulative_var <- cumsum(var_explained[1:n_pcs])

  return(list(
    ranked_sectors = ranked_sectors,
    distances = distances,
    weighted_distances = weighted_distances,
    loadings = loadings,
    var_explained = var_explained[1:n_pcs],
    cumulative_var = tail(cumulative_var, 1),
    n_pcs = n_pcs
  ))
}

# pagerank_rank_sectors: ranks sectors by PageRank weighted by total output
pagerank_rank_sectors <- function(A_star, x_vector) {
  g <- igraph::graph_from_adjacency_matrix(A_star, mode = "directed",
                                           weighted = TRUE, diag = FALSE)

  pr <- igraph::page_rank(g)$vector

  # weight by total output (xi)
  weighted_pr <- pr * as.vector(x_vector)

  ranked_sectors <- order(weighted_pr, decreasing = TRUE)
  return(list(ranked_sectors = ranked_sectors, pagerank = pr, weighted_pagerank = weighted_pr))
}

# compute_cheap_rankings: computes 5 "cheap" sector rankings that don't need DIIM
# All structural methods are weighted by total output (xi) to reflect economic size
# 1. Total Output  2. PCA x xi  3. PageRank x xi  4. BL x xi  5. FL x xi
compute_cheap_rankings <- function(A, A_star, x) {
  n <- nrow(A)
  rankings <- list()

  # 1. Total Output: rank sectors by total output (xi)
  rankings[["Total Output"]] <- order(as.vector(x), decreasing = TRUE)

  # Compute Leontief inverse (shared by PCA, BL, FL)
  I_minus_A <- diag(n) - A
  L <- solve(I_minus_A)

  # 2. PCA x xi: 2-PC distance on influence matrix H = A * L, weighted by xi
  H <- A %*% L
  eig <- eigen(H)
  pc_loadings <- Re(eig$vectors[, 1:min(2, n)])
  pca_dist <- sqrt(rowSums(pc_loadings^2))
  rankings[["PCA x xi"]] <- order(pca_dist * as.vector(x), decreasing = TRUE)

  # 3. PageRank x xi: PageRank centrality weighted by total output
  pr_results <- pagerank_rank_sectors(A_star, x)
  rankings[["PageRank x xi"]] <- pr_results$ranked_sectors

  # 4. BL x xi: Backward Linkage (column sums of L) weighted by total output
  BL <- colSums(L)
  rankings[["BL x xi"]] <- order(BL * as.vector(x), decreasing = TRUE)

  # 5. FL x xi: Forward Linkage (row sums of L) weighted by total output
  FL <- rowSums(L)
  rankings[["FL x xi"]] <- order(FL * as.vector(x), decreasing = TRUE)

  return(rankings)
}
