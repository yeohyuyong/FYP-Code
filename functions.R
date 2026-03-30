library(openxlsx)

download_data <- function() {
  # 2019 IOT: extract 15x15 intermediate flows, total output, and final demand
  data <- read.xlsx("dataset/io_tables/iot_2019_2021_combined.xlsx", sheet = "2019IOTable")
  iot2019 <- data[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
  x <- data[2:(2 + 15 - 1), ncol(data)]
  c <- data[2:(2 + 15 - 1), ncol(data) - 1]
  x <- as.matrix(sapply(x, as.numeric))
  c <- as.matrix(sapply(c, as.numeric))

  # Technical coefficient matrix A
  A <- read.xlsx("dataset/io_tables/iot_2019_2021_combined.xlsx", sheet = "A")
  A <- A[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
  A <- as.matrix(sapply(A, as.numeric))

  # Sector initial inoperability q(0) and perturbation c*
  q0_df <- read.xlsx("dataset/covid_data/unemployment_and_impact_analysis.xlsx", sheet = "sector inoperability", colNames = FALSE)
  q0 <- as.numeric(q0_df[2:16, 3])

  c_star_df <- read.xlsx("dataset/covid_data/unemployment_and_impact_analysis.xlsx", sheet = "c_star")
  c_star <- as.numeric(c_star_df[1:15, 7])

  # Interdependency matrix: A* = x^{-1} A x
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
  # 2022 IOT (15-sector aggregation) for manpower disruption scenario
  data_A <- read.xlsx("dataset/manpower_disruption_data/iot_2022_15_sectors.xlsx", sheet = "A", colNames = FALSE)
  num_sectors <- nrow(data_A) - 1
  A <- data_A[2:(num_sectors + 1), 2:(num_sectors + 1)]
  A <- as.matrix(sapply(A, as.numeric))

  data_x <- read.xlsx("dataset/manpower_disruption_data/iot_2022_15_sectors.xlsx", sheet = "x")
  x <- as.numeric(data_x$Total_Output)

  data_q0 <- read.xlsx("dataset/manpower_disruption_data/sector_initial_inoperability_15_sectors.xlsx", sheet = "Sector_initial_inoperability")
  q0 <- as.numeric(data_q0$Sector_initial_inoperability)

  # No external perturbation in manpower scenario (c* = 0)
  c <- rep(0, num_sectors)
  c_star <- rep(0, num_sectors)

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

DIIM <- function(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = NULL, risk_management = 1, days_in_year = 366, intervention_magnitude = 0.10) {
  a_ii <- diag(A_star)
  T <- total_duration

  num_sectors <- length(q0)
  inoperability_evolution <- matrix(NA, nrow = num_sectors, ncol = total_duration)

  if (!is.null(key_sectors)) {
    q0[key_sectors] <- q0[key_sectors] * (1 - intervention_magnitude)
  }

  # Assume 99% recovery by time T; avoid log(0) for zero-inoperability sectors
  q0[q0 == 0] <- 1e-8
  qT <- q0 * 1 / 100

  # Recovery rate from Santos & Haimes (2004)
  k <- log(q0 / qT) / (T * (1 - a_ii))
  K <- diag(as.vector(k))

  inoperability_evolution[, 1] <- q0

  for (t in 2:total_duration) {
    if (t <= lockdown_duration) {
      # Lockdown phase: c* active, risk_management dampens propagation
      inoperability_evolution[, t] <- inoperability_evolution[, t - 1] + K %*% (A_star %*% inoperability_evolution[, t - 1] + c_star - inoperability_evolution[, t - 1]) * risk_management
    } else {
      # Recovery phase: c* = 0, no dampening
      inoperability_evolution[, t] <- inoperability_evolution[, t - 1] + K %*% (A_star %*% inoperability_evolution[, t - 1] - inoperability_evolution[, t - 1])
    }
  }

  # Cumulative economic loss (daily output in $M)
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

simulation_ml_vs_diim <- function(q0, A, A_star, c_star, x, lockdown_duration, total_duration) {
  # Baseline (no intervention)
  model <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration)
  EL_evolution <- model$EL_evolution

  # DIIM top-5: sectors with highest peak economic loss
  max_econ_loss <- apply(EL_evolution, 1, max)
  sorted_indices <- order(max_econ_loss, decreasing = TRUE)
  top_econ_loss_5 <- sorted_indices[1:5]
  model_diim <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, key_sectors = top_econ_loss_5)

  # PCA x xi top-5
  pca_results_local <- pca_rank_sectors(A, x, n_pcs = 2)
  ml_key_sectors <- pca_results_local$ranked_sectors[1:5] # ml here refers to PCA
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


# Compare baseline vs DIIM top-k vs PCA top-k intervention
compare_methods <- function(q0, A_star, c_star, x, lockdown_duration, total_duration, pca_sectors, days_in_year = 366) {
  model_base <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration, days_in_year = days_in_year)
  EL_base <- model_base$EL_evolution

  # DIIM top-k: sectors with highest peak loss
  max_econ_loss <- apply(EL_base, 1, max)
  diim_sectors <- order(max_econ_loss, decreasing = TRUE)[1:length(pca_sectors)]

  model_diim <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
    key_sectors = diim_sectors, days_in_year = days_in_year
  )
  model_pca <- DIIM(q0, A_star, c_star, x, lockdown_duration, total_duration,
    key_sectors = pca_sectors, days_in_year = days_in_year
  )

  loss_base <- model_base$total_economic_loss
  loss_diim <- model_diim$total_economic_loss
  loss_pca <- model_pca$total_economic_loss

  reduction_diim <- loss_base - loss_diim
  reduction_pca <- loss_base - loss_pca
  pca_wins <- reduction_pca > reduction_diim
  overlap <- length(intersect(pca_sectors, diim_sectors))

  return(list(
    loss_base = loss_base, loss_diim = loss_diim, loss_pca = loss_pca,
    reduction_diim = reduction_diim, reduction_pca = reduction_pca,
    pca_wins = pca_wins, diim_sectors = paste(diim_sectors, collapse = ","),
    overlap = overlap
  ))
}

# Rank sectors by PCA distance on H = A*L, weighted by total output xi
pca_rank_sectors <- function(A_matrix, x_vector, n_pcs = 2) {
  n <- nrow(A_matrix)
  I_minus_A <- diag(n) - A_matrix
  L <- solve(I_minus_A) # Leontief inverse
  H <- A_matrix %*% L # influence matrix

  eig <- eigen(H)
  eigenvalues <- Re(eig$values)
  eigenvectors <- Re(eig$vectors)
  n_pcs <- min(n_pcs, ncol(eigenvectors))
  loadings <- eigenvectors[, 1:n_pcs, drop = FALSE]

  distances <- sqrt(rowSums(loadings^2))
  weighted_distances <- distances * as.vector(x_vector)
  ranked_sectors <- order(weighted_distances, decreasing = TRUE)

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

# Five xi-weighted structural rankings (no DIIM run needed)
compute_simplified_rankings <- function(A, A_star, x) {
  n <- nrow(A)
  rankings <- list()
  xi <- as.vector(x)

  rankings[["Total Output"]] <- order(xi, decreasing = TRUE)

  L <- solve(diag(n) - A) # Leontief inverse

  # PCA x xi: 2-PC distance on H = A*L
  H <- A %*% L
  eig <- eigen(H)
  pc_loadings <- Re(eig$vectors[, 1:min(2, n)])
  pca_dist <- sqrt(rowSums(pc_loadings^2))
  rankings[["PCA x xi"]] <- order(pca_dist * xi, decreasing = TRUE)

  pr <- igraph::page_rank(igraph::graph_from_adjacency_matrix(
    A_star,
    mode = "directed", weighted = TRUE, diag = FALSE
  ))$vector
  rankings[["PageRank x xi"]] <- order(pr * xi, decreasing = TRUE)

  BL <- colSums(L)
  rankings[["BL x xi"]] <- order(BL * xi, decreasing = TRUE)

  FL <- rowSums(L)
  rankings[["FL x xi"]] <- order(FL * xi, decreasing = TRUE)

  return(rankings)
}
