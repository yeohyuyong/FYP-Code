library(openxlsx)
library(ggplot2)
library(igraph)
library(factoextra)
library(FactoMineR)
library(dendextend)
library(tidyverse)
library("corrr")
library(ggcorrplot)

data <- read.xlsx("dataset/Input_Output_Tables_2019_2021.xlsx", sheet = "2019IOTable")
iot2019 <- data[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
x <- data[2:(2 + 15 - 1), ncol(data)]

A <- read.xlsx("dataset/Input_Output_Tables_2019_2021.xlsx", sheet = "A")
A <- A[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
A <- as.matrix(sapply(A, as.numeric))
n_sectors <- nrow(A)

I_minus_A <- diag(nrow(A)) - A
L <- solve(I_minus_A)

calculate_linkages <- function(L_matrix, x_vector = NULL) {
  BL <- colSums(L_matrix)
  BL_normalized <- BL / mean(BL)
  FL <- rowSums(L_matrix)
  FL_normalized <- FL / mean(FL)
  total_linkages <- (BL_normalized + FL_normalized) / 2
  return(data.frame(
    Sector_ID = 1:length(BL),
    BL = BL, BL_normalized = BL_normalized,
    FL = FL, FL_normalized = FL_normalized,
    Total_Linkage = total_linkages
  ))
}

linkages <- calculate_linkages(L)
cat("\nRasmussen-Hirschman Linkages:\n")
print(linkages %>% arrange(desc(Total_Linkage)))

construct_H_matrix <- function(A) {
  I_minus_A <- diag(nrow(A)) - A
  L <- solve(I_minus_A)
  H <- A %*% L
  return(H)
}

H <- construct_H_matrix(A)

perform_pca <- function(H_matrix) {
  eigen_result <- eigen(H_matrix)
  eigenvalues <- eigen_result$values
  eigenvectors <- eigen_result$vectors
  variance_explained <- eigenvalues / sum(eigenvalues)
  cumulative_variance <- cumsum(variance_explained)
  return(list(
    eigenvalues = Re(eigenvalues),
    eigenvectors = Re(eigenvectors),
    variance_explained = Re(variance_explained),
    cumulative_variance = Re(cumulative_variance)
  ))
}

pca_result <- perform_pca(H)

PC1 <- pca_result$eigenvectors[, 1]
PC2 <- pca_result$eigenvectors[, 2]

pca_data <- data.frame(
  Sector_ID = 1:nrow(A),
  PC1 = PC1, PC2 = PC2,
  Eigenvalue_1 = pca_result$eigenvalues[1],
  Eigenvalue_2 = pca_result$eigenvalues[2]
)
# Compute Euclidean distance from origin in PC1-PC2 space
pca_data$Distance <- sqrt(pca_data$PC1^2 + pca_data$PC2^2)

# Weight by total output (xi) for xi-weighted ranking
pca_data$Weighted_Distance <- pca_data$Distance * as.vector(x)

# Rank sectors by xi-weighted distance (descending)
pca_data <- pca_data[order(pca_data$Weighted_Distance, decreasing = TRUE), ]

cat("\nPCA x xi Key Sector Ranking:\n")
cat("Top 5 sectors:", pca_data$Sector_ID[1:5], "\n\n")
print(pca_data[, c("Sector_ID", "PC1", "PC2", "Distance", "Weighted_Distance")])


plot_scree <- function(pca_result, sector_names = NULL) {
  n_components <- min(10, length(pca_result$variance_explained))
  scree_data <- data.frame(
    Component = 1:n_components,
    Variance_Explained = pca_result$variance_explained[1:n_components] * 100,
    Cumulative = pca_result$cumulative_variance[1:n_components] * 100
  )
  p <- ggplot(scree_data, aes(x = Component)) +
    geom_line(aes(y = Variance_Explained, color = "Individual"), size = 1) +
    geom_point(aes(y = Variance_Explained, color = "Individual"), size = 3) +
    geom_line(aes(y = Cumulative, color = "Cumulative"), size = 1) +
    geom_point(aes(y = Cumulative, color = "Cumulative"), size = 3) +
    geom_hline(yintercept = 70, linetype = "dashed", color = "gray", size = 0.5) +
    geom_hline(yintercept = 80, linetype = "dashed", color = "gray", size = 0.5) +
    labs(title = "Scree Plot: Variance Explained by Principal Components",
         x = "Principal Component", y = "Variance Explained (%)", color = "Type") +
    theme_minimal() + theme(legend.position = "right")
  print(p)
  return(scree_data)
}

scree_data <- plot_scree(pca_result)

cat("\nVariance Explained by Principal Components:\n")
print(data.frame(
  PC = 1:5,
  Variance = pca_result$variance_explained[1:5] * 100,
  Cumulative = pca_result$cumulative_variance[1:5] * 100
))

plot_factor_map <- function(pca_data, sector_labels = NULL) {
  if (is.null(sector_labels)) {
    sector_labels <- paste("S", pca_data$Sector_ID, sep = "")
  }
  p <- ggplot(pca_data, aes(x = PC1, y = PC2)) +
    geom_point(size = 4, color = "steelblue", alpha = 0.6) +
    geom_text(aes(label = Sector_ID), size = 3, hjust = -0.3) +
    geom_vline(xintercept = 0, linetype = "dashed", color = "gray", alpha = 0.5) +
    geom_hline(yintercept = 0, linetype = "dashed", color = "gray", alpha = 0.5) +
    labs(title = "Factor Map: First Two Principal Components",
         x = paste("PC1 (", round(pca_result$variance_explained[1] * 100, 1), "% variance)"),
         y = paste("PC2 (", round(pca_result$variance_explained[2] * 100, 1), "% variance)")) +
    theme_minimal() + coord_equal()
  print(p)
  return(p)
}

factor_map <- plot_factor_map(pca_data)

pca_matrix <- pca_data[, 2:3]
str(pca_matrix)

set.seed(100)

wcss <- function(k) {
  kmeans(pca_matrix, k, nstart = 10)$tot.withinss
}

k.values <- 1:10
wcss_k <- sapply(k.values, wcss)
plot(k.values, wcss_k,
  type = "b", pch = 19, frame = FALSE,
  xlab = "Number of clusters K",
  ylab = "Total within-clusters sum of squares"
)

library(ggplot2)
elbow <- data.frame(k.values, wcss_k)
ggplot(elbow, aes(x = k.values, y = wcss_k)) +
  geom_point() +
  geom_line() +
  scale_x_continuous(breaks = seq(1, 30, by = 1))

# Select the elbow point k=3
set.seed(100)
k3 <- kmeans(pca_matrix, centers = 3, nstart = 10)
str(k3)
k3

hc1 <- hclust(dist(pca_matrix), method = "complete")
plot(hc1, cex = 0.6, hang = -.1)
rect.hclust(hc1, k = 3, border = 2:8)
