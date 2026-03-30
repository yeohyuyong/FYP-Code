# Set working directory to project root
if (!file.exists("functions.R")) setwd("..")

library(openxlsx)
library(ggplot2)
library(reshape2)
library(gridExtra)
library(dplyr)
library(tidyverse)

source("functions.R")

data <- download_manpower_data()
A <- data$A
x <- data$x
c <- data$c
q0 <- data$q0
q0[q0 == 0] <- 1e-8
c_star <- data$c_star
A_star <- data$A_star
a_ii <- diag(A_star)
qT <- q0/100
k <- log(q0 / qT) / (751 * (1 - a_ii))

DIIM_model <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751, days_in_year = 365)

inoperability_evolution <- DIIM_model$inoperability_evolution
max_inoperability <- apply(inoperability_evolution, 1, max)
sorted_indices_inop <- order(max_inoperability, decreasing = TRUE)
sorted_inoperability <- inoperability_evolution[sorted_indices_inop, ]

num_sectors <- nrow(sorted_inoperability)
png("manpower-disruption/Inoperability_Evolution.png", units = "in", width = 10, height = 7, res = 300)
matplot(t(sorted_inoperability),
    type = "l", lty = 1, col = rainbow(num_sectors),
    xlab = "Days", ylab = "Inoperability",
    main = "Manpower Disruption: Inoperability Evolution")
legend("topright", legend = paste("Sector", sorted_indices_inop[1:10]), col = rainbow(num_sectors)[1:10], lty = 1, cex = 0.6, title = "Top 10")
dev.off()

# Display inline
matplot(t(sorted_inoperability),
    type = "l", lty = 1, col = rainbow(num_sectors),
    xlab = "Days", ylab = "Inoperability",
    main = "Manpower Disruption: Inoperability Evolution")
legend("topright", legend = paste("Sector", sorted_indices_inop[1:10]), col = rainbow(num_sectors)[1:10], lty = 1, cex = 0.6, title = "Top 10")

EL_evolution <- DIIM_model$EL_evolution
max_econ_loss <- apply(EL_evolution, 1, max)
sorted_indices_el <- order(max_econ_loss, decreasing = TRUE)
sorted_econ_loss <- EL_evolution[sorted_indices_el, ]

png("manpower-disruption/Economic_Loss_Evolution.png", units = "in", width = 10, height = 7, res = 300)
matplot(t(sorted_econ_loss),
    type = "l", lty = 1, col = rainbow(num_sectors),
    xlab = "Days", ylab = "Economic Loss (Millions)",
    main = "Manpower Disruption: Economic Loss Evolution")
legend("topright", legend = paste("Sector", sorted_indices_el[1:10]), col = rainbow(num_sectors)[1:10], lty = 1, cex = 0.6, title = "Top 10")
dev.off()

# Display inline
matplot(t(sorted_econ_loss),
    type = "l", lty = 1, col = rainbow(num_sectors),
    xlab = "Days", ylab = "Economic Loss (Millions)",
    main = "Manpower Disruption: Economic Loss Evolution")
legend("topright", legend = paste("Sector", sorted_indices_el[1:10]), col = rainbow(num_sectors)[1:10], lty = 1, cex = 0.6, title = "Top 10")

