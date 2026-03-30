# Set working directory to project root
if (!file.exists("functions.R")) setwd("..")

library(openxlsx)
library(ggplot2)
library(reshape2)
library(gridExtra)
library(dplyr)
library(tidyverse)

source("functions.R")

data <- download_data()
A <- data$A
x <- data$x
c <- data$c
q0 <- data$q0
c_star <- data$c_star
A_star <- data$A_star

DIIM_model <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751)

inoperability_evolution <- DIIM_model$inoperability_evolution
max_inoperability <- apply(inoperability_evolution, 1, max)
sorted_indices <- order(max_inoperability, decreasing = TRUE)
sorted_inoperability <- inoperability_evolution[sorted_indices, ]

num_sectors <- nrow(sorted_inoperability)
png("covid/Inoperability_Evolution.png", units = "in", width = 10, height = 7, res = 300)
matplot(t(sorted_inoperability),
        type = "l", lty = 1, col = rainbow(num_sectors),
        xlab = "Days", ylab = "Inoperability",
        main = "Inoperability Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices), col = rainbow(num_sectors), lty = 1, cex = 0.6)
dev.off()

# Also display inline
matplot(t(sorted_inoperability),
        type = "l", lty = 1, col = rainbow(num_sectors),
        xlab = "Days", ylab = "Inoperability",
        main = "Inoperability Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices), col = rainbow(num_sectors), lty = 1, cex = 0.6)

EL_evolution <- DIIM_model$EL_evolution
max_econ_loss <- apply(EL_evolution, 1, max)
sorted_indices <- order(max_econ_loss, decreasing = TRUE)
sorted_econ_loss <- EL_evolution[sorted_indices, ]

num_sectors <- nrow(sorted_econ_loss)
png("covid/Economic_Loss_Evolution.png", units = "in", width = 10, height = 7, res = 300)
matplot(t(sorted_econ_loss),
        type = "l", lty = 1, col = rainbow(num_sectors),
        xlab = "Days", ylab = "Economic Loss",
        main = "Economic Loss Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices), col = rainbow(num_sectors), lty = 1, cex = 0.6)
dev.off()

# Also display inline
matplot(t(sorted_econ_loss),
        type = "l", lty = 1, col = rainbow(num_sectors),
        xlab = "Days", ylab = "Economic Loss",
        main = "Economic Loss Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices), col = rainbow(num_sectors), lty = 1, cex = 0.6)
