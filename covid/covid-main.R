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


# Calculate maximum inoperability for each sector over time
inoperability_evolution <- DIIM_model$inoperability_evolution
max_inoperability <- apply(inoperability_evolution, 1, max)

# Get the order of sectors sorted by descending maximum inoperability
sorted_indices <- order(max_inoperability, decreasing = TRUE)

# Reorder the inoperability evolution matrix
sorted_inoperability <- inoperability_evolution[sorted_indices, ]

# Optional: get sector names if available (replace 'sector_names')
# sector_names_sorted <- sector_names[sorted_indices]

# Setup colors and labels for plotting
num_sectors <- nrow(sorted_inoperability)
# Save Inoperability Evolution plot
png("covid/Inoperability_Evolution.png", units = "in", width = 10, height = 7, res = 300)
matplot(t(sorted_inoperability),
        type = "l", lty = 1, col = rainbow(num_sectors),
        xlab = "Days", ylab = "Inoperability",
        main = "Inoperability Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices), col = rainbow(num_sectors), lty = 1, cex = 0.6)
dev.off()


EL_evolution <- DIIM_model$EL_evolution
max_econ_loss <- apply(EL_evolution, 1, max)
sorted_indices <- order(max_econ_loss, decreasing = TRUE)
sorted_econ_loss <- EL_evolution[sorted_indices, ]


num_sectors <- nrow(sorted_econ_loss)
# Save Economic Loss Evolution plot
png("covid/Economic_Loss_Evolution.png", units = "in", width = 10, height = 7, res = 300)
matplot(t(sorted_econ_loss),
        type = "l", lty = 1, col = rainbow(num_sectors),
        xlab = "Days", ylab = "Economic Loss",
        main = "Economic Loss Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices), col = rainbow(num_sectors), lty = 1, cex = 0.6)
dev.off()


output_40_days <- DIIM(q0, A_star, c_star, x, lockdown_duration = 40, total_duration = 751)
economics_loss_40_days <- as.matrix(output_40_days$EL_end)

output_55_days <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751)
economics_loss_55_days <- as.matrix(output_55_days$EL_end)

output_70_days <- DIIM(q0, A_star, c_star, x, lockdown_duration = 70, total_duration = 751)
economics_loss_70_days <- as.matrix(output_70_days$EL_end)



# evolution of the inoperability under policy 1 where the inoperability is 95% of the original inoperability
policy_j_0 <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751, risk_management = 1)
policy_j_1 <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751, risk_management = 0.95)

policy_j_0$total_economic_loss
policy_j_1$total_economic_loss


policy_j_3 <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55 - 20, total_duration = 751, risk_management = 1)
policy_j_3$total_economic_loss



# ml_key_sectors = c(3,2,5,8,10,4,7)
# trad_key_sectors = c(2,3,5,12,8,9,10)

lockdown_duration_vals <- c(10, 20, 30, 40)
total_duration_vals <- c(300, 400, 500, 600)

nsim <- length(lockdown_duration_vals) * length(total_duration_vals)

col_names <- c(
        "lockdown_duration",
        "total_duration",
        "model_tot_econ_loss",
        "model_diim_tot_econ_loss",
        "model_ml_tot_econ_loss"
)

sim_matrix <- matrix(data = NA, nrow = nsim, ncol = 5)
colnames(sim_matrix) <- col_names

row_idx <- 1
