library(openxlsx)
library(ggplot2)
library(reshape2)
library(gridExtra)
library(dplyr)
library(tidyverse)

# Assuming running from project root
if (file.exists("functions.R")) {
    source("functions.R")
} else {
    source("../functions.R") # Fallback if running from subdir
}

# 1. Load Data
data <- download_manpower_data()
A <- data$A
x <- data$x
c <- data$c
q0 <- data$q0
c_star <- data$c_star
A_star <- data$A_star

# Preprocess q0 to handle zeros (since DIIM uses log(q0))
# Replace 0 or very small values with a small epsilon
q0[q0 < 1e-6] <- 1e-6

# 2. Main Simulation
# Manpower disruption likely has different durations.
# Using 55 days composed to match covid-main.R for now, user can adjust.
# Days in year = 365 for 2022
DIIM_model <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751, days_in_year = 365)


# 3. Analyze Inoperability
inoperability_evolution <- DIIM_model$inoperability_evolution
max_inoperability <- apply(inoperability_evolution, 1, max)

# Sort sectors by max inoperability
sorted_indices_inop <- order(max_inoperability, decreasing = TRUE)
sorted_inoperability <- inoperability_evolution[sorted_indices_inop, ]

# Plot Inoperability
num_sectors <- nrow(sorted_inoperability)
# Plot top 10 for clarity, or all if preferred. matching covid-main which plots all?
# covid-main plots all: matplot(t(sorted_inoperability), ...)
matplot(t(sorted_inoperability),
    type = "l", lty = 1, col = rainbow(num_sectors),
    xlab = "Days", ylab = "Inoperability",
    main = "Manpower Disruption: Inoperability Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices_inop[1:10]), col = rainbow(num_sectors)[1:10], lty = 1, cex = 0.6, title = "Top 10")


# 4. Analyze Economic Loss
EL_evolution <- DIIM_model$EL_evolution
max_econ_loss <- apply(EL_evolution, 1, max) # Max cumulative loss is the end value usually

sorted_indices_el <- order(max_econ_loss, decreasing = TRUE)
sorted_econ_loss <- EL_evolution[sorted_indices_el, ]

# Plot Economic Loss
matplot(t(sorted_econ_loss),
    type = "l", lty = 1, col = rainbow(num_sectors),
    xlab = "Days", ylab = "Economic Loss (Millions)",
    main = "Manpower Disruption: Economic Loss Evolution"
)
legend("topright", legend = paste("Sector", sorted_indices_el[1:10]), col = rainbow(num_sectors)[1:10], lty = 1, cex = 0.6, title = "Top 10")

print("Top 5 Sectors by Economic Loss:")
print(sorted_indices_el[1:5])


# 5. Sensitivity Analysis / Policy Simulations (Modeled after covid-main.R)

# Different lengths
output_40_days <- DIIM(q0, A_star, c_star, x, lockdown_duration = 40, total_duration = 751, days_in_year = 365)
loss_40 <- sum(output_40_days$total_economic_loss)

output_55_days <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751, days_in_year = 365)
loss_55 <- sum(output_55_days$total_economic_loss)

output_70_days <- DIIM(q0, A_star, c_star, x, lockdown_duration = 70, total_duration = 751, days_in_year = 365)
loss_70 <- sum(output_70_days$total_economic_loss)

print(paste("Total Loss (40 days):", loss_40))
print(paste("Total Loss (55 days):", loss_55))
print(paste("Total Loss (70 days):", loss_70))


# Risk Management Policies
# Policy 1: Risk management = 0.95 (5% reduction in inoperability propagation)
policy_rm_1 <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751, risk_management = 0.95, days_in_year = 365)
print(paste("Total Loss (RM=0.95):", policy_rm_1$total_economic_loss))

# Policy comparison
print(paste("Savings from RM=0.95:", loss_55 - policy_rm_1$total_economic_loss))
