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

DIIM_model <- DIIM(q0, A_star, c_star, x, lockdown_duration = 55, total_duration = 751, days_in_year = 365)

inoperability_evolution <- DIIM_model$inoperability_evolution
max_inoperability <- apply(inoperability_evolution, 1, max)
sorted_indices <- order(max_inoperability, decreasing = TRUE)
sorted_inoperability <- inoperability_evolution[sorted_indices, ]

EL_evolution <- DIIM_model$EL_evolution
max_econ_loss <- apply(EL_evolution, 1, max)
sorted_indices <- order(max_econ_loss, decreasing = TRUE)
sorted_econ_loss <- EL_evolution[sorted_indices, ]

