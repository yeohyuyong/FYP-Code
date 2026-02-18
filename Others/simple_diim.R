library(openxlsx)
library(ggplot2)

data <- read.xlsx("dataset/Input_Output_Tables_2019_2021.xlsx", sheet = "2019IOTable")
iot2019 <- data[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
x <- data[2:(2 + 15 - 1), ncol(data)]
c <- data[2:(2 + 15 - 1), ncol(data) - 1]
x <- as.matrix(sapply(x, as.numeric))
c <- as.matrix(sapply(c, as.numeric))

A <- read.xlsx("dataset/Input_Output_Tables_2019_2021.xlsx", sheet = "A")
A <- A[2:(2 + 15 - 1), 3:(3 + 15 - 1)]
A <- as.matrix(sapply(A, as.numeric))


q0_df <- read.xlsx("dataset/Unemployment_and_Impact_Analysis.xlsx", sheet = "sector inoperability", colNames = FALSE)
q0 <- as.numeric(q0_df[2:16, 3])

# verify x = Ax + c
x_test <- A %*% x + c
x - x_test



# we assume after 2 years the economic activity return to 99% of pre lock down level
A_star <- solve(diag(as.vector(x))) %*% A %*% diag(as.vector(x))
qT <- q0 * 1 / 100
# qT = rep(0.001, times = 15)
a_ii <- diag(A_star)
T <- 55 + 974
k <- log(q0 / qT) / (T * (1 - a_ii))


K <- diag(as.vector(k))

# Number of time steps for simulation (e.g., 90 days)
time_steps <- T

# now we obtain the value of c-c_tilda from the ADB dataset

# c_diff = c-c_tilda
# c_diff = read.xlsx("dataset/Unemployment_and_Impact_Analysis.xlsx", sheet = "c_diff")
# Note: c_diff sheet not found in original file. Using c_star logic instead or placeholder.
if ("c_star" %in% getSheetNames("dataset/Unemployment_and_Impact_Analysis.xlsx")) {
  c_star_df <- read.xlsx("dataset/Unemployment_and_Impact_Analysis.xlsx", sheet = "c_star")
  c_star <- as.numeric(c_star_df[1:15, 7])
} else {
  c_diff <- data.frame(rep(0, 15)) # Placeholder
  c_star <- solve(diag(as.vector(x))) %*% (c_diff)
}
c_star <- c_star

num_sectors <- length(q0)
inoperability_evolution <- matrix(NA, nrow = num_sectors, ncol = time_steps)
inoperability_evolution[, 1] <- q0

for (t in 2:time_steps) {
  if (t <= 55) {
    inoperability_evolution[, t] <- inoperability_evolution[, t - 1] + K %*% (A_star %*% inoperability_evolution[, t - 1] + c_star - inoperability_evolution[, t - 1])
  } else { # after 55 days, c* become 0 as no more external shock
    inoperability_evolution[, t] <- inoperability_evolution[, t - 1] + K %*% (A_star %*% inoperability_evolution[, t - 1] - inoperability_evolution[, t - 1])
  }
}

library(reshape2)
df <- melt(inoperability_evolution)
colnames(df) <- c("Sector", "Time", "Inoperability")

sector_names <- c(
  "1. Agriculture, Fishing, Quarrying, Utilities and Sewerage & Waste Management",
  "2. Manufacturing",
  "3. Construction",
  "4. Wholesale & Retail Trade",
  "5. Transportation & Storage",
  "6. Accommodation & Food Services",
  "7. Information & Communications",
  "8. Financial & Insurance Services",
  "9. Real Estate Services",
  "10. Professional Services",
  "11. Administrative & Support Services",
  "12. Public Administration & Education",
  "13. Health & Social Services",
  "14. Arts, Entertainment & Recreation",
  "15. Other Community, Social & Personal Services"
)

df$Time <- df$Time
df$Sector <- factor(df$Sector, levels = 1:15, labels = sector_names)

# Calculate max inoperability per sector and order descending
max_inoperability <- aggregate(Inoperability ~ Sector, df, max)
max_inoperability <- max_inoperability[order(-max_inoperability$Inoperability), ]
ordered_sectors_inop <- as.character(max_inoperability$Sector)

# Split sectors into three groups of 5 each for inoperability
groups_inop <- split(ordered_sectors_inop, ceiling(seq_along(ordered_sectors_inop) / 5))

# Plot top 5, next 5, last 5 inoperability graphs
for (i in 1:3) {
  sub_df <- df[df$Sector %in% groups_inop[[i]], ]
  p <- ggplot(sub_df, aes(x = Time, y = Inoperability, color = Sector)) +
    geom_line(size = 1) +
    theme_minimal() +
    labs(
      title = paste0("Evolution of Sectoral Inoperability (Group ", i, ")"),
      x = "Time",
      y = "Inoperability",
      color = "Sector"
    ) +
    theme(legend.position = "right")
  print(p)
}

# Economic Loss calculation

# vector x is the planned YEARLY output of all sectors
# convert to planned DAILY output of all sectors
# note that output is in millions

x_daily <- x / 366 # 366 days in 2020
EL_evolution <- t(apply(inoperability_evolution, 1, cumsum)) * as.vector(x_daily)

# Melt economic loss matrix for plotting
EL_df <- melt(EL_evolution)
colnames(EL_df) <- c("Sector", "Time", "EconomicLoss")
EL_df$Sector <- factor(EL_df$Sector, levels = 1:15, labels = sector_names)

# Calculate max economic loss per sector and order descending
max_econ_loss <- aggregate(EconomicLoss ~ Sector, EL_df, max)
max_econ_loss <- max_econ_loss[order(-max_econ_loss$EconomicLoss), ]
ordered_sectors_econ <- as.character(max_econ_loss$Sector)

# Split sectors into three groups of 5 each for economic loss
groups_econ <- split(ordered_sectors_econ, ceiling(seq_along(ordered_sectors_econ) / 5))

# Plot top 5, next 5, last 5 economic loss graphs
for (i in 1:3) {
  sub_EL_df <- EL_df[EL_df$Sector %in% groups_econ[[i]], ]
  p <- ggplot(sub_EL_df, aes(x = Time, y = EconomicLoss, color = Sector)) +
    geom_line(size = 1) +
    theme_minimal() +
    labs(
      title = paste0("Evolution of Economic Loss by Sector (Group ", i, ")"),
      x = "Time",
      y = "Cumulative Economic Loss",
      color = "Sector"
    ) +
    theme(legend.position = "right")
  print(p)
}

total_economic_loss <- sum(EL_evolution[, ncol(EL_evolution)]) # sum of last column in EL_evolution
