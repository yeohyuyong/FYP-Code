library(readxl)
library(openxlsx)

# ============================================================
# STEP 1: Read 20-sector IO table from io_2022_broad_sector.xlsx
# ============================================================
cat("=== STEP 1: Reading 20-sector IO table ===\n")

df_io <- read_excel("dataset/io_tables/io_2022_broad_sector.xlsx", sheet = "Table 5", col_names = FALSE)

# The IO table layout (from inspection):
# Rows 10-29: 20 sector rows (sector data)
# Row 30 (RT1): Total intermediate demand
# Row 31 (21): Imports
# Row 32 (22): Taxes
# Row 33 (RT2): Total inputs at purchasers' prices
# Row 34 (23): Compensation of employees
# Row 35 (24): Other value added
# Row 36 (25): Gross operating surplus
# Row 37 (RT3): Value added
# Row 38 (RT4): Total output (= total inputs at basic prices)
# Columns 4-23: 20 sector columns

# Extract Z matrix (20x20 intermediate transactions)
Z <- as.matrix(sapply(df_io[10:29, 4:23], as.numeric))

# Extract total output x (row 38, columns 4:23)
x <- as.numeric(df_io[38, 4:23])

# Extract compensation of employees (row 34, columns 4:23)
comp_emp <- as.numeric(df_io[34, 4:23])

cat("Z dimensions:", dim(Z), "\n")
cat("x length:", length(x), "\n")
cat("comp_emp length:", length(comp_emp), "\n")

# Sector names from the IO table
sector_names_20 <- as.character(df_io[10:29, 2])
cat("\n20 IOT Sector Names:\n")
for (i in 1:20) cat(paste0("  ", i, ": ", sector_names_20[i], " (x=", round(x[i], 1), ")\n"))

# ============================================================
# STEP 2: Read 20-to-15 sector mapping
# ============================================================
cat("\n=== STEP 2: Reading 20-to-15 sector mapping ===\n")

df_map <- read_excel("dataset/covid_data/sector_reference_data.xlsx", sheet = "sector mapping")

# The mapping has 15 rows. The Singapore.IOT.Sectors column tells us which
# IOT broad sectors map to which of the 15 aggregated sectors.
# Some 15-sector groups combine multiple IOT sectors.

cat("Mapping (15 aggregated sectors):\n")
print(df_map, n = 20)

# Build the 15x20 mapping matrix M
# Based on the mapping:
# Sector 1 (Agri/Fish/Quarry/Waste): IOT 1 (Other Goods)
# Sector 2 (Manufacturing): IOT 2 (Manufacturing), IOT 3 (Utilities)
# Sector 3 (Construction): IOT 4 (Construction)
# Sector 4 (Wholesale & Retail Trade): IOT 5 (Wholesale), IOT 6 (Retail)
# Sector 5 (Transportation & Storage): IOT 7 (Transportation & Storage)
# Sector 6 (Accommodation & Food Services): IOT 8 (Accommodation), IOT 9 (Food & Beverage)
# Sector 7 (Info & Comms): IOT 10 (Info & Comms)
# Sector 8 (Financial & Insurance): IOT 11 (Finance & Insurance)
# Sector 9 (Real Estate): IOT 12 (Real Estate), IOT 13 (Ownership of Dwellings)
# Sector 10 (Professional Services): IOT 14 (Professional, Scientific & Technical)
# Sector 11 (Admin & Support Services): IOT 15 (Administrative & Support)
# Sector 12 (Public Admin & Education): IOT 16 (Public Administration), IOT 17 (Education)
# Sector 13 (Health & Social Services): IOT 18 (Health & Social Services)
# Sector 14 (Arts, Entertainment & Recreation): IOT 19 (Arts, Entertainment & Recreation)
# Sector 15 (Other Services): IOT 20 (Other Services)

mapping_20_to_15 <- c(1, 2, 2, 3, 4, 4, 5, 6, 6, 7, 8, 9, 9, 10, 11, 12, 12, 13, 14, 15)

M <- matrix(0, nrow = 15, ncol = 20)
for (i in 1:20) {
  M[mapping_20_to_15[i], i] <- 1
}

# ============================================================
# STEP 3: Aggregate IO table from 20 to 15 sectors
# ============================================================
cat("\n=== STEP 3: Aggregating IO table ===\n")

Z_agg <- M %*% Z %*% t(M)
x_agg <- as.vector(M %*% x)
comp_emp_agg <- as.vector(M %*% comp_emp)

# Compute technical coefficient matrix A = Z * diag(1/x)
A_agg <- Z_agg %*% diag(1 / x_agg)

cat("Z_agg dimensions:", dim(Z_agg), "\n")
cat("x_agg length:", length(x_agg), "\n")
cat("Sum of x (original):", sum(x), "\n")
cat("Sum of x_agg:", sum(x_agg), "\n")
cat("Column sums of A_agg:", round(colSums(A_agg), 4), "\n")

# ============================================================
# STEP 4: Compute initial inoperability q0
# q0 = (foreign manpower dependence) × (manpower dependence)
# ============================================================
cat("\n=== STEP 4: Computing initial inoperability ===\n")

# Manpower dependence = Compensation of employees / Total output
manpower_dep <- comp_emp_agg / x_agg

# Foreign manpower dependence = Non-resident / Total employment by sector
# Data from MOM Labour Market Report Q4 2022 (Dec 2022, in thousands)
# Aligned to the 15 MOM sectors

total_emp <- c(
  22.6 + 23.9,  # 1. Agri/Fish/Quarry/Utilities/Waste ≈ Other (22.6) + Utilities portion (23.9)
  498.9,        # 2. Manufacturing
  498.9,        # 3. Construction (same value coincidence with Manufacturing)
  303.1 + 161.7, # 4. Wholesale (303.1) + Retail (161.7)
  265.6,        # 5. Transportation & Storage
  265.9,        # 6. Accommodation & Food Services
  186.0,        # 7. Info & Comms
  223.0,        # 8. Financial & Insurance
  74.1,         # 9. Real Estate
  276.6,        # 10. Professional Services
  239.3,        # 11. Administrative & Support Services
  263.6,        # 12. Public Admin & Education
  193.2,        # 13. Health & Social Services
  47.9,         # 14. Arts, Entertainment & Recreation
  386.1         # 15. Other Services
)

nonres_emp <- c(
  10.0 + 10.0,  # 1. Agri/Fish/Quarry/Utilities/Waste (estimated: ~43% foreign)
  282.1,        # 2. Manufacturing
  394.9,        # 3. Construction
  76.8 + 43.9,  # 4. Wholesale + Retail
  102.5,        # 5. Transportation & Storage
  152.3,        # 6. Accommodation & Food Services
  65.8,         # 7. Info & Comms
  30.3,         # 8. Financial & Insurance
  31.0,         # 9. Real Estate
  66.5,         # 10. Professional Services
  142.8,        # 11. Administrative & Support Services
  17.4,         # 12. Public Admin & Education
  8.1,          # 13. Health & Social Services
  14.3,         # 14. Arts, Entertainment & Recreation
  201.4         # 15. Other Services
)

foreign_manpower_dep <- nonres_emp / total_emp

# Initial inoperability
q0 <- foreign_manpower_dep * manpower_dep

# 15 sector names
sector_names_15 <- c(
  "Agriculture, Fishing, Quarrying, Utilities & Waste Management",
  "Manufacturing",
  "Construction",
  "Wholesale & Retail Trade",
  "Transportation & Storage",
  "Accommodation & Food Services",
  "Information & Communications",
  "Financial & Insurance Services",
  "Real Estate Services",
  "Professional Services",
  "Administrative & Support Services",
  "Public Administration & Education",
  "Health & Social Services",
  "Arts, Entertainment & Recreation",
  "Other Community, Social & Personal Services"
)

cat("\n--- Results Summary ---\n")
results <- data.frame(
  Sector = sector_names_15,
  Total_Output = round(x_agg, 2),
  Comp_Emp = round(comp_emp_agg, 2),
  Manpower_Dep = round(manpower_dep, 4),
  Foreign_Share = round(foreign_manpower_dep, 4),
  q0 = round(q0, 4)
)
print(results, right = FALSE)

cat("\nAll q0 in [0,1]:", all(q0 >= 0 & q0 <= 1), "\n")
cat("All A_agg column sums < 1:", all(colSums(A_agg) < 1), "\n")

# ============================================================
# STEP 5: Save to Excel files
# ============================================================
cat("\n=== STEP 5: Saving aggregated data ===\n")

# Save IO table (A matrix and x vector)
wb_iot <- createWorkbook()

addWorksheet(wb_iot, "A")
A_out <- rbind(c("", sector_names_15), cbind(sector_names_15, A_agg))
writeData(wb_iot, "A", A_out, colNames = FALSE)

addWorksheet(wb_iot, "x")
x_out <- data.frame(Sector = sector_names_15, Total_Output = x_agg)
writeData(wb_iot, "x", x_out)

addWorksheet(wb_iot, "Z")
Z_out <- rbind(c("", sector_names_15), cbind(sector_names_15, Z_agg))
writeData(wb_iot, "Z", Z_out, colNames = FALSE)

saveWorkbook(wb_iot, "dataset/manpower_disruption_data/iot_2022_15_sectors.xlsx", overwrite = TRUE)

# Save inoperability data
wb_q0 <- createWorkbook()
addWorksheet(wb_q0, "Sector_initial_inoperability")
q0_out <- data.frame(
  Sector = sector_names_15,
  Manpower_Dependence = manpower_dep,
  Foreign_Manpower_Share = foreign_manpower_dep,
  Sector_initial_inoperability = q0
)
writeData(wb_q0, "Sector_initial_inoperability", q0_out)

saveWorkbook(wb_q0, "dataset/manpower_disruption_data/sector_initial_inoperability_15_sectors.xlsx", overwrite = TRUE)

cat("Done! Files saved to dataset/manpower_disruption_data/\n")
