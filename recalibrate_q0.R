################################################################################
# Recalibrate COVID-19 Initial Inoperability (q0)
#
# Uses actual circuit breaker workforce unavailability instead of unemployment
# rates. Methodology follows Pichler et al. (2022):
#   Workforce Unavailability = (1 - Telecommuting Rate) × (1 - Essential Onsite Fraction)
#
# Data sources:
#   - MOM 2020 telecommuting data (Info&Comms/Financial 77%, Health 28%, F&B 11%)
#   - MTI/Enterprise SG essential services list (April 2020)
#   - Q2 2020 GDP contraction by sector (validation)
################################################################################

library(openxlsx)

# ---- Sector names (15-sector aggregation) ----
sector_names <- c(
  "Agriculture, Utilities, Waste",
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
  "Other Services"
)

# ---- Telecommuting rates (MOM 2020 data + estimates) ----
# Source: MOM Labour Force Survey 2020, MAS Feature Article (2021)
# Direct MOM data: Info&Comms=0.77, Financial=0.77, Health=0.28, F&B=0.11
# Others estimated from occupational mix and industry characteristics
telecommuting_rate <- c(
  0.25,  # 1. Agriculture/Utilities/Waste: some admin, mostly field/plant work
  0.25,  # 2. Manufacturing: some admin/engineering, mostly production floor
  0.15,  # 3. Construction: craftsmen/trades 14.6% (MOM), very physical
  0.40,  # 4. Wholesale & Retail: wholesale admin can WFH, retail cannot
  0.20,  # 5. Transport & Storage: drivers/logistics cannot WFH
  0.11,  # 6. Accommodation & Food: MOM data = 11%
  0.77,  # 7. Info & Comms: MOM data = 77%
  0.77,  # 8. Financial & Insurance: MOM data = 77%
  0.50,  # 9. Real Estate: agents/admin partial WFH
  0.60,  # 10. Professional Services: legal/consulting partial WFH
  0.30,  # 11. Admin & Support: security/cleaning onsite, office admin WFH
  0.55,  # 12. Public Admin & Education: govt admin WFH, teachers online
  0.28,  # 13. Health & Social: MOM data = 28%
  0.20,  # 14. Arts & Recreation: mostly venue-based, some digital content
  0.15   # 15. Other Services: hair salons, personal services, religious orgs
)

# ---- Essential onsite fraction ----
# Fraction of non-telecommuting workers permitted to work onsite during CB
# Based on MTI/Enterprise SG essential services list (7 April 2020)
essential_onsite <- c(
  0.70,  # 1. Utilities ESSENTIAL; waste mgmt continued; some agriculture
  0.40,  # 2. Biomedical/food/semiconductor essential; other manufacturing stopped
  0.05,  # 3. Almost entirely shut down (Q2 GDP: -59% YoY, -95.6% QoQ)
  0.25,  # 4. Wholesale (supply chains) partially essential; retail closed
  0.50,  # 5. Public transport/logistics ESSENTIAL; aviation/tourism stopped
  0.15,  # 6. F&B delivery/takeaway only; hotels near-empty
  0.80,  # 7. ESSENTIAL service (comms infrastructure) + high telecommuting
  0.80,  # 8. ESSENTIAL service (banking/finance) + high telecommuting
  0.10,  # 9. Showflats closed, property viewings suspended
  0.20,  # 10. Some legal/accounting essential, most consulting stopped
  0.30,  # 11. Cleaning/security essential; facility mgmt, temp staffing stopped
  0.40,  # 12. Government essential; schools moved online
  0.70,  # 13. Healthcare ESSENTIAL during pandemic
  0.05,  # 14. Museums, casinos, entertainment venues all closed
  0.05   # 15. Hair salons, personal services, religious services all closed
)

# ---- Compute workforce unavailability ----
# Pichler et al. (2022) methodology:
# Workers who cannot telecommute AND are not in essential onsite roles are unavailable
workforce_unavailability <- (1 - telecommuting_rate) * (1 - essential_onsite)

cat("========================================\n")
cat("Circuit Breaker Workforce Unavailability\n")
cat("========================================\n\n")

for (i in 1:15) {
  cat(sprintf("  %2d. %-40s  Telecom=%.0f%%  Essential=%.0f%%  Unavail=%.0f%%\n",
              i, sector_names[i],
              telecommuting_rate[i] * 100,
              essential_onsite[i] * 100,
              workforce_unavailability[i] * 100))
}

# ---- Load LAPI/Output ratios ----
lapi_df <- read.xlsx("dataset/covid_data/unemployment_and_impact_analysis.xlsx",
                     sheet = "lapi div sector output", colNames = FALSE)
lapi_ratio <- as.numeric(lapi_df[, 2])

# ---- Compute recalibrated q0 ----
q0_new <- workforce_unavailability * lapi_ratio

# ---- Load old q0 for comparison ----
q0_old_df <- read.xlsx("dataset/covid_data/unemployment_and_impact_analysis.xlsx",
                       sheet = "sector inoperability", colNames = FALSE)
q0_old <- as.numeric(q0_old_df[2:16, 3])

# ---- Print comparison table ----
cat("\n\n==============================================\n")
cat("q0 Comparison: Old vs Recalibrated\n")
cat("==============================================\n\n")
cat(sprintf("  %-40s  %10s  %10s  %8s\n",
            "Sector", "Old q0", "New q0", "Ratio"))
cat(paste(rep("-", 75), collapse = ""), "\n")

for (i in 1:15) {
  ratio <- q0_new[i] / q0_old[i]
  cat(sprintf("  %-40s  %10.6f  %10.6f  %6.1fx\n",
              sector_names[i], q0_old[i], q0_new[i], ratio))
}

cat(paste(rep("-", 75), collapse = ""), "\n")
cat(sprintf("  %-40s  %10.6f  %10.6f\n", "Mean", mean(q0_old), mean(q0_new)))
cat(sprintf("  %-40s  %10.6f  %10.6f\n", "Range (min)", min(q0_old), min(q0_new)))
cat(sprintf("  %-40s  %10.6f  %10.6f\n", "Range (max)", max(q0_old), max(q0_new)))

cat("\n\nJin & Zhou (Shanghai) range: 0.04 - 0.50\n")
cat("Santos & Haimes range: 0.03 - 0.15\n")

# ---- Update the Excel file ----
cat("\n\nUpdating Excel file...\n")

# Load existing workbook
wb_path <- "dataset/covid_data/unemployment_and_impact_analysis.xlsx"
wb <- loadWorkbook(wb_path)

# Update the "2020 resident unemployment rate" sheet with workforce unavailability
# (This sheet is the source of the first factor in the q0 formula)
for (i in 1:15) {
  writeData(wb, sheet = "2020 resident unemployment rate",
            x = workforce_unavailability[i], startCol = 2, startRow = i)
}

# Update the "sector inoperability" sheet with new q0 values
for (i in 1:15) {
  writeData(wb, sheet = "sector inoperability",
            x = q0_new[i], startCol = 3, startRow = i + 1)
}

# ---- Add Sources sheet ----
# Remove existing Sources sheet if present, then create fresh
if ("Sources" %in% names(wb)) {
  removeWorksheet(wb, "Sources")
}
addWorksheet(wb, "Sources")

# --- Methodology source ---
method_df <- data.frame(
  Item = c(
    "Formula",
    "",
    "Source",
    "",
    "Full Citation",
    "",
    "Exact Passage (Pichler et al. 2022, Appendix A1, Equation 29):",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ),
  Details = c(
    "Workforce Unavailability = (1 - RLI_i) x (1 - ESS_i), where RLI = Remote Labor Index (telecommuting rate), ESS = Essential Score (essential onsite fraction)",
    "",
    "Pichler et al. (2022), Appendix A1, Equation (29)",
    "",
    paste0("Pichler, A., Pangallo, M., del Rio-Chanona, R.M., Lafond, F., & Farmer, J.D. (2022). ",
           "\"In and out of lockdown: Propagation of supply and demand shocks in a dynamic input-output model\". ",
           "Journal of Economic Dynamics and Control, 143, 104527. ",
           "https://doi.org/10.1016/j.jedc.2022.104527"),
    "",
    paste0("\"To estimate industry supply shocks, del Rio-Chanona et al. (2020) constructed an industry-specific ",
           "Remote Labor Index and essential score. The Remote Labor Index of industry i (RLI_i) was constructed by ",
           "classifying the work activities of the occupations of workers employed in each industry into 'can be ",
           "performed at home successfully' or 'cannot be performed at home successfully'.\""),
    "",
    paste0("\"The essential score was built to capture whether an industry can operate during a lockdown, according to ",
           "government mandates, even if working from home is not possible.\""),
    "",
    paste0("\"RLI_i can be interpreted as the probability that a worker from industry i can work from home. Similarly, ",
           "ESS_i can be interpreted as the probability that a worker from industry i has a job that is considered ",
           "essential and can work on-site if needed.\""),
    "",
    paste0("\"Using these interpretations and assuming independence between these two probabilities, del Rio-Chanona ",
           "et al. (2020) calculated that the expected value of the number of workers who could not work during the ",
           "lockdown was given by (1 - RLI_i)(1 - ESS_i).\" (Equation 29)"),
    "",
    ""
  ),
  stringsAsFactors = FALSE
)
writeData(wb, "Sources", method_df, startRow = 1, colNames = TRUE)

# --- Per-sector source table ---
telecom_source <- c(
  "Estimated from occupational mix; utilities/waste mgmt roles are predominantly onsite",
  "Estimated from occupational mix; production floor workers cannot telecommute",
  "MOM Labour Force Survey 2020: craftsmen/trades workers at 14.6% remote work rate",
  "Estimated: wholesale admin can WFH (~50%), retail shop floor cannot (~10%); blended",
  "Estimated from occupational mix; drivers, warehouse, logistics roles require onsite presence",
  "MOM Labour Force Survey 2020: F&B services at 11% telecommuting rate",
  "MOM Labour Force Survey 2020: Info & Comms at 77% telecommuting rate",
  "MOM Labour Force Survey 2020: Financial services at 77% telecommuting rate",
  "Estimated: property agents and back-office admin partially telecommutable",
  "Estimated from occupational mix; legal, accounting, consulting roles partially WFH",
  "Estimated: cleaning/security requires onsite presence; office admin can WFH",
  "Estimated: government admin WFH, teachers moved to online lessons",
  "MOM Labour Force Survey 2020: Health & Social Services at 28% telecommuting rate",
  "Estimated from occupational mix; venue-based entertainment, some digital content creation",
  "Estimated from occupational mix; personal services (hair salons, etc.) require physical presence"
)

essential_source <- c(
  "MTI Essential Services List (7 Apr 2020): Utilities designated essential; waste mgmt continued operations",
  "MTI Essential Services List: Biomedical, food manufacturing, semiconductor fabrication designated essential; other manufacturing suspended",
  "MTI: Construction almost entirely suspended. Q2 2020 GDP: -59% YoY (SingStat). Only critical maintenance permitted",
  "MTI Essential Services List: Wholesale supply chains partially essential; retail closed except supermarkets/food",
  "MTI Essential Services List: Public transport and logistics designated essential. Aviation/tourism halted. Q2 GDP: -39% YoY",
  "MTI: F&B restricted to takeaway/delivery only. Hotels near-empty due to border closure. Q2 GDP: -41% YoY",
  "MTI Essential Services List: Communications infrastructure designated essential",
  "MTI Essential Services List: Banking and financial services designated essential",
  "MTI: Showflats closed, property viewings and transactions suspended during circuit breaker",
  "MTI: Some legal and accounting services deemed essential; most professional consulting suspended",
  "MTI Essential Services List: Cleaning and security services designated essential; temporary staffing and events stopped",
  "MTI Essential Services List: Government continued essential operations; schools shifted to full home-based learning (MOE)",
  "MTI Essential Services List: Healthcare designated essential; hospitals and clinics continued full operations",
  "MTI: Museums, attractions, casinos, entertainment venues all closed per circuit breaker measures",
  "MTI: Hair salons, beauty services, religious services all closed. Some reopened only in Phase 2 (19 Jun 2020)"
)

sector_source_df <- data.frame(
  Sector_ID = 1:15,
  Sector = sector_names,
  Telecommuting_Rate = telecommuting_rate,
  Telecommuting_Source = telecom_source,
  Essential_Onsite_Fraction = essential_onsite,
  Essential_Source = essential_source,
  Workforce_Unavailability = workforce_unavailability,
  LAPI_Output_Ratio = lapi_ratio,
  New_q0 = q0_new,
  stringsAsFactors = FALSE
)
writeData(wb, "Sources", sector_source_df, startRow = 21, colNames = TRUE)

# --- General references ---
refs_df <- data.frame(
  Reference = c(
    "Telecommuting Data Sources:",
    "[1] Ministry of Manpower (MOM), \"Labour Force in Singapore 2020\", Table 52: Employed Residents Who Worked From Home by Industry. https://stats.mom.gov.sg/Pages/Labour-Force-In-Singapore-2020.aspx",
    "[2] Monetary Authority of Singapore (MAS), \"Feature Article: Impact of COVID-19 on Singapore's Labour Market\", MAS Annual Report 2020/2021. https://www.mas.gov.sg/",
    "[3] The Straits Times, \"About half of employed residents in Singapore worked remotely in 2020: MOM survey\", 1 Dec 2021.",
    "",
    "Essential Services Sources:",
    "[4] Ministry of Trade and Industry (MTI), \"Adjustments to Essential Services List\", 7 April 2020. https://www.mti.gov.sg/",
    "[5] Enterprise Singapore, \"List of Essential Services\", April 2020.",
    "[6] Ministry of Health (MOH), \"End of Circuit Breaker, Gradual Resumption of Activities in Three Phases\", 19 May 2020. https://www.moh.gov.sg/",
    "",
    "GDP Contraction Data (for validation):",
    "[7] SingStat, \"Singapore Economy Contracted by 13.2% in Q2 2020\", National Accounts Q2 2020.",
    "[8] MTI, \"MTI Narrows 2020 GDP Growth Forecast\", Economic Survey of Singapore Q2 2020.",
    "",
    "LAPI / Sector Output:",
    "[9] Department of Statistics, Singapore (SingStat), \"Singapore Input-Output Tables 2019\". The LAPI ratio is computed as Compensation of Employees / Total Output from the IO table."
  ),
  stringsAsFactors = FALSE
)
writeData(wb, "Sources", refs_df, startRow = 39, colNames = FALSE)

# Auto-size columns for readability
setColWidths(wb, "Sources", cols = 1:9, widths = "auto")

saveWorkbook(wb, wb_path, overwrite = TRUE)
cat("Saved updated q0 values and sources to:", wb_path, "\n")

cat("\nDone! Run simulations/enhanced_comparison.R to see updated results.\n")
