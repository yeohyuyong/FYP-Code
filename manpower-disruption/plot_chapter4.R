# =============================================================================
# Manpower Disruption Case Study — Chapter 4 Report Figures
# =============================================================================
# Generates the 2 plots used in Chapter 4 of the FYP report:
#   1. manpower_inoperability.png  — Inoperability evolution (top 10 highlighted)
#   2. manpower_econ_loss.png      — Economic loss evolution (top 10 highlighted)
#
# Usage: source("manpower-disruption/plot_chapter4.R") or Rscript manpower-disruption/plot_chapter4.R
# =============================================================================

if (!file.exists("functions.R")) setwd("..")

library(ggplot2)
library(reshape2)

# Run manpower-disruption.R to get inoperability_evolution and EL_evolution
source("manpower-disruption/manpower-disruption.R")

inoperability_evolution <- DIIM_model$inoperability_evolution
EL_evolution            <- DIIM_model$EL_evolution

num_sectors <- nrow(inoperability_evolution)
num_days    <- ncol(inoperability_evolution)

# ---- Sector labels (short names for plots) ----
sector_labels <- c(
  "Agriculture & Utilities", "Manufacturing", "Construction",
  "Wholesale & Retail", "Transportation & Storage", "Accommodation & Food",
  "Info & Communications", "Financial Services", "Real Estate",
  "Professional Services", "Admin & Support", "Public Admin & Education",
  "Health & Social", "Arts & Recreation", "Other Community Services"
)

# ---- Helper: build long data frame ----
sector_display <- paste0("S", 1:num_sectors, ": ", sector_labels)

build_long_df <- function(mat, value_name) {
  df <- as.data.frame(t(mat))
  colnames(df) <- sector_display
  df$Day <- 1:num_days
  melt(df, id.vars = "Day", variable.name = "Sector", value.name = value_name)
}

# ---- Helper: evolution line chart with top-N highlighted ----
plot_evolution <- function(mat, value_col, title, subtitle, ylab, top_n = 10,
                           rank_metric_fn = function(m) apply(m, 1, max)) {
  metric     <- rank_metric_fn(mat)
  sorted_idx <- order(metric, decreasing = TRUE)
  top_sectors <- sector_display[sorted_idx[1:top_n]]

  long_df <- build_long_df(mat, value_col)
  long_df$is_top <- long_df$Sector %in% top_sectors
  long_df$Sector <- factor(long_df$Sector, levels = sector_display[sorted_idx])

  bg_df <- long_df[!long_df$is_top, ]
  fg_df <- long_df[long_df$is_top, ]

  ggplot() +
    geom_line(data = bg_df, aes(x = Day, y = .data[[value_col]], group = Sector),
              color = "grey80", linewidth = 0.5, alpha = 0.7) +
    geom_line(data = fg_df, aes(x = Day, y = .data[[value_col]], color = Sector),
              linewidth = 0.9) +
    scale_color_manual(
      values = setNames(
        c("#E69F00", "#009E73", "#0072B2", "#CC79A7", "#D55E00",
          "#56B4E9", "#F0E442", "#E34234", "#999999", "#6BAED6"),
        top_sectors),
      breaks = top_sectors) +
    labs(title = title, subtitle = subtitle, x = "Days", y = ylab, color = "Sector") +
    theme_minimal(base_size = 13) +
    theme(plot.title = element_text(face = "bold", size = 16),
          plot.subtitle = element_text(color = "grey50", size = 11),
          legend.position = "bottom", legend.text = element_text(size = 8),
          panel.grid.minor = element_blank(),
          plot.background = element_rect(fill = "white", color = NA),
          panel.background = element_rect(fill = "white", color = NA)) +
    guides(color = guide_legend(nrow = 5))
}

# ---- Figure 1: Inoperability Evolution ----
p_inop <- plot_evolution(inoperability_evolution, "Inoperability",
  "Manpower Disruption: Inoperability Evolution",
  "Top 10 sectors highlighted; remaining sectors in grey", "Inoperability")
ggsave("manpower-disruption/manpower_inoperability.png", p_inop, width = 10, height = 7, dpi = 300)
cat("Saved: manpower_inoperability.png\n")

# ---- Figure 2: Economic Loss Evolution ----
p_econ <- plot_evolution(EL_evolution, "Economic_Loss",
  "Manpower Disruption: Economic Loss Evolution",
  "Top 10 sectors highlighted; remaining sectors in grey", "Economic Loss")
ggsave("manpower-disruption/manpower_econ_loss.png", p_econ, width = 10, height = 7, dpi = 300)
cat("Saved: manpower_econ_loss.png\n")

cat("\nAll Chapter 4 figures saved to manpower-disruption/\n")
