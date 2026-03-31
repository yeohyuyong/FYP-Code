# =============================================================================
# COVID-19 Case Study — Chapter 3 Report Figures
# =============================================================================
# Generates the 3 plots used in Chapter 3 of the FYP report:
#   1. inoperability.png  — Inoperability evolution (top 10 highlighted)
#   2. econ_loss.png      — Economic loss evolution (top 10 highlighted)
#   3. ordinal.png        — Joint impact matrix (inoperability rank vs loss rank)
#
# Usage: source("covid/plot_chapter3.R") or Rscript covid/plot_chapter3.R
# =============================================================================

if (!file.exists("functions.R")) setwd("..")

library(ggplot2)
library(reshape2)
library(ggrepel)

# Run covid-main.R to get inoperability_evolution and EL_evolution
source("covid/covid-main.R")

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
  "COVID-19: Inoperability Evolution",
  "Top 10 sectors highlighted; remaining sectors in grey", "Inoperability")
ggsave("covid/inoperability.png", p_inop, width = 10, height = 7, dpi = 300)
cat("Saved: inoperability.png\n")

# ---- Figure 2: Economic Loss Evolution ----
p_econ <- plot_evolution(EL_evolution, "Economic_Loss",
  "COVID-19: Economic Loss Evolution",
  "Top 10 sectors highlighted; remaining sectors in grey", "Economic Loss")
ggsave("covid/econ_loss.png", p_econ, width = 10, height = 7, dpi = 300)
cat("Saved: econ_loss.png\n")

# ---- Figure 3: Joint Impact Matrix ----
peak_inop <- apply(inoperability_evolution, 1, max)
cum_econ   <- EL_evolution[, ncol(EL_evolution)]
inop_rank  <- rank(-peak_inop)
loss_rank  <- rank(-cum_econ)

assign_zone <- function(ir, lr) {
  if (ir <= 5 & lr <= 5)   return("Top 5")
  if (ir <= 7 & lr <= 7)   return("Top 7")
  if (ir <= 10 & lr <= 10) return("Top 10")
  return("Other")
}

ordinal_df <- data.frame(
  Sector    = sector_labels,
  Loss_Rank = loss_rank,
  Inop_Rank = inop_rank,
  Zone      = factor(mapply(assign_zone, inop_rank, loss_rank),
                     levels = c("Top 5", "Top 7", "Top 10", "Other"))
)

p_ordinal <- ggplot(ordinal_df, aes(x = Loss_Rank, y = Inop_Rank)) +
  geom_point(aes(color = Zone, size = Zone)) +
  geom_text_repel(aes(label = Sector, color = Zone), size = 3.5, fontface = "bold",
                  max.overlaps = 20, seed = 42, box.padding = 0.4, point.padding = 0.3,
                  show.legend = FALSE) +
  scale_color_manual(values = c("Top 5" = "#E34234", "Top 7" = "#E69F00",
                                "Top 10" = "#0072B2", "Other" = "#AAAAAA")) +
  scale_size_manual(values = c("Top 5" = 5, "Top 7" = 4.5, "Top 10" = 4, "Other" = 3.5)) +
  scale_x_continuous(breaks = 1:num_sectors, limits = c(0.5, num_sectors + 0.5)) +
  scale_y_continuous(breaks = 1:num_sectors, limits = c(0.5, num_sectors + 0.5)) +
  labs(title = "Joint Impact Matrix",
       subtitle = "Rank by cumulative economic loss (x-axis) vs rank by peak inoperability (y-axis)",
       x = "Economic Loss Rank (1 = highest loss)",
       y = "Inoperability Rank (1 = highest inoperability)") +
  theme_minimal(base_size = 13) +
  theme(plot.title = element_text(face = "bold", size = 16),
        plot.subtitle = element_text(color = "grey50", size = 11),
        panel.grid.minor = element_blank(), legend.position = "right",
        plot.background = element_rect(fill = "white", color = NA),
        panel.background = element_rect(fill = "white", color = NA)) +
  guides(size = guide_legend(title = "Zone"), color = guide_legend(title = "Zone"))
ggsave("covid/ordinal.png", p_ordinal, width = 10, height = 8, dpi = 300)
cat("Saved: ordinal.png\n")

cat("\nAll Chapter 3 figures saved to covid/\n")


