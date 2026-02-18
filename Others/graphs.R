set.seed(123)
A_star <- matrix(runif(42 * 42), nrow = 42, ncol = 42)
dim(A_star)  # 42 42

# Dependencies
suppressPackageStartupMessages({
  library(plotly)
  library(viridisLite)
})

# If A_star not defined, create a placeholder
if (!exists("A_star")) {
  set.seed(42)
  A_star <- matrix(runif(42*42), nrow = 42, ncol = 42)
}

n <- nrow(A_star)
i_idx <- rep(1:n, times = n)
j_idx <- rep(1:n, each  = n)
z_val <- as.numeric(A_star)

# Bar half-size
hs <- 0.45

# Color by height
z_min <- min(z_val, na.rm = TRUE)
z_max <- max(z_val, na.rm = TRUE)
rng <- if (z_max > z_min) (z_val - z_min) / (z_max - z_min) else rep(0, length(z_val))
cols <- viridis(256)[pmax(1, round(rng * 255) + 1)]

# Function to build one bar as a prism (8 vertices, 12 triangles)
bar_mesh <- function(x, y, h, color) {
  if (is.na(h) || h <= 0) return(NULL)
  x0 <- x - hs; x1 <- x + hs
  y0 <- y - hs; y1 <- y + hs
  z0 <- 0; z1 <- h
  
  vx <- c(x0, x1, x1, x0, x0, x1, x1, x0)
  vy <- c(y0, y0, y1, y1, y0, y0, y1, y1)
  vz <- c(z0, z0, z0, z0, z1, z1, z1, z1)
  
  # Triangles by vertex indices (0-based for plotly)
  tris <- rbind(
    c(0,1,2), c(0,2,3),       # bottom
    c(4,6,5), c(4,7,6),       # top
    c(0,4,5), c(0,5,1),       # side 1
    c(1,5,6), c(1,6,2),       # side 2
    c(2,6,7), c(2,7,3),       # side 3
    c(3,7,4), c(3,4,0)        # side 4
  )
  
  list(x = vx, y = vy, z = vz, i = tris[,1], j = tris[,2], k = tris[,3],
       facecolor = rep(color, nrow(tris)))
}

# Build all bars
parts <- vector("list", length(z_val))
for (k in seq_along(z_val)) parts[[k]] <- bar_mesh(i_idx[k], j_idx[k], z_val[k], cols[k])
parts <- Filter(Negate(is.null), parts)

# Combine meshes into one trace by concatenating vertices and reindexing triangles
vx <- vy <- vz <- c()
ii <- jj <- kk <- c()
fc <- c()
offset <- 0
for (p in parts) {
  nv <- length(p$x)
  vx <- c(vx, p$x); vy <- c(vy, p$y); vz <- c(vz, p$z)
  ii <- c(ii, p$i + offset); jj <- c(jj, p$j + offset); kk <- c(kk, p$k + offset)
  fc <- c(fc, p$facecolor)
  offset <- offset + nv
}

fig <- plot_ly(type = "mesh3d",
               x = vx, y = vy, z = vz,
               i = ii, j = jj, k = kk,
               facecolor = fc,
               showscale = FALSE) |>
  layout(
    scene = list(
      xaxis = list(title = "Sector number", tickmode = "array",
                   tickvals = c(1, seq(5, n, 5), n)),
      yaxis = list(title = "Sector number", tickmode = "array",
                   tickvals = c(1, seq(5, n, 5), n)),
      zaxis = list(title = "A*"),
      camera = list(eye = list(x = 1.6, y = 1.2, z = 0.8))
    ),
    title = "3D bar plot of A* (OpenGL-free via plotly)"
  )

fig





















