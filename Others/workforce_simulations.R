library(openxlsx)
library(ggplot2)

# extracting data
IOT_2019_full = read.xlsx('/Users/yeohyy/Library/CloudStorage/GoogleDrive-yeoh.yuyong@gmail.com/My Drive/School/NTU/FYP/Data/2019_full_IOT.xlsx', sheet = "2019IOT")
x = IOT_2019_full[4:(4+107-1), ncol(IOT_2019_full)]
IOT_2019_full = IOT_2019_full[4:(4+107-1), 4:(4+107-1)]
dependence_ratio = read.xlsx('/Users/yeohyy/Library/CloudStorage/GoogleDrive-yeoh.yuyong@gmail.com/My Drive/School/NTU/FYP/Data/2019_full_IOT.xlsx', sheet = "dependence ratio")
dependence_ratio = dependence_ratio[1:(1+107-1), ncol(dependence_ratio)]

A = read.xlsx('/Users/yeohyy/Library/CloudStorage/GoogleDrive-yeoh.yuyong@gmail.com/My Drive/School/NTU/FYP/Data/2019_full_IOT.xlsx', sheet = "A")
A = A[2:(2+107-1), 3:(3+107-1)]
A = as.matrix(sapply(A, as.numeric))


attack_rate = 0.3
q0 = dependence_ratio * attack_rate # sector initial inoperability

# we assume after 2 years the economic activity return to 99% of pre lock down level

A_star = solve(diag(as.vector(x)))%*%A%*%diag(as.vector(x))
qT = q0 * 1/100
a_ii = diag(A_star)
T=55+974


k = log(q0/qT)/(T*(1-a_ii))
K <- diag(as.vector(k))

set.seed(123)
c_star = runif(n=107,min=0,max=0.2)

time_steps = T

num_sectors = length(q0)
inoperability_evolution = matrix(NA,nrow = num_sectors,ncol=time_steps)
inoperability_evolution[,1] = q0

inoperability_evolution[,2] = inoperability_evolution[,2-1] + K %*% (A_star %*% inoperability_evolution[,2-1] + c_star - inoperability_evolution[,2-1]) 
inoperability_evolution[,3] = inoperability_evolution[,3-1] + K %*% (A_star %*% inoperability_evolution[,3-1] + c_star - inoperability_evolution[,3-1]) 


for (t in 2:time_steps) {
  if (t<=55){
    inoperability_evolution[,t] = inoperability_evolution[,t-1] + K %*% (A_star %*% inoperability_evolution[,t-1] + c_star - inoperability_evolution[,t-1]) 
  }
  else { # after 55 days, c* become 0 as no more external shock
    inoperability_evolution[,t] = inoperability_evolution[,t-1] + K %*% (A_star %*% inoperability_evolution[,t-1] - inoperability_evolution[,t-1])
  }
}























