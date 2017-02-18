library(SixSigma)
library(qcc)
library(qualityTools)

library(AcceptanceSampling)

# Simple sampling plan
x <- OC2c(10,1)
x

assess(x, PRP=c(0.01, 0.95), CRP=c(0.05, 0.04))

# Double sampling plan
x <- OC2c(c(125,125), c(1,4), c(4,5),
          pd = seq(0,0.1,0.001))
x

# Simple sampling plan
x <- OC2c(125,1)
x

## A standard binomial sampling plan
x <- OC2c(10,1)
x ## print out a brief summary
plot(x) ## plot the OC curve
plot(x, xlim=c(0,0.5)) ## plot the useful part of the OC curve

## A double sampling plan
x <- OC2c(c(125,125), c(1,4), c(4,5), pd=seq(0,0.1,0.001))
x
plot(x) ## Plot the plan

## Assess whether the plan can meet desired risk points
assess(x, PRP=c(0.01, 0.95), CRP=c(0.05, 0.04))

## A plan based on the Hypergeometric distribution
x <- OC2c(10,1, type="hypergeom", N=5000, pd=seq(0,0.5, 0.025))
plot(x)

## The summary
x <- OC2c(10,1, type="hypergeom", N=5000, pd=seq(0,0.5, 0.1))
summary(x, full=TRUE)

## Plotting against a function which generates P(defective)
xm <- seq(-3, 3, 0.05) ## The mean of the underlying characteristic
x <- OC2c(10, 1, pd=1-pnorm(0, mean=xm, sd=1))
plot(xm, x) ## Plot P(accept) against mean

data("pistonrings")
ss.study.ca(pistonrings$diameter,
            LSL = 73.9, USL = 74.1, Target = 74)

library(qualityTools)
cp(x = pistonrings$diameter,
   lsl = 73.9, usl = 74.1,
   target = 74)
library(qualityTools)
data(defects)
paretoChart(defect, las = 2)

library(xtable)

# PRP
# The Producer Risk Point in the form of a two element numeric vector of the
# form c(pd, pa). The first element, pd, specifies the quality level at which to
# evaluate the plan. The second element, pa, indicates the minimum probability of
# acceptance to be achieved by the plan.
# 
# CRP
# The Consumer Risk Point in the form of a two element numeric vector of the
# form c(pd, pa). The first element, pd, specifies the quality level at which to
# evaluate the plan. The second element, pa, indicates the maximum probability
# of acceptance to be achieved by the plan.

PRP <- c(0.01, 0.90)
CRP <- c(0.02, 0.05)

find.plan(PRP, CRP, type="binomial")
find.plan(PRP, CRP, type="hypergeom", N)
find.plan(PRP, CRP, type="normal", s.type="unknown")