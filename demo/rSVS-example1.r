# demo showing loading of example data provided in package



# SouthernPine.csv

sp <- system.file( "extdata", "SouthernPine.csv", package="rSVS" )
spo <- read.csv(sp)

# SouthernPine2.csv
sp2 <- system.file( "extdata", "SouthernPine2.csv", package="rSVS" )
sp2o <- read.csv(sp2)

# SouthernPine3.csv
sp3 <- system.file( "extdata", "SouthernPine3.csv", package="rSVS" )
sp3o <- read.csv(sp3)
