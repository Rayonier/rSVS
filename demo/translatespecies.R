# translating species demo

Trees <- read.csv(system.file("extdata", "SouthernPine2.csv", package="rSVS"))
head(Trees)    # display first 6 rows of Trees data frame

TreesFIA <- Translate_NRCS2FIA( Trees )    # convert using NRCS2FIA and save as TreesFIA
head(TreesFIA)

TreesNRCS <- Translate_FIA2NRCS(TreesFIA)
head(TreesNRCS)
