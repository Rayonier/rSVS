# demo showing loading of example data provided in package

# load .svs example file
df <- system.file( "extdata", "Douglas-fir.svs", package="rSVS" )
svs(df)

