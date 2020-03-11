# tests for rSVS package to make sure various components of package are available

test_that( "BMP2PNG.exe files exist", {
  expect_is( system.file( "bin", "BMP2PNG.exe", package="rSVS" ), 'character' )
})

test_that( "python38.zip file exists", {
  expect_is( system.file( "bin", "python38.zip", package="rSVS" ), 'character' )
})

test_that( "rSVS_Species.csv files exist", {
  expect_is( system.file( "bin", "rSVS_Species.csv", package="rSVS" ), 'character' )
})

test_that( "rSVS_Species.xlsx files exist", {
  expect_is( system.file( "bin", "rSVS_Species.xlsx", package="rSVS" ), 'character' )
})

test_that( "unzip.exe files exist", {
  expect_is( system.file( "bin", "unzip.exe", package="rSVS" ), 'character' )
})

# test to make sure SVS files exist
test_that( "winsvs.exe files exist", {
  expect_is( system.file( "bin/SVS", "winsvs.exe", package="rSVS" ), 'character' )
})

test_that( "tbl2svs.dll files exist", {
  expect_is( system.file( "bin/SVS", "tbl2svs.dll", package="rSVS" ), 'character' )
})

test_that( "FIA.trf files exist", {
  expect_is( system.file( "bin/SVS", "FIA.trf", package="rSVS" ), 'character' )
})

test_that( "NRCS.trf files exist", {
  expect_is( system.file( "bin/SVS", "NRCS.trf", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: BottomlandHardwood.svs", {
  expect_is( system.file( "extdata", "BottomlandHardwood.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: Douglas-fir.svs", {
  expect_is( system.file( "extdata", "Douglas-fir.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: LodgepolePine.svs", {
  expect_is( system.file( "extdata", "LodgepolePine.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: MixedConifer.svs", {
  expect_is( system.file( "extdata", "MixedConifer.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: MontaneOak-Hickory.svs", {
  expect_is( system.file( "extdata", "MontaneOak-Hickory.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: PacificSilverFir-Hemlock.svs", {
  expect_is( system.file( "extdata", "PacificSilverFir-Hemlock.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: Redwood.svs", {
  expect_is( system.file( "extdata", "Redwood.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: SouthernPine.svs", {
  expect_is( system.file( "extdata", "SouthernPine.svs", package="rSVS" ), 'character' )
})

test_that( "Example data files exist: Spruce-fir.svs", {
  expect_is( system.file( "extdata", "Spruce-fir.svs", package="rSVS" ), 'character' )
})

test_that( "Python component StandViz.py exists", {
  expect_is( system.file( "python", "StandViz.py", package="rSVS" ), 'character' )
})

# tests to make sure parts of package are working correctly
