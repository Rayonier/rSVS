# tests for rSVS package

test_that( "winsvs.exe files exist", {
  expect_is( system.file( "bin/SVS", "winsvs.exe", package="rSVS" ), 'character' )
})

test_that( "python.exe file exists", {
  expect_is( system.file( "bin", "python.exe", package="rSVS" ), 'character' )
})

