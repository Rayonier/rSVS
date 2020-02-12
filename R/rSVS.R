# Package Documentation
#
#' A packge for stand level visualization using the Stand Visualization System (SVS).
#'
#' The rSVS package provides and interface to perform SVS visualiations from R.
#' @docType package
#' @name rSVS
NULL

# You can learn more about package authoring with RStudio at:
#
#   http://r-pkgs.had.co.nz/
#
# Some useful keyboard shortcuts for package authoring:
#
#   Build and Reload Package:  'Ctrl + Shift + B'
#   Check Package:             'Ctrl + Shift + E'
#   Test Package:              'Ctrl + Shift + T'

#' Visualize stand using the Stand Visualization System (SVS)
#'
#' The SVS() function will create stand level visualizations of data frames containing appropriate information.  The
#' function has the abillity to generate coordinates if they are not provided.
#'
#' pystandviz internal format:
#' stand, year, species, treeno, x, y, dbh, height, crownratio, crownradius, tpa, live, status, condition, svsstatus, bearing, brokenht, brokenoffset, dmr, leanangle, rootwad
#'
#' rsvs data frame format:
#' stand, year, treeno, species, dbh, height, crownratio, crownradius, tpa, x, y, live, status, condition, (svsstatus, brokenht, brokenoffset, bearing, dmr, leanangle, rootwad)
#'
#' @param data compatible data frame (see details)
#' @param output what and were to product output
#' @param clumped generate clumped coordinates
#' @param random generate random coordiantes
#' @param row generate coordinates for rows (plantation)
#' @param uniform generate uniform coordinates
#' @param randommess control "noise" of coordinates generated
#' @param clumpiness adjust clump strength
#' @param clumpratio adjust number/size of clumps
#' @author james mccarter
#' @examples
#' svs( d )
#' svs( d, random=true )    # visualize stand in svs using random tree locations
#' svs( d, row=true )       # visualize stand in svs using rows
#' @export
SVS <- function( data, output='svs', clumped=false, random=false, row=true, uniform=false, randomness=null, clumpiness=null, clumpratio=null ) {
    #print( data )
    if( ! "reticulate" %in% .packages() ) print( paste0( "reticulate package NOT loaded" ) )
    if( ! "reticulate" %in% rownames(installed.packages()) ) print( paste0( "reticulate package NOT installed" ) )
    svsexe <- system.file( "bin/svs", "winsvs.exe", package="rSVS" )
    cmdline <- paste0( svsexe, " c:/Users/mccarterj/McCarter/Projects/rSVS/inst/bin/SouthernPine.svs" )
    system( cmdline, invisible=FALSE )
}


#' List known species codes
#'
#' List known species codes.
#'
#' This function simply reads and returns the list of known species codes distributed with the rSVS
#' package (rSVS_Species.csv). This file is also used as the basis for species translation by the FIA2NRCS() and
#' NRCS2FIA() funtions.  The file should be syncronized with the FIA.TFM and NRCS.TFM files, but
#' this is not guaranteed.
#'
#' The rSVS package supports FIA number and NRCS alphabetic codes.  The displayed table will include FIA, NRCS, Genus,
#' Species, Common, and Comment columns.
#'
#' @return DataFrame with species names and codes
#' @examples
#' SVS_Species()
#' head(SVS_Species())
#' length(SVS_Species()$FIA)
#' @export
SVS_Species <- function() {
    Species <- read.csv( system.file( "bin", "rSVS_Species.csv", package="rSVS" ) )     # read rSVS_Species.csv
    return( Species )                                                                   # and return DataFrame
}

#' Convert species codes from FIA number to NRCS code
#'
#' Details
#'
#' @param MyData DataFrame with tree information
#' @examples
#' FIA2NRCS( MyData )
#' @export
FIA2NRCS <- function( DataFrame ) {

}

#' Convert species codes from NRCS code to FIA number
#'
#' Details
#'
#' @param MyData DataFrame with tree information
#' @examples
#' NRCS2FIA( MyData )
#' @export
NRCS2FIA <- function( DataFrame ) {

}

