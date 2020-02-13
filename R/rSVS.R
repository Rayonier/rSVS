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

#' Demonstrate Stand Visualiztion on several stand types
#'
#' Display one of several included SVS files to demonstrate different stand types.
#'
#' @param StandType Name of stand to display
#' @author Jim McCarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' SVS_Demo( 'SouthernPine' )
#' SVS_Demo( 'Douglas-fir' )
#' @export
SVS_Demo <- function( Stand=NULL ) {
    svsexe <- system.file( "bin/svs", "winsvs.exe", package="rSVS" )
    if( is.null(Stand) ) {
        print( paste0( "Please pick one of: BottomlandHardwood, Douglas-fir, LodgepolePine, MixedConifer, MontaneOak-Hickory, ",
                       "PacificSilverFir-Hemlock, Redwood, SouthernPine, or Spruce-Fir" ) )
        return('SVS_Demo() exited.')
    } else if( Stand=='BottomlandHardwood' ) {
        svsfile <- system.file( "bin", "BottomlandHardwood.svs", package="rSVS" )
    } else if( Stand == 'Douglas-fir' ) {
        svsfile <- system.file( "bin", "Douglas-fir.svs", package="rSVS" )
    } else if( Stand == 'LodgepolePine' ) {
        svsfile <- system.file( "bin", "LodgepolePine.svs", package="rSVS" )
    } else if( Stand == 'MixedConifer' ) {
        svsfile <- system.file( "bin", "MixedConifer.svs", package="rSVS" )
    } else if( Stand == 'MontaneOak-Hickory' ) {
        svsfile <- system.file( "bin", "MontaneOak-Hickory.svs", package="rSVS" )
    } else if( Stand == 'PacificSilverFir-Hemlock' ) {
        svsfile <- system.file( "bin", "PacificSilverFir-Hemlock.svs", package="rSVS" )
    } else if( Stand == 'Redwood' ) {
        svsfile <- system.file( "bin", "Redwood.svs", package="rSVS" )
    } else if( Stand == 'SouthernPine' ) {
        svsfile <- system.file( "bin", "SouthernPine.svs", package="rSVS" )
    } else if( Stand == 'Spruce-Fir') {
        svsfile <- system.file( "bin", "Spruce-Fir.svs", package="rSVS" )
    }
    cmdline <- paste0( svsexe, " ", svsfile )
    #print( cmdline )
    system( cmdline, invisible=FALSE )
    return( 'SVS existed.' )
}

#' Visualize stand using the Stand Visualization System (SVS)
#'
#' The SVS() function will create stand level visualizations of data frames containing appropriate information.  The
#' function has the abillity to generate coordinates if they are not provided.
#'
#' StandViz internal format:
#' stand, year, species, treeno, x, y, dbh, height, crownratio, crownradius, tpa, live, status, condition, svsstatus, bearing, brokenht, brokenoffset, dmr, leanangle, rootwad
#'
#' rsvs data frame format:
#' stand, year, treeno, species, dbh, height, crownratio, crownradius, tpa, x, y, live, status, condition, (svsstatus, brokenht, brokenoffset, bearing, dmr, leanangle, rootwad)
#'
#' @param data compatible data frame (see details)
#' @param output what and were to product output (SVS | BITMAP | WEB | CSV )
#' @param clumped if TRUE generate clumped coordinates
#' @param random if TRUE generate random coordiantes
#' @param row if TRUE generate coordinates for rows (plantation)
#' @param uniform if TRUE generate uniform coordinates
#' @param randommess control "noise" of coordinates generated
#' @param clumpiness adjust clump strength
#' @param clumpratio adjust number/size of clumps
#' @author James Mccarter \email{jim.mccarter@@rayonier.com}
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
    demosvsfile <- system.file( "bin", "SouthernPine.svs", package="rSVS" )
    print( demosvsfile )
    cmdline <- paste0( svsexe, " ", demosvsfile )
    print( cmdline )
    system( cmdline, invisible=FALSE )
    # if reticulate
    # library(reticulate)
    # SVS <- import_from_path( "StandViz", path="inst/python" ) 
    # else
    # pyexe <- system.file( "bin/python38", "python.exe", package="rSVS" )
    # StandViz <- system.file( "python", "StandViz.py", package="rSVS" )
    # cmdline <- paste0( pyexe, " ", StandViz, " arguments go here" )
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

