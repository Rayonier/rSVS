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
    svsexe <- system.file( "bin/svs", "winsvs.exe", package="rSVS" )
    cmdline <- paste0( svsexe, " c:/Users/mccarterj/McCarter/Projects/rSVS/inst/bin/SouthernPine.svs" )
    system( cmdline, invisible=FALSE )
}
