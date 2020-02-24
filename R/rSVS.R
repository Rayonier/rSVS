# Package Documentation
#
#' A packge for stand level visualization using the Stand Visualization System (SVS).
#'
#' This rSVS package provides and interface to perform SVS visualiations from R.
#'
#' The package includes the following functions:
#' \itemize{
#'     \item SVS() - main function for performing visualziations
#'     \item SVS_Demo() - show reginal example visulizations
#'     \item FIA2NRCS() - convert species codes from FIA # to NRCS code
#'     \item NRCS2FIA() - convert species codes from NRCS code to FIA #
#'     \item SVS_Species() - list known species
#' }
#'
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
#' Display one of several stand types using example SVS files includes with the package.
#'
#' The list of stand types includes:
#' \itemize{
#'    \item BottomlandHardwood
#'    \item Douglas-fir
#'    \item LodgepolePine
#'    \item MixedConifer
#'    \item MontaneOak-Hickory
#'    \item PacificSilverFir-Hemlock
#'    \item Redwood
#'    \item SouthernPine
#'    \item Spruce-Fir
#' }
#'
#' @param StandType Name of stand to display
#' @author Jim McCarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' SVS_Demo( 'SouthernPine' )
#' SVS_Demo( 'Douglas-fir' )
#' @export
SVS_Demo <- function( Stand=NULL ) {
    SavedDir <- getwd()                                                             # get and save current working directory
    setwd( path.package("rSVS") )                                                   # set working directory to package location
    svsexe <- system.file( "bin/svs", "winsvs.exe", package="rSVS" )                # get location of winsvs.exe
    if( is.null(Stand) ) {                                                          # if no stand type provided, print message and return
        print( paste0( "Please pick from: BottomlandHardwood, Douglas-fir, LodgepolePine, MixedConifer, ",
                       "MontaneOak-Hickory, PacificSilverFir-Hemlock, Redwood, SouthernPine, or Spruce-Fir" ) )
        return('SVS_Demo() exited.')
    } else if( grepl( 'BottomlandHardwood', Stand, ignore.case=TRUE ) ) {           # check, ignoring case to eliminate some case typos
        svsfile <- system.file( "bin", "BottomlandHardwood.svs", package="rSVS" )   # get location of SVS file
    } else if( grepl( 'Douglas-fir', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "Douglas-fir.svs", package="rSVS" )
    } else if( grepl( 'LodgepolePine', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "LodgepolePine.svs", package="rSVS" )
    } else if( grepl( 'MixedConifer', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "MixedConifer.svs", package="rSVS" )
    } else if( grepl( 'MontaneOak-Hickory', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "MontaneOak-Hickory.svs", package="rSVS" )
    } else if( grepl( 'PacificSilverFir-Hemlock', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "PacificSilverFir-Hemlock.svs", package="rSVS" )
    } else if( grepl( 'Redwood', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "Redwood.svs", package="rSVS" )
    } else if( grepl( 'SouthernPine', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "SouthernPine.svs", package="rSVS" )
    } else if( grepl( 'Spruce-Fir', Stand, ignore.case=TRUE ) ) {
        svsfile <- system.file( "bin", "Spruce-Fir.svs", package="rSVS" )
    } else {                                                                        # demo file not found, print message and return
        print( paste0( "Unknown demo file: '", Stand,"'. Please pick from: BottomlandHardwood, Douglas-fir, LodgepolePine, ",
                       "MixedConifer, MontaneOak-Hickory, PacificSilverFir-Hemlock, Redwood, SouthernPine, or Spruce-Fir" ) )
        return('SVS_Demo() exited.')
    }
    cmdline <- paste0( svsexe, " ", svsfile )                                       # create command line
    #print( cmdline )
    system( cmdline, invisible=FALSE )                                              # spawn SVS program
    setwd( SavedDir )                                                               # restore working directory to original
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
#' Live/Dead: live or l|dying|dead or d|stump or s
#' Status: standing or s|broken or b|brokentop|deadtop|down or d
#' Condition: Live:  1 or dominant or d|2 or codominant or c|3 or intermediate or i|4 or suppressed
#'            Dying: 1|2|3
#'            Dead:  1|2|3|4|5
#'
#' @param data compatible data frame or string containing filename with path (see details)
#' @param output what and were to product output (SVS | BITMAP | WEB | CSV )
#' @param clumped if TRUE generate clumped coordinates
#' @param random if TRUE generate random coordiantes
#' @param row if TRUE generate coordinates for rows (plantation)
#' @param uniform if TRUE generate uniform coordinates
#' @param randommess control "noise" of coordinates generated
#' @param clumpiness adjust clump strength
#' @param clumpratio adjust number/size of clumps
#' @param verbose turn on verbose output
#' @author James Mccarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' SVS( d )
#' SVS( '../MyFiles/Stand1.csv' )
#' SVS( '../MyFiles/Stand2.xlsx', Sheet='Sheet1' )
#' SVS( d, random=true )    # visualize stand in svs using random tree locations
#' SVS( d, row=true )       # visualize stand in svs using rows
#' @export
SVS <- function( data, sheet=FALSE, output='svs', clumped=FALSE, random=TRUE, row=FALSE, uniform=FALSE, randomness=NULL, clumpiness=NULL, clumpratio=NULL, verbose=FALSE ) {
    #print( data )
    if( typeof(data) == "character" ) {
        if( verbose ) print( paste0( "Data = \"character\"" ) )
        cmdline <- paste0( ".\\python38\\python.exe ", system.file( "python", "StandViz.py", package="rSVS" ), " -D -v ", data )
        if( verbose ) print( cmdline )
        system( cmdline, invisible=FALSE, wait=TRUE )
    } else if( typeof(data) == "list" ) {
        if( verbose ) print( paste0( "Data = \"list\", need to save data or pass through reticulate interface" ) )
        # if 
    } else {
        print( paste0( "Don't know how to handle data or type ", typeof(data) ) )
    }
    if( ! "reticulate" %in% .packages() ) print( paste0( "reticulate package NOT loaded" ) )
    if( ! "reticulate" %in% rownames(installed.packages()) ) print( paste0( "reticulate package NOT installed" ) )
    #svsexe <- system.file( "bin/svs", "winsvs.exe", package="rSVS" )
    #demosvsfile <- system.file( "bin", "SouthernPine.svs", package="rSVS" )
    #print( demosvsfile )
    #cmdline <- paste0( svsexe, " ", demosvsfile )
    #print( cmdline )
    #system( cmdline, invisible=FALSE )
    #cmdline <- paste0( ".\\python38\\python.exe -i .\\inst\\python\\StandViz.py -D -v -A")
    #cmdline <- paste0( ".\\python38\\python.exe ", system.file( "python", "StandViz.py", package="rSVS" ), " -D -v ", system.file( "bin", data, package="rSVS") )
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

