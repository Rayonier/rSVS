# Package Documentation
#
#' A packge for stand level visualization using the Stand Visualization System (SVS).
#'
#' This rSVS package provides and interface to perform SVS visualiations from R.
#'
#' The package includes the following functions:
#' \itemize{
#'     \item SVS()             - main function for performing visualziations
#'     \item SVS_Environment() - show reginal example visulizations
#'     \item SVS_Example()     - show reginal example visulizations
#'     \item SVS_Species()     - list known species
#'     \item FIA2NRCS()        - convert species codes from FIA # to NRCS code
#'     \item NRCS2FIA()        - convert species codes from NRCS code to FIA #
#' }
#'
#' This package includes a number of executable programs that will be run as part of the package.
#' This limits where the package can be hosted (e.g not on CRAN).
#'
#' NOTE: SVS is a Windows only program, therefor limiting this package to only work on Windows
#' computers.
#'
#' @docType package
#' @name rSVS
NULL

#' Check SVS environment and return path to components
#'
#' SVS_Enviroment() checks the package enviroment and alternatively will install a python distribution
#'
#' Possible package components to investigate include:
#' \itemize{
#'   \item All
#'   \item SVS
#'   \item Python
#'   \item BMP2PNG
#'   \item Zip
#' }
#'
#' When testing for one component the path to the relavant executable will be returned. When testing
#' for All paths to each component will be provided.
#'
#' When testing for Python the function will first test to see if a copy of Python is available on
#' the PATH. If there is not a system wide copy of Python available the function will check for a
#' package #' internal copy. The first time SVS_Environment('Python') is run, the user will be
#' prompted to allow the unzipping of the required python files (the package includes a zipped copy of
#' Python 3.8 that can be installed into the package). Subsequent calls to SVS_Enviroment('Python')
#' will located the "python.exe" in the package and return that path.
#'
#' @param component which part of the enviroment to check, default all
#' @param verbose echo status messages as environment is being examined
#' @param debug toggle to turn on extra output while function is running
#' @return path path to individual component returned and messages printed on console
#' @author Jim McCarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' SVS_Environment( 'python' )  # investigate Python environment for running backend code
#' SVS_Environment( 'svs' )     # make sure all SVS components are available to run program
#' @export
SVS_Environment <- function( component='all', verbose=FALSE, debug=FALSE ) {
    #MyComponents <- c( 'SVS', 'Python', 'BMP2PNG', 'Zip' )  # list of components that exist
    if ( component == 'all' ) {                                                 # test all components with recursive call
        SvsPath <- SVS_Environment( 'svs', verbose, debug )                     # call ourselved to get SVS path
        PyPath <- SVS_Environment( 'python', verbose, debug )                   # call ourselves to get Python path
        BmpPath <- SVS_Environment( 'bmp2png', verbose, debug )                 # call ourselves to get BMP2PNG path
        ZipPath <- SVS_Environment( 'zip', verbose, debug )                     # call ourselfes to get Zip path
        return( c(SvsPath, PyPath, BmpPath, ZipPath) )                          # return paths
    } else if( component == 'svs' ) {                                           # handle SVS component
        SvsExePath <- system.file( "bin/SVS", "winsvs.exe", package="rSVS" )    # check for winsvs.exe
        if( SvsExePath == "" ) {                                                # if file not found in package
            print( "Error in package!  Will not be able to visualize stands." ) # print error message
        } else {                                                                # else, check to make sure all required files are there
            if( verbose ) print( "winsvs.exe file located in package" )
			FILELIST <- c('DEFAULT.TRF', 'FIA.TRF', 'fvs2svs.dll', 'fvs2svs.hlp', 'NRCS.TRF', 'org2svs.dll', 'svslib.hlp',   # list of files
                         'svslib.ini', 'tbl2svs.dll', 'tbl2svs.exe', 'tbl2svs.hlp', 'winsvs.exe', 'winsvs.hlp', 'winsvs.ini')
			FileMissing <- FALSE                                                # start assuming no fils missing
			for( F in FILELIST ) {                                              # loop across files
				FileTest = system.file( "bin/SVS", F, package="rSVS" )          # look for file
				if( debug ) print( FileTest )                                   # if Debug, show process
				if( FileTest == "" ) FileMissing <- TRUE                        # if "" returned, file not in package
			}
			if( FileMissing ) print( "One or more files missing from SVS program filder!" ) # report error
			else if( verbose ) print( "SVS and support files are available." )  # esle, echo everything looks good
		}
        return( SvsExePath )                                                    # return path to winsvs.exe in package
	} else if( component == 'python' ) {                                        # handle Python component
		if( verbose ) print('Testing for Python')                               # message
        PyPath <- ""                                                            # default, no path
		SysPyPath <-  Sys.which( "python" )                                     # check for system Python on PATH
        IntPyPath <- system.file( "bin","python38/python.exe",package="rSVS" )  # check for internal python
        if( debug ) print( paste0( "SysPyPath=", SysPyPath, ", IntPyPath=", IntPyPath ) )   # if Debug, echo values
        if( SysPyPath == "" ) {                                                 # no python on PATH
            if( verbose ) print( "no Sys.which('python)' located" )
        } else {                                                                # python on PATH
            if( IntPyPath == "" ) {                                             # have system Python, but no internal
                PyPath = SysPyPath                                              # use system Python
                # should go additional testing to make sure we have all the bits we need
            } else {                                                            # have system and internal python,
                PyPath = IntPyPath                                              # use internal python
            }
            return( PyPath )                                                    # return path
        }
        if( IntPyPath == "" ) {                                                 # no bin/python38/python.exe
            if( verbose ) print( "No 'python38/python.exe' found")
            if( debug ) print( paste0( "SysPyPath=", SysPyPath, ", IntPyPath=", IntPyPath ) )
            PyZipPath <- system.file( "bin", "python38.zip", package="rSVS" )   # check for python38.zip file
            if( PyZipPath == "") {                                              # no python38.zip file found
                print( "No python38.zip in package, need python to run" )       # can't work, exit
                return()                                                        # exit with no path
            } else {                                                            # found zip, and extract
                Response <- readline( prompt=paste0( "No python found on system, but I can install package internal copy?  Y/N: " ) )
                if( Response == 'Y' ) {                                         # if use agrees
                    SaveWD = getwd()                                            # save current working directory
                    setwd( system.file("bin","",package="rSVS") )               # set working directory to bin folder in package
                    system( "unzip.exe python38.zip", invisible=TRUE )          # unzip python38.zip
                    setwd( SaveWD )                                             # restore working directory to orignal
                    IntPyPath <- system.file( "bin","python38/python.exe",package="rSVS" )  # confirm path to file
                }
            }
            if( debug ) print( paste0( "SysPyPath=", SysPyPath, ", IntPyPath=", IntPyPath ) )
            if( verbose ) print( paste0( "Python located at ", IntPyPath ) )
            return( IntPyPath )
        } else {
            if( verbose ) print( paste0( "Python located at ", IntPyPath ) )
            return( IntPyPath )
        }
        return("should never get here")
	} else if( component == 'bmp2png' ) {                                       # handle BMP2PNG component
        Bmp2PngExe <- system.file( "bin", "BMP2PNG.EXE", package="rSVS" )
        if( Bmp2PngExe == "" ) {
            print( "Error in package!  Will not be able to convert BMP files to PNG file for web page presentation of visualizations")
        } else {
			if( verbose ) print( 'BMP2PNG.EXE, used to convert bitmap files to web friendly PNG graphics files, is available.' )
		}
        return( Bmp2PngExe )
	} else if( component == 'zip' ) {                                           # handle Info-Zip component
        ZipExe <- system.file( "bin", "zip.exe", package="rSVS" )
        if( ZipExe == "" ) {
            print( "Error in package!  Will not be able to extract python38.zip if no system defined python exists." )
        } else {
			if( verbose ) print( 'Info-Zip zip.exe and unzip.exe are available.' )
		}
        return( ZipExe )
    }
}

#' Demonstrate Stand Visualiztion on several stand types
#'
#' Display one of several stand types using example SVS files included with package.
#'
#' The list of availabel stand type examples include:
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
#' @param Example Name of stand/stand type example to display
#' @author Jim McCarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' SVS_Demo( 'SouthernPine' )
#' SVS_Demo( 'Douglas-fir' )
#' SVS_Demo()                   # gives list of possible options
#' @export
SVS_Example <- function( Example=NULL ) {
    SavedDir <- getwd()                                                             	# get and save current working directory
    setwd( path.package("rSVS") )                                                   	# set working directory to package location
    svsexe <- SVS_Environment( "svs" )                	                                # get location of winsvs.exe
    if( is.null(Example) ) {                                                          	# if no stand type provided, print message and return
        print( paste0( "Please pick from: BottomlandHardwood, Douglas-fir, LodgepolePine, MixedConifer, ",
                       "MontaneOak-Hickory, PacificSilverFir-Hemlock, Redwood, SouthernPine, or Spruce-Fir" ) )
        return('SVS_Demo() exited.')
    } else if( grepl( 'BottomlandHardwood', Example, ignore.case=TRUE ) ) {            	# check, ignoring case to eliminate typos or case
        svsfile <- system.file( "extdata", "BottomlandHardwood.svs", package="rSVS" )   # get location of SVS file
    } else if( grepl( 'Douglas-fir', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "Douglas-fir.svs", package="rSVS" )
    } else if( grepl( 'LodgepolePine', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "LodgepolePine.svs", package="rSVS" )
    } else if( grepl( 'MixedConifer', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "MixedConifer.svs", package="rSVS" )
    } else if( grepl( 'MontaneOak-Hickory', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "MontaneOak-Hickory.svs", package="rSVS" )
    } else if( grepl( 'PacificSilverFir-Hemlock', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "PacificSilverFir-Hemlock.svs", package="rSVS" )
    } else if( grepl( 'Redwood', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "Redwood.svs", package="rSVS" )
    } else if( grepl( 'SouthernPine', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "SouthernPine.svs", package="rSVS" )
    } else if( grepl( 'Spruce-Fir', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "Spruce-Fir.svs", package="rSVS" )
    } else {                                                                        	# example file not found, print message and return
        print( paste0( "Unknown demo file: '", Example,"'. Please pick from: BottomlandHardwood, Douglas-fir, LodgepolePine, ",
                       "MixedConifer, MontaneOak-Hickory, PacificSilverFir-Hemlock, Redwood, SouthernPine, or Spruce-Fir" ) )
        return('SVS_Example() exited.')
    }
    cmdline <- paste0( svsexe, " ", svsfile )                                       	# create command line
    #print( cmdline )
    system( cmdline, invisible=FALSE )                                              	# spawn SVS program to display file
    setwd( SavedDir )                                                               	# restore working directory to original
    return( 'SVS_Example() existed.' )
}

# hidden function to conver FMD plot data in R to .csv file
FMD2CSV <- function( data ) {
    if( ! file.exists('svsfiles') ) dir.create( 'svsfiles' )                            # if svsfiles does not exist, create it
    CSVFilename <- paste0( "svsfiles/FMD_data.csv"  )
    #tl <- data
    tl <- data[,c(3,9:12,14:21)]
    write.csv( tl, CSVFilename, row.names=FALSE )
    return( CSVFilename )
}

# hidden function to convert StandObject in R to .csv file
StandObject2CSV <- function( data ) {
    if( ! file.exists( 'svsfiles' ) ) dir.create( 'svsfiles' )                          # if svsfiles does not exist, create it
    CSVFilename <- paste0( "svsfiles/", data$header$standid, "_", data$header$ysp, ".csv" ) # format filename from $header$standid and $header$ysp
    #print( paste0( "CSVFilename=", CSVFilename ) )
    #tr <- cbind( stand=data$header$standid, age=data$header$ysp, data$treelist )[,c(-9)]    # extract treelist to new dictionary with standid and ysp included
    tr <- cbind( data$treelist[,c(2,4,5,6,3)], crad=0, status=1, pc=0, cc=0 )           # extract treelist to new dictionary with standid and ysp included
    tr <- tr[,c(1,2,3,4,6,7,8,9,5)]                                                     # extract and re-order columns we want
    write.csv( tr, CSVFilename, row.names=FALSE )                                       # write .csv file
    return( CSVFilename )                                                               # return filename
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
    if( verbose ) print( paste0( "class(data) = ", class(data) ) )
    PyExePath <- SVS_Environment('python')
    #PyExePath <- ".\\python38\\python.exe"
    if( class(data) == "character" ) {
        if( verbose ) print( paste0( "Data = \"character\"" ) )
        cmdline <- paste0( PyExePath, " ", system.file( "python", "StandViz.py", package="rSVS" ), " -D -v ", data )
        if( verbose ) print( cmdline )
        RetValue <- system( cmdline, invisible=FALSE, wait=TRUE )
        if( RetValue == 0 ) return( "SVS() completed" )
        else print( paste0( "Error running command!  Error = ", RetValue, " for command: ", cmdline ) )
    } else if( class(data) == "list" ) {
        if( (attributes(data)$names[1]=="header") & (attributes(data)$names[2]=="treelist"))  {
            if( verbose ) print( "Pretty sure we have a ray.pacific.tlm/organon stand object" )
            CsvFile <- StandObject2CSV(data)
            cmdline <- paste0( PyExePath, " \"", system.file( "python", "StandViz.py", package="rSVS" ), "\" -D -v ", CsvFile )
            if( verbose ) print( cmdline )
            RetValue <- system( cmdline, invisible=FALSE, wait=TRUE )
            if( RetValue == 0 ) return( "SVS() completed" )
            else print( paste0( "Error running command!  Error = ", RetValue, " for command: ", cmdline ) )
        } else {
            print( paste0( "Not sure what object type we have here: ", attribute(data)$names, str(data) ) )
        }
    } else if( class(data) == "data.frame" ) {
        if( (attributes(data)$names[1]=="DataSource") & (attributes(data)$names[3]=="PlotKey") ) {       # have FMD treelist for plots
            # attributes(data)$class = "data.frame"
            if( verbose ) print( "Pretty sure we have an FMD tree data frame")
            #print( attributes(data)$names )
            print( paste0( "Will create visualizations for: ", paste(unique(data$PlotKey),collapse=' ') ) )
            CsvFile <- FMD2CSV(data)
            cmdline <- paste0( PyExePath, " \"", system.file( "python", "StandViz.py", package="rSVS" ), "\" -D -v ", CsvFile )
            if( verbose ) print( cmdline )
            RetValue <- system( cmdline, invisible=FALSE, wait=TRUE )
            if( RetValue == 0 ) return( "SVS() completed" )
            else print( paste0( "Error running command!  Error = ", RetValue, " for command: ", cmdline ) )
        } else {
            print( "Some unknown data.frame format:" )
            print(str(data))
            print(attributes(data)$names)
        }
    } else {
        print( paste0( "Don't know how to handle this type of data: ", typeof(data) ) )
        print(str(data))
    }
    #if( ! "reticulate" %in% .packages() ) if( verbose ) print( paste0( "reticulate package NOT loaded" ) )
    #if( ! "reticulate" %in% rownames(installed.packages()) ) if( verbose ) print( paste0( "reticulate package NOT installed" ) )
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


#' List species codes
#'
#' List known species codes.
#'
#' This function simply reads and returns the list of known species codes distributed with the rSVS
#' package (rSVS_Species.csv). This file is also used as the basis for species translation by the
#' FIA2NRCS() and #' NRCS2FIA() funtions.
#'
#' The rSVS package supports FIA number and NRCS alphabetic codes. The displayed table will include
#' FIA, NRCS, Genus, Species, Common, and Comment, FIA.trf, NRCS.trf, FVS.trf, FVSSpCode, FVS.East,
#' and FVS.West.
#'
#' The FIA.trf and NRCS.trf columns should list the species that currently exist in each SVS tree
#' form file (NOTE: this species list should be synchronized with FIA.trf and NRCS.trf, but this is
#' not guaranteed).
#'
#' The FVS.trf, FVSSpCode, FVS.East, and FVS.West columns are used for species audits and keeping
#' track of the source of tree form definitions (largely from FIA variant tree form files from the
#' original SVS distribution) before the tree form definitions were exanded to handle more live and
#' dead classes.
#'
#' @return DataFrame with species names and codes
#' @examples
#' Species <- SVS_Species()            # store species list to Species
#' names(SVS_Species)                  # column names in returned species list
#' head(SVS_Species())                 # first 6 species records
#' length(SVS_Species()$FIA)           # number of species records
#' length(unique(SVS_Species()$FIA))   # number of FIA # known
#' sort(unique(SVS_Species()$FIA))     # list of FIA #'s known
#' length(unique(SVS_Species()$NRCS))  # number of NRCS codes known
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
#' @author Jim McCarter \email{jim.mccarter@rayonier.com}
#' @examples
#' FIA2NRCS( MyData )
#' @export
FIA2NRCS <- function( DataFrame ) {
    print( "FIA2NRCS() not implemented yet!")
}

#' Convert species codes from NRCS code to FIA number
#'
#' Details
#'
#' @param MyData DataFrame with tree information
#' @author Jim McCarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' NRCS2FIA( MyData )
#' @export
NRCS2FIA <- function( DataFrame ) {
    print( "NRCS2FIA() not implemented yet!")
}

# You can learn more about package authoring with RStudio at:
#
#   http://r-pkgs.had.co.nz/
#
# Some useful keyboard shortcuts for package authoring:
#
#   Build and Reload Package:  'Ctrl + Shift + B'
#   Check Package:             'Ctrl + Shift + E'
#   Test Package:              'Ctrl + Shift + T'

