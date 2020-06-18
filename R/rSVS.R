# Package Documentation
#
#' A package for stand level visualization using the Stand Visualization System (SVS, Robert J. McGaughey, USDA Forest Service, PNW Research Station).
#'
#' This rSVS package provides an interface to perform SVS visualiations from R.
#'
#' The package includes the following functions:
#' \itemize{
#'     \item SVS()             - main function for performing visualziations
#'     \item SVS_Environment() - check package environment and returns path to components
#'     \item SVS_Example()     - show reginal example visualizations
#'     \item SVS_ExampleData() - generate stand data for visualizations
#'     \item SVS_Species()     - list known species
#'     \item svsfiles_clean()  - clean out svsfiles folder containing temporary files for visualizations
#'     \item FIA2NRCS()        - convert species codes from FIA # to NRCS code
#'     \item NRCS2FIA()        - convert species codes from NRCS code to FIA #
#' }
#'
#' The package supports either species identified by FIA # or NRCS Plants Database code. The default
#' behavior can be configured by setting either hidden variable ".UseFIA <- TRUE" or ".UseNRCS <- TRUE".
#' NOTE: Depending on the source data, setting the incorrect treeform file will cause trees to be
#' displayed as an Unknown tree species (e.g. red)
#'
#' The SVS() function supports multiple data format types, including:
#' \itemize{
#'     \item FMD data frame - a data frame with at least TreeKey, Species, MeasDate, Status, DBH, Height, CrownRatio, TPA
#'     \item LMS Object - a list with stand, measurement, and treelist slots
#'     \item StandObject - a list with header, treelist, and management slots
#'     \item SVScsv - a data frame or .csv file containing: Species, PlantID, PlntClass, CrwnClass, TreeStat, DBH, Height, LAng, FAng, EndDia, CRad1, CRat1,
#'                    CRad2, CRat2, CRad3, CRat3, CRad4, CRat4, TPA, MarkCode, X, Y, Z
#'     \item StandViz - a data frame or .csv file containing: Stand, Year/Age, Species, TreeNo, Live/Dead, Status, Condition, DBH, Height, CrownRatio, CrownRadius, TPA
#'     \item StandVizExtended - a data frame or .csv file containing: Stand, Year/Age, Species, TreeNo, Live/Dead, Status, Condition, DBH, Height, CrownRatio,
#'                              CrownRadius, TPA, BrokenHt, Offset, Bearing, Lean, RootWad, X, Y
#'     \item TBL2SVS data frame - a data frame or .csv file containing: Species, DBH, Height, CrownRatio, CrownRadius, Status, PlantClass, CrownClass, TPA
#' }
#'
#' FMD data frame details:
#'
#' A FMD data frame contains PlotName in column 3, TreeNo, Species, MeasDate, MeasAge in columns 9-12, and Status, Condition, Damage,
#' Screen, DBH, Height, CrownRatio in columns 14-21.
#'
#' LMS object details:
#'
#' A LMS object is a list with $stand, $measurement, and $treelist slots.  The $stand slot must contain $STANDNAME, the $measurement must
#' contain $MEASDATA where year can be extracted from the first 4 character positions.  The $treelist slot must contain tree, species,
#' tpa, dbh, height, and cr columns in that order.
#'
#' StandObject details:
#'
#' A StandObject is a list that contains at a minimum $header and #treelist slots.  The $header slot must contain $standid and
#' $ysp (years since planting) fields.  The $treelist slopt must contain treeid, species, tpa, dbh, height, and cr columns in
#' that order.
#'
#' StandViz details:
#'
#' A data frame or .csv file containing: Stand, Year/Age, Species, TreeNo, Live/Dead, Status, Condition, DBH, Height, CrownRatio, CrownRadius, TPA
#'
#' StandVizExtended details:
#'
#' A data frame or .csv file containing: Stand, Year/Age, Species, TreeNo, Live/Dead, Status, Condition, DBH, Height, CrownRatio, CrownRadius, TPA,
#' BrokenHt, Offset, Bearing, Lean, RootWad, X, Y.
#'
#' This format can be used for visualizing mapped stand data and supports a variety of enhancedments that allow for improved visualizations of
#' forest health conditions, including broken trees (BrokenHt, Offset, Bearing), dead top trees(Condition, BrokenHt), blown down trees (RootWad).
#' If TPA > 1 then coordinates are generated for the stand.  To provide a mapped stand all TPA values must be 1.0 and X,Y locations filled in.
#'
#' TBL2SVS details:
#'
#' A data frame or .csv file containing: Species, DBH, Height, CrownRatio, CrownRadius, Status, PlantClass, CrownClass, TPA [, X, Y, MarkCode, FAng, SDia]
#'
#' Status Codes:
#' \tabular{lll}{
#'   \tab 0 or 10 \tab Plant is cut, has branches, on ground \cr
#'   \tab 1 or 11 \tab plant is standing \cr
#'   \tab 2 or 12 \tab stump \cr
#'   \tab 3 or 13 \tab plant is cut, has no branches, on ground \cr
#' }
#'
#' The following example input files are provided:
#' \tabular{lcl}{
#'   - BottomlandHardwood.svs   \tab - \tab SVS format file of mixed bottomland hardwood from southeast \cr
#'   - Douglas-fir.svs          \tab - \tab SVS format file of Douglas-fir from Pacfic Northwest \cr
#'   - LodgepolePine.svs        \tab - \tab SVS format file of Lodgepole pine from Rocky Mountains \cr
#'   - MixedConifer.svs         \tab - \tab SVS format file of Mixed conifer stand from California \cr
#'   - MontaneOak-Hickory.svs   \tab - \tab SVS format file of Oak/Hickory forest \cr
#'   - SilverFir-Hemlock.svs    \tab - \tab SVS format file of Pacific silver fir/hemlock stand from Pacfic Northest \cr
#'   - Redwood.svs              \tab - \tab SVS format file of Redwood stand from Calinornia \cr
#'   - SouthernPine.svs         \tab - \tab SVS format file of Southern pine stand from southeast \cr
#'   - SouthernPine.csv         \tab - \tab Same as above, but in SVScsv format \cr
#'   - SouthernPine2.csv        \tab - \tab Abreviated version of SouthernPine without 4 crown dimensions, but including X,Y coordinates \cr
#'   - SouthernPine3.csv        \tab - \tab Abreviated version of SouthernPine without X, Y coordinates \cr
#'   - Spruce-Fir.csv           \tab - \tab SVS format file of Spruce/Fir from Rocky Mountains \cr
#' }
#'
#' NOTE: This package includes a number of executable programs that will be run as part of the package,
#' limiting where the package can be hosted (e.g NOT on CRAN).
#'
#' The package understands a number of data formats and can create visualizations for single or multiple
#' stands and years based on what information is contained in the specific data format. Files for
#' visualizations are created in the svsfiles folder in the current working directory for R. These
#' temporary files (*.asc, *.bmp, *.csv, *.png, *.SVS, *.opt) are intermediate files used for creating
#' the visualizations. This folder can be cleaned out using the \strong{svsfiles_clean()} function.
#'
#' NOTE: SVS is a Windows only program, therefor limiting this package to only work on Windows
#' computers.
#'
#' @docType package
#' @name rSVS-package
NULL

# data conversion functions
FMD2CSV <- function( data ) {                                                           # hidden function to convert FMD plot data in R to .csv file
    if( ! file.exists('svsfiles') ) dir.create( 'svsfiles' )                            # if svsfiles does not exist, create it
    CSVFilename <- paste0( "svsfiles/FMD_data.csv"  )                                   # create filename
    tl <- data[,c(3,9:11,14:16,18:21)]                                                        # select columns
    write.csv( tl, CSVFilename, row.names=FALSE )                                       # write .csv file
    return( CSVFilename )                                                               # return filename written
}

LMSObject2CSV <- function( data ) {                                                     # hidden function to convert LMS data in R to .csv file
    if( ! file.exists( 'svsfiles' ) ) dir.create( 'svsfiles' )                          # if svsfiles does not exist, create it
    CSVFilename <- paste0( "svsfiles/", data$stand$STANDNAME, "_", substr(data$measurement$MEASDATE,1,4), ".csv" ) # format filename from $header$standid and $header$ysp
    year <- substr(data$measurement$MEASDATE,1,4)
    tr <- cbind( data$treelist[,c(3,1,4,7,11,12,13)], year=year, cr=0, status=1, pc=0, cc=0 )           # extract treelist to new dictionary with standid and ysp included
    #print(head(tr))
    #Stand,ObjectID,Species,TreeQuality,QDBH,TPA,Height,Year,CR,Status,PC,CC
    write.csv( tr, CSVFilename, row.names=FALSE )                                       # write .csv file
    return( CSVFilename )                                                               # return filename written
}

StandObject2CSV <- function( data ) {                                                   # hidden function to convert StandObject in R to .csv file
    if( ! file.exists( 'svsfiles' ) ) dir.create( 'svsfiles' )                          # if svsfiles does not exist, create it
    CSVFilename <- paste0( "svsfiles/", data$header$standid, "_", data$header$ysp, ".csv" ) # format filename from $header$standid and $header$ysp
    year <- data$header$ysp
    tr <- cbind( data$treelist[,c(2,4,5,6,3)], crad=0, status=1, pc=0, cc=0 )           # extract treelist to new dictionary with standid and ysp included
    if( nrow(data$cut.trees) > 0 ) {                                                    # if we have records in cut-trees (need to validate year)
      tt <- cbind( data$cut.trees[,c(3,5,6,7,4)], crad=0, status=3, pc=0, cc=0 )
      tr <- rbind( tr, tt )
    }
    tr <- tr[,c(1,2,3,4,6,7,8,9,5)]                                                     # extract and re-order columns we want
    write.csv( tr, CSVFilename, row.names=FALSE )                                       # write .csv file
    return( CSVFilename )                                                               # return filename written
}

SVScsvObject2CSV <- function( data, name=deparse(substitute(data)) ) {
    if( ! file.exists( 'svsfiles' ) ) dir.create( 'svsfiles' )
    CSVFilename <- paste0( "svsfiles/", deparse(substitute(data)), ".csv" )
    print( paste0( "CSVFilename = ", CSVFilename, ", name= ", name ) )
    write.csv( data, CSVFilename, row.names=FALSE )
    return( CSVFilename )
}

# turn comment into #' to export and make available to user
#' @export
Detect_DataType <- function( data, verbose=FALSE ) {                                    # hidden function to detect data type of object or file
    DataType <- NULL                                                                    # start with data type not known
    if( verbose ) print( paste0( "class(data) = ", class(data) ) )                      # echo what type of data we have
    if( class(data) == "character" ) {                                                  # have a string which is a filename
        if( verbose ) print( paste0( "Data = \"character\"" ) )                         # echo if verbose
        if( grepl( ".svs", tolower(data) ) ) DataType <- 'SVSFile'                               # have a .svs file
        else if( grepl( '.csv', tolower(data) ) ) {                                              # have a .csv file
            if( file.exists( data ) ) {                                                 # make sure file exists
                tf <- read.csv( data )                                                  # read file to check format
                if( length(attributes(tf)$names) < 14 ) DataType <- 'TBL2SVSFile'       # TBL2SVS format file
                else if( attributes(tf)$names[12]=="CRat1") DataType <- 'SVScsvFile'    # SVScsv format file
                else if( (length(attributes(tf)$names)>=19) & (attributes(tf)$names[17]=="RootWad") ) DataType <- 'StandVizExtendedFile'
                else if( (length(attributes(tf)$names)>=14) & (attributes(tf)$names[2]=="Year.Age") ) DataType <- 'StandVizFile'
                else DataType <- 'CSVFile'                                              # some other .csv format
            }
        }
        if( ! file.exists( data ) ) print( paste0( "Error: File '", data, "' does not exist!" ) )   # warn if file does not exist
    } else if( class(data) == "list" ) {                                                # have a list, now test of what type of data
        if( (attributes(data)$names[1]=="header") & (attributes(data)$names[2]=="treelist"))  {     # should be organon/cipsanon/ryn.c2g stand object
            if( verbose ) print( "Detected organon/cips/plc stand object" )             # echo verbose
            DataType <- 'StandObject'                                                   # set DataType to StandObject type
        } else if( (attributes(data)$names[1]=="stand") & (attributes(data)$names[2]=="measurement") & (attributes(data)$names[3]=="treelist") )  {
            if( verbose ) print( "Detected LMS stand object" )                          # echo if verbose
            DataType <- 'LMSObject'                                                     # set DataType to LMSObject type
        } else {
            print( paste0( "Not sure what object type we have here: ", attributes(data)$names, str(data) ) )    # don't know format
        }
    } else if( class(data) == "data.frame" ) {                                          # have a data frame
        if( (attributes(data)$names[1]=="DataSource") & (attributes(data)$names[3]=="PlotKey") ) {       # have FMD treelist for plots
            if( verbose ) print( "Detected FMD tree data frame")                        # if verbose, echo type detected
            DataType <- 'FMDObject'                                                     # set DataType to FMDObject
        } else if( length(attributes(data)$names)<14 ) { DataType <- 'TBL2SVSObject'    # TBL2SVS format object
        } else if( (length(attributes(data)$names)>=19) & (attributes(data)$names[17]=='RootWad') ) { DataType <- 'StandVizExtendedObject'
        } else if( (length(attributes(data)$names)>=14) & (attributes(data)$names[2]=='Year.Age') ) { DataType <- 'StandVizObject'
        } else if( attributes(data)$names[12]=='CRat1' ) { DataType <- 'SVScsvObject'   # SVScsv format object
        } else {
            print( "Some unknown data.frame format:" )                                  # don't know this format
            print(str(data))
            print(attributes(data)$names)
        }
    }
    return( DataType )                                                                  # return detected data type
}

#' Visualize stand using the Stand Visualization System (SVS)
#'
#' The SVS() function will create stand level visualizations from data frames and files containing appropriate information.  The
#' function has the abillity to generate coordinates if they are not provided. Additionally missing information (height, crown ratio, crown width) values will be
#' "dubbed" if missing.
#'
#' StandViz internal format:
#' stand, year, species, treeno, x, y, dbh, height, crownratio, crownradius, tpa, live, status, condition, svsstatus, bearing, brokenht, brokenoffset,
#' dmr, leanangle, rootwad
#'
#' rSVS data frame format:
#' stand, year, treeno, species, dbh, height, crownratio, crownradius, tpa, x, y, live, status, condition, (svsstatus, brokenht, brokenoffset, bearing, dmr,
#' leanangle, rootwad)
#'
#' Live/Dead: live or l|dying|dead or d|stump or s
#'
#' Status: standing or s|broken or b|brokentop|deadtop|down or d
#'
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
SVS <- function( data, sheet=FALSE, output='svs', clumped=FALSE, random=TRUE, row=FALSE, uniform=FALSE, randomness=NULL, clumpiness=NULL, clumpratio=NULL,
                 verbose=FALSE, debug=FALSE ) {
    if( exists(".Development") ) PyExePath <- ".\\python38\\python.exe"                 # if under development use local copy of python
    else PyExePath <- SVS_Environment('python')                                         # else test for and optionally install package copy of python
    DataType <- Detect_DataType( data, verbose )
    if( debug ) print( paste0( "DataType=", DataType ))
    StandVizOpt <- " "                                                                  # start with StandViz.py options empty
    if( verbose ) StandVizOpt <- paste0( StandVizOpt, " -v" )                           # if verbose=TRUE, append -v command line option
    if( debug ) StandVizOpt <- paste0( StandVizOpt, " -D" )                             # if debug=TRUE, append -D
    if( clumped ) StandVizOpt <- paste0( StandVizOpt, " -gc" )                          # if clumped=TRUE, add command line option for clumped coordinates
    else if( row ) StandVizOpt <- paste0( StandVizOpt, " -gf" )                         # if row=TRUE, add command line option for fixed coordinates
    else if( uniform ) StandVizOpt <- paste0( StandVizOpt, " -gu" )                     # if uniform=TRUE, add command line option for uniform coordinates
    else StandVizOpt <- paste0( StandVizOpt, " -gr" )                                   # esle add command line option for random coordinates
    if( !is.null(randomness) ) StandVizOpt <- paste0( StandVizOpt, " -rf ", randomness )    # add randomness factor if specified
    if( !is.null(clumpiness) ) StandVizOpt <- paste0( StandVizOpt, " -cf ", clumpiness )    # add clumpiness factor if specified
    if( !is.null(clumpratio) ) StandVizOpt <- paste0( StandVizOpt, " -cr ", clumpratio )    # add clumpratio option if specified
    if( exists(".UseNRCS") ) if( .UseNRCS ) StandVizOpt <- paste0( StandVizOpt, " -N" ) # tell StandViz.py to use NRCS treeform file
    svcmdline <- paste0( PyExePath, ' "', system.file( "python", "StandViz.py", package="rSVS" ), '" ', StandVizOpt )    # create path to StandViz.py program
    if( debug ) print(svcmdline)
    if( DataType %in% c('CSVFile','SVSFile', 'SVScsvFile', 'StandVizFile', 'StandVizExtendedFile', 'TBL2SVSFile') ) { # have a string which is a filename
        cmdline <- paste0( svcmdline, " ", data )                                            # add data to command line
    } else if( DataType == "StandObject" ) {                                            # have a organon/cips/c2g stand object
        CsvFile <- StandObject2CSV(data)                                                # convert data to .csv file
        cmdline <- paste0( svcmdline, " ", CsvFile )                                         # add CsvFile to command line
    } else if( DataType == "LMSObject" ) {
        CsvFile <- LMSObject2CSV(data)                                                  # convert data from LMS object to .csv file
        cmdline <- paste0( svcmdline, " ", CsvFile )                                     # add CsvFile to command line
    } else if( DataType=="FMDObject" ) {
        # need to figure out how many visualizations will be created.  Unique(PlotKey and 
        nViz <- length( unique( paste0( data$PlotKey,"_",data$MeasDate) ) )
        print( paste0( "Will create visualizations for: ", paste(unique(data$PlotKey),collapse=' ') ) )
        if( nViz > 3 ) {
            ans <- readline( prompt=paste0( "You are requestion ", nViz, " visualizations.  Proceed (Y/N)? " ) )
            if( !(tolower(ans) %in% c('y', 'yes')) ) return( "Visualizations Cancelled!")
        }
        CsvFile <- FMD2CSV(data)
        cmdline <- paste0( svcmdline, " ", CsvFile )                                     # add CsvFile to command line
    } else if( DataType=="StandVizObject" ) {
        CsvFile <- paste0( "svsfiles/", bquote(data), ".csv" ) # format filename from object name
        write.csv( data, CsvFile, row.names=FALSE )                                       # write .csv file
        #cmdline <- paste0( PyExePath, " \"", system.file( "python", "StandViz.py", package="rSVS" ), "\"", StandVizOpt, CsvFile )
        cmdline <- paste0( svcmdline, " ", CsvFile )                                     # add CsvFile to command line
        #if( verbose ) print( paste0( "cmdline: ", cmdline )  )
        #RetValue <- system( cmdline, invisible=FALSE, wait=TRUE )
        #if( RetValue == 0 ) return( "SVS() completed" )
        #else print( paste0( "Error running command!  Error = ", RetValue, " for command: ", cmdline ) )
    } else if( DataType=="SVScsvObject" ) {
        CsvFile <- SVScsvObject2CSV( data )
        print( paste0( "SVScsvObject: CSVFile = ", CsvFile ) )
        cmdline <- paste0( svcmdline, " ", CsvFile )
    } else if( DataType=="TBL2SVSObject" ) {
        print( "Processing TBLS2SVSObject..." )
        CsvFile <- paste0( "svsfiles/", deparse(substitute(data)), ".csv" ) # format filename from object name
        write.csv( data, CsvFile, row.names=FALSE )
        #cmdline <- paste0( PyExePath, " ", system.file( "python", "StandViz.py", package="rSVS" ), StandVizOpt, CsvFilename )
        cmdline <- paste0( svcmdline, " ", CsvFile )                                     # add CsvFile to command line
        #if( verbose ) print( paste0( "cmdline: ", cmdline )  )
        #RetValue <- system( cmdline, invisible=FALSE, wait=TRUE )
        #if( RetValue == 0 ) return( "SVS() completed" )
        #else print( paste0( "Error running command!  Error = ", RetValue, " for command: ", cmdline ) )
    } else {
        print( paste0( "Don't know how to handle this type of data: ", typeof(data) ) )
        print(str(data))
        return
    }
    if( verbose ) print( paste0( "cmdline: ", cmdline )  )
    RetValue <- system( cmdline, invisible=FALSE, wait=TRUE )                       # execute and save return value
    if( RetValue == 0 ) return( "SVS() completed" )                                 # return success
    else print( paste0( "Error running command!  Error = ", RetValue, " for command: ", cmdline ) ) # return error number and commmand line that failed
    #if( ! "reticulate" %in% .packages() ) if( verbose ) print( paste0( "reticulate package NOT loaded" ) )
    #if( ! "reticulate" %in% rownames(installed.packages()) ) if( verbose ) print( paste0( "reticulate package NOT installed" ) )
    # if reticulate
    # library(reticulate)
    # SVS <- import_from_path( "StandViz", path="inst/python" )
}

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
#'   \item Defaults
#' }
#'
#' When testing for one component, the path to the relavant executable will be returned. When testing
#' for All paths, a path to each component will be provided as a list.
#'
#' When testing for Python the function will first test to see if a copy of Python is available on
#' the system PATH. If there is no a system wide copy of Python available the function will check for a
#' package internal copy. The first time SVS_Environment('Python') is run, the user will be
#' prompted to allow the installation (unzipping) of the required python files (the package includes a zipped copy of
#' Python 3.8 that can be installed into the package). Subsequent calls to SVS_Enviroment('Python')
#' will located "python.exe" in the package and return the path.
#'
#' @param component which part of the enviroment to check, default all
#' @param verbose echo status messages as environment is being examined
#' @param debug toggle to turn on extra output while function is running
#' @return path of primary executable for the individual component returned and messages printed on console
#' @author Jim McCarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' SVS_Environment( 'python' )  # investigate Python environment for running backend code
#' SVS_Environment( 'svs' )     # make sure all SVS components are available to run program
#' @export
SVS_Environment <- function( component='all', verbose=FALSE, debug=FALSE ) {
    if ( tolower(component)=='all' ) {                                          # test all components with recursive call
        SvsPath <- SVS_Environment( 'svs', verbose, debug )                     # call ourselved to get SVS path
        PyPath <- SVS_Environment( 'python', verbose, debug )                   # call ourselves to get Python path
        BmpPath <- SVS_Environment( 'bmp2png', verbose, debug )                 # call ourselves to get BMP2PNG path
        ZipPath <- SVS_Environment( 'zip', verbose, debug )                     # call ourselfes to get Zip path
        SVS_Environment( 'defaults', verbose, debug )                           # call ourselfes to check defaults defined by hidden variables
        return( c(SvsPath, PyPath, BmpPath, ZipPath) )                          # return paths
    } else if( tolower(component)=='svs' ) {                                    # handle SVS component
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
	} else if( tolower(component)=='python' ) {                                 # handle Python component
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
                # should do additional testing to make sure we have all the bits we need
            } else {                                                            # have system and internal python,
                PyPath = IntPyPath                                              # use internal python
            }
            return( PyPath )                                                    # return selected path
        }
        if( IntPyPath == "" ) {                                                 # no bin/python38/python.exe
            if( verbose ) print( "No 'python38/python.exe' found")              # echo message if verbose
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
            if( debug ) print( paste0( "SysPyPath=", SysPyPath, ", IntPyPath=", IntPyPath ) )   # echo debug information
            if( verbose ) print( paste0( "Python located at ", IntPyPath ) )    # echo which python if verbose
            return( IntPyPath )                                                 # return package internal python path
        } else {
            if( verbose ) print( paste0( "Python located at ", IntPyPath ) )    # echo which python if verbose
            return( IntPyPath )                                                 # return package internal python path
        }
        return("should never get here")                                         # should never fall through to this line
	} else if( tolower(component)=='bmp2png' ) {                                # handle BMP2PNG component
        Bmp2PngExe <- system.file( "bin", "BMP2PNG.EXE", package="rSVS" )       # get path to BMP2PNG.exe
        if( Bmp2PngExe == "" ) {
            print( "Error in package!  Will not be able to convert BMP files to PNG file for web page presentation of visualizations")
        } else {
			if( verbose ) print( 'BMP2PNG.EXE, used to convert bitmap files to web friendly PNG graphics files, is available.' )
		}
        return( Bmp2PngExe )                                                    # return path to executable
	} else if( tolower(component)=='zip' ) {                                    # handle Info-Zip component
        ZipExe <- system.file( "bin", "zip.exe", package="rSVS" )               # get package internal path to zip.exe
        if( ZipExe == "" ) {
            print( "Error in package!  Will not be able to extract python38.zip if no system defined python exists." )
        } else {
			if( verbose ) print( 'Info-Zip zip.exe and unzip.exe are available.' )  # echo if verbose
		}
        return( ZipExe )                                                        # return path to zip.exe
    } else if( tolower(component)=='defaults' ) {                               # check for default behaviors set by hidden variables
        if( exists( ".UseFIA" ) ) print( paste0( ".UseFIA = ", .UseFIA ) )
        if( exists( ".UseNRCS" ) ) print( paste0( ".UseNRCS = ", .UseNRCS ) )
        if( exists( ".Development" ) ) print( paste0( ".Development = ", .Development ) )
    }
}

#' Demonstrate Stand Visualiztion on several stand types
#'
#' Display one of several stand types using example SVS files included with package.
#'
#' The list of available stand types include:
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
#' @param Example Name of stand/stand type to display
#' @author Jim McCarter \email{jim.mccarter@@rayonier.com}
#' @examples
#' SVS_Example( 'SouthernPine' )
#' SVS_Example( 'Douglas-fir' )
#' SVS_Example()                   # gives list of possible options
#' @export
SVS_Example <- function( Example=NULL ) {
    SavedDir <- getwd()                                                             	# get and save current working directory
    setwd( path.package("rSVS") )                                                   	# set working directory to package location
    svsexe <- SVS_Environment( "svs" )                	                                # get location of winsvs.exe
    if( is.null(Example) ) {                                                          	# if no stand type provided, print message and return
        print( paste0( "Please pick from: BottomlandHardwood, Douglas-fir, LodgepolePine, MixedConifer, ",
                       "MontaneOak-Hickory, SilverFir-Hemlock, Redwood, SouthernPine, or Spruce-Fir" ) )
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
    } else if( grepl( 'SilverFir-Hemlock', Example, ignore.case=TRUE ) ) {
        svsfile <- system.file( "extdata", "SilverFir-Hemlock.svs", package="rSVS" )
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
    return( 'SVS_Example() existed.' )                                                  # echo that SVS has exited
}

# hidden treelist function for use by SVS_ExampleData()
# @export
treelist <- function( species, dbh, tpa, scale=4, shape=2, n=30, dmin=0.0001, dmax=100, incr=0.1  ) {
    # exponential decrease (scale=2.5, shape=1); left skew(scale=4, shape=2); normal(scale=10, shape=3.6); right skew(scale=15, shape=10)
    d = data.frame( scale, shape, species, dbh, tpa)                     # create data frame of inputs
    tr = list()                                                     # create empty list for the returned treelist
    nSpp = nrow(d)                                                  # get number of species entered
    for( i in 1:nrow(d) ) {                                         # loop across species
        z = d[i,]                                                   # get row of dictionary
        #print( paste0( "dmin=", dmin, ", dmax=", dmax, ", incr=", incr) )
        di = seq( dmin, dmax, by=incr )                             # get initial sequence for distribution
        pdf = dweibull( di, shape=z$shape, scale=z$scale )          # generate pdf for 2 parameter weibull
        k = which(pdf > 0.01 )                                      # filter to values > 0.01
        dlims = range(di[k])                                        # get range of values with prob > 0.01
        #print( paste0( "dbh=", z$dbh, ", dlims[1]=", dlims[1], ", dlims[2]=", dlims[2] ) )
        di = seq( z$dbh-mean(dlims)-dlims[1], z$dbh+mean(dlims)-dlims[1], length.out=n/nSpp )           # create new sequence of diamters with prob > 0.01
        pdf = dweibull(di, shape=z$shape, scale=z$scale )           # create new pdf across diameter range
        o = data.frame( pdf, dbh=round(di,2) )                      # create data frame with pdf and diameters
        o$species = z$species                                       # add appropriate species
        o$wi = o$pdf / sum(o$pdf)                                   # add weight (proportion of pdf)
        o$tpa = z$tpa * o$wi                                        # compute tpa from weight
        tr[[i]] <- o                                                # add to treelist
    }
    return( do.call(rbind, tr) )                                    # return combined treelist
}

# Need function to generate tree data for visulizations:
#
# SVS_Generate( (DF, 12.5, 75.0, 100), (WH, 7.8, 68, 150), Random=TRUE )
#

#' Create SVS example data
#'
#' Create different kinds of example data understood by this package.
#'
#' This function will return data data frame in one of the following formats:
#' \itemize{
#   \item StandObject format
#'   \item StandViz format
#'   \item StandVizExtended format
#'   \item TBL2SVS format
#' }
#'
#' @param datatype Type of example data to create
#' @param species individual or list of species for stand
#' @param dbh average diameter breast height for each species
#' @param tpa trees per acre for each species
#' @param scale weibull scale parameter (default=4)
#' @param shape weibull shape parameter (default=2)
#' @param dmin minimum diameter for distribution (default 0.001")
#' @param dmax maximum diameter for distribution (default 100")
#' @param incr diameter increment between dmin and dmax (default 0.1")
#' @param hd height to diameter ratio for height dubbing (default 7)
#' @return data frame of list with example data for use with SVS()
#' @examples
#' SV <- SVS_ExampleData( species=c(202,263), tpa=c(300,275), dbh=c(7.8,4.7) )
#' SVS_ExampleData( datatype=NULL )     # force function to return known data types
#' @author James McCarter \email{jim.mccarter@@rayonier.com}
#' @export
SVS_ExampleData <- function( datatype='TBL2SVS', species, dbh, tpa, scale=4, shape=2, n=30, dmin=0.001, dmax=100, incr=0.1, hd=7 ) {
    if( is.null(datatype) ) {
        return( "Known data types are: StandViz (default) and TBL2SVS." )
    }
    tr <- treelist( species, dbh, tpa, scale, shape, n, dmin, dmax, incr )
    tr2 <- tr[c(3,2,5)]                                                     # get subset of columns we want: species, dbh, tpa
    tr2$ht <- tr2$dbh * rnorm(n,hd)                                         # dib in ht scaled from dbh with random normal noise around hd factor
    tr2$dbh[tr2$ht<4.5] <- 0.01                                             # if ht < 4.5, reset dbh to very small
    if( datatype=='TBL2SVS' ) {
        # Species, DBH, Height, CRat, Crad, Status, PlantClass, CrownClass, TPA
        tr2$CrownRatio <- 0.45                      # add CrownRatio
        tr2$CrownRadius <- tr2$ht * tr2$CrownRatio * 0.33 / 2.0     # add CrownRadius
        tr2$Status <- 1
        tr2$PlantClass <- 0
        tr2$CrownClass <- 0
        tr2 <- tr2[,c(1,2,4,5,6,7,8,9,3)]
        names(tr2) <- c("Species", "DBH", "Height", "CrownRatio", "CrownRadius", "Status", "PlantClass", "CrownClass", "TPA")
    } else if( datatype=='StandViz' ) {
        # Stand, Year/Age, Species, TreeNo, Live/Dead, Status, Condiiion, DBH, Height, CrowRat, CrownRad, TPA
        tr2$Stand <- 'Stand'                        # and stand name
        tr2$Year.Age <- substr( Sys.time(), 1, 4 )      # add year
        tr2$TreeNo <- seq.int(nrow(tr2) )           # add TreeNo
        tr2$CrownRatio <- 0.45                      # add CrownRatio
        tr2$CrownRadius <- tr2$ht * tr2$CrownRatio * 0.33 / 2.0     # add CrownRadius
        tr2$Live.Dead <- NA
        tr2$Status <- NA
        tr2$Condition <- NA
        tr2 <- tr2[,c(5,6,1,7,10,11,12,2,4,8,9,3)]
        names(tr2) <- c("Stand", "Year.Age", "Species", "TreeNo", "Live.Dead", "Status", "Condition", "DBH", "Height", "CrownRatio", "CrownRadius", "TPA")
    #} else if( datatype=='StandVizExtended' ) {
    #    print( "Return StandViz Extended format example data" )
    }
    return( tr2 )
}

#' List species codes
#'
#' List known species codes.
#'
#' This function simply reads and returns the list of known species codes distributed with the rSVS
#' package (rSVS_Species.csv). This file is also used as the basis for species translation by the
#' FIA2NRCS() and NRCS2FIA() funtions.
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
#' original SVS distribution).  Note that the tree form definitions have been exanded to handle more
#' live and dead classes for improved visualizations of forest health conditions.
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

#' Clean out svsfiles temorary directory
#'
#' @author James McCarter \email{jim.mccarter@rayonier.com}
#' @examples
#' svsfiles_clean()
#' @export
svsfiles_clean <- function() {
    filelist <- list.files( "svsfiles" )                # get list if files in svsfiles folder
    for( FILE in filelist ) {                           # loop across files to remove them
        file.remove( paste0( "svsfiles/", FILE ) )      # remove current file
    }
}

#' List files in svsfiles folder
#'
#' @author James McCarter \email{jim.mccarter@rayonier.com}
#' @examples
#' svsfiles_list()
#' @export
svsfiles_list <- function() {
    if( file.exists( 'svsfiles' ) ) return( dir( 'svsfiles' ) )
    else return( "'svsfiles' folder does not exist")
}

#' Convert species codes from FIA number to NRCS code
#'
#' Function not implemented yet.
#'
#' @param MyData DataFrame with tree information
#' @return Data frame with converted species codes
#' @author Jim McCarter \email{jim.mccarter@rayonier.com}
#' @examples
#' FIA2NRCS( MyData )
#' @export
FIA2NRCS <- function( Data ) {
    DataType <- Detect_DataType( Data )
    SppRef <- read.csv( system.file( "bin", "rSVS_Species.csv", package="rSVS" ) )[,c(1,2)] # read rSVS_Species.csv and get just FIA and NRCS fields
    if( DataType %in% c('StandVizObject','StandVizExtendedObject','SVScsvObject','TBL2SVSObject') ) {
        Data$Species <- as.character( Data$Species )                                        # coerce Species to character from factor
        SppRef$FIA <- as.character( SppRef$FIA )                                                # coerce Species to character from factor
        Data2 <- dplyr::left_join( Data, SppRef, by=c("Species" = "FIA") )                  # join original data with species code map
        Data2$FIA <- Data2$Species                                                          # copy original Species to FIA
        Data2$Species <- Data2$NRCS                                                         # copy NRCS over Species
        Data2$NRCS <- NULL                                                                  # remove NRCS column
        Data2$FIA <- NULL                                                                  # remove NRCS column
        return( Data2 )
    } else {
        print( paste0( "FIA2NRCS(): Don't know how to convert ", DataType, " object yet!") )
    }
}

#' Convert species codes from NRCS code to FIA number
#'
#' Function not implemented yet.
#'
#' @param MyData DataFrame with tree information
#' @return Data frame with converted species codes
#' @author Jim McCarter \email{jim.mccarter@rayonier.com}
#' @examples
#' NRCS2FIA( MyData )
#' @export
NRCS2FIA <- function( Data ) {
    DataType <- Detect_DataType( Data )
    SppRef <- read.csv( system.file( "bin", "rSVS_Species.csv", package="rSVS" ) )[,c(1,2)] # read rSVS_Species.csv and get just FIA and NRCS fields
    if( DataType %in% c('StandVizObject','StandVizExtendedObject','SVScsvObject','TBL2SVSObject') ) {
        Data$Species <- as.character( Data$Species )                                        # coerce Species to character from factor
        SppRef$NRCS <- as.character( SppRef$NRCS )                                          # coerce NRCS to character from factor
        Data2 <- dplyr::left_join( Data, SppRef, by=c("Species" = "NRCS") )                 # join original data with species code map
        Data2$NRCS <- Data2$Species                                                         # copy original Species to NRCS
        Data2$Species <- Data2$FIA                                                          # copy FIA over Species
        Data2$FIA <- NULL                                                                   # remove FIA column
        Data2$NRCS <- NULL                                                                   # remove FIA column
        return( Data2 )
    } else {
        print( paste0( "NCRS2FIA(): Don't know how to convert ", DataType, " object yet!") )
    }
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

