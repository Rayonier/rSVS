#!/usr/bin/python
#
# Filename:...: StandViz.py
# Description.: Stand level visualzation package using the Stand Visualization System (SVS)
# Author......: James B. McCarter
# Copyright...: 2020, Rayonier, Inc
# Requirements: Python 3.x, with numpy and pandas available
#

import argparse     # ArgumentParser()
import math         # math.ceil(), .floor(), .sqrt()
import os           # os.path.split(), os.path.splitext(), os.system()
import pandas as pd # Pandas DataFrame, .read_csv()
import platform     # platform.system()
import random       # random.seed(), .uniform()
import re           # re.search(), .sub()
import sys          # sys.argv, .exc_info(), .exit()
import time         # time.asctime()
#import win32com.client, pythoncom, _winreg

(__file_version__, __file_date__) = ( '$Revision: 1.0.5 $', '$Date: 2020/05/31 17:39:00 $' )
(_MyPath, _MyFile, _MyOS) = (os.path.split(sys.argv[0])[0], os.path.split(sys.argv[0])[1], platform.system())
(_MyVersion, _MyDate) = ( __file_version__.split()[1], '{} - {}'.format(__file_date__.split()[1], __file_date__.split()[2]) )

##################################
# Begin Global Data Declarations #
##################################

#OWNPATH = os.getcwd()
(OWNPATH, file) = os.path.split(sys.argv[0])
#global OWNPATH
#VERBOSE = 0

#########################################################
# Begin Main Program - implement command line inferface #
#########################################################
def main():     # implement __main__ scope for handling of command line execution of script
    global DEBUG, NOTIFY, VERBOSE
    try:
        (DEBUG, NOTIFY, VERBOSE) = (False, False, False)
        # ABC.E.GHIJKLM.OPQRSTUVWXYZ abcdef..ijklm..pqrs.u..xyz 0123456789
        SARG = argparse.ArgumentParser( add_help=False, usage=" %(prog)s [-o bitmap|csv|svs|web] [-g clumped|fixed|random|uniform] [-cf #] [-cr #] [-rf #] file [file [...]]\n" +
                                                             "\t%(prog)s [-v] [-F|-N] [-w worksheet] excelfile [excelfile [...]]\n"
                                                             "\t%(prog)s [-v] [-ta] [-tc FIA|NRCS] [-CF]" )
        SARGO = SARG.add_argument_group( "Output arguments" )
        SARGO.add_argument( "-D", action="store_true", help="Debug output" )
        SARGO.add_argument( "-n", action="store_true", help="Notify progress in DOS window" )
        SARGO.add_argument( "-o", action="store", nargs=1, metavar="format", help="Output format (bitmap|csv|svs(default)|web)" )
        SARGO.add_argument( "-v", action="store_true", help="Verbose output" )
        #SARGO.add_argument( "-os", action="store_true", help="Output to SVS (default)" )
        #SARGO.add_argument( "-ow", action="store_true", help="Output to HTML (create .png, generate .html page)" )
        #SARGO.add_argument( "-ox", action="store_true", help="Output to eXcel (.csv file)" )

        SARGC = SARG.add_argument_group( "Coordinate generation arguments" )
        SARGC.add_argument( "-cf", action="store", nargs=1, metavar="#", help="Clumpiness Factor (default 0.75)" )
        SARGC.add_argument( "-cr", action="store", nargs=1, metavar="#", help="Clump Ratio (n clumps = (0.01-0.5)*TPA)" )
        SARGC.add_argument( "-g", action="store", nargs=1, metavar="method", help="generate [clumped|fixe|random|uniform] coordinates" )    # or -c Clumped|c|Fixed|f|Random|r|Uniform|u
        #SARGC.add_argument( "-gc", action="store_true", help="generate clumped coordinates" )    # or -c Clumped|c|Fixed|f|Random|r|Uniform|u
        #SARGC.add_argument( "-gf", action="store_true", help="generate fixed coordinates" )
        #SARGC.add_argument( "-gr", action="store_true", help="generate random coordinates" )
        #SARGC.add_argument( "-gu", action="store_true", help="generate Uniform coordinates" )
        SARGC.add_argument( "-rf", action="store", nargs=1, metavar="#", help="Randomness Factor (0=perf rows, 0.4-0.8=plantation, <.8=clumps)" )

        SARGT = SARG.add_argument_group( "Treeform arguments" )
        SARGT.add_argument( "-CF", action="store_true", help="Create FIA.TRF from rSVS_Species.csv" )
        SARGT.add_argument( "-F", action="store_true", help="Use FIA treeform file" )
        SARGT.add_argument( "-N", action="store_true", help="Use NRCS treeform file (default FIA)" )
        SARGT.add_argument( "-ta", action="store_true", help="Audit rSVS_Species.xlsx file" )
        SARGT.add_argument( "-tc", action="store", nargs=1, metavar="TRFile", help="Compare treeform file versus rSVS_Species.xlsx" )

        SARGG = SARG.add_argument_group( "General arguments" )
        SARGG.add_argument( "-cd", action="store", nargs=1, metavar="factor", help="Crown Dubbing from height: CR=Height*Factor" )
        SARGG.add_argument( "-ds", action="store", nargs=1, metavar="factor", help="Diameter scaling: dbh*factor; 25:50 dbh>10*1.25, dbh>20*1.50" )
        SARGG.add_argument( "-e", action="store", nargs='?', metavar="factor", help="Expand coordinates by factor (default 2.0)", type=float, const=2.0, default=2.0 )
        SARGG.add_argument( "-h", action="store_true", help="display help" )
        SARGG.add_argument( "-hd", action="store", nargs=1, metavar="factor", help="Height Dubbing from DBH: Height(ft)=DBH(in)*Factor" )
        # or -s Diameter|d|XY|xy
        # or -e expand XY to fit into acre (e.g. for Postex plots)
        #SARGG.add_argument( "-m", action="store_true", help="Mechanical (row) thinning (only for Fixed coordinates)" )
        SARGG.add_argument( "-t", action="store_true", help="Test and debugging option" )
        SARGG.add_argument( "-w", action="store", nargs=1, metavar="name", help="worksheet name for Excel input" )
        #SARGG.add_argument( "-z", action="store_true", help="zip file for transfer" )

        SARG.add_argument( "FILELIST", nargs="*", help="Files [File [...]]")
        SOPT = SARG.parse_args()
        nFile = len( SOPT.FILELIST )

        if( SOPT.D ): DEBUG = 1
        if( SOPT.v ): VERBOSE = 1
        if( SOPT.n ): NOTIFY = 1

        OriginalWindowsPath = os.getcwd()                                       # save starting path
        ScriptPath = _MyPath                                                    # 
        if( ScriptPath == '' ): ScriptPath = OriginalWindowsPath
        if( DEBUG ): print( "OriginalPath={}, ScriptPath={}, os.path.realpath()={}".format(OriginalWindowsPath, ScriptPath, os.path.realpath(ScriptPath)) )
        SVSPath = os.path.normpath( os.path.split(ScriptPath)[0] + '/bin/SVS/winsvs.exe' )
        if( not os.path.exists( SVSPath ) ): print( "This command will fail!: {}".format(SVSPath) )
        if( DEBUG ): print( "SVSPath = {}".format(SVSPath))

        #os.chdir( ScriptPath )

        if( SOPT.ta ):
            Audit_rSVS_Species_File(ScriptPath)                                 # audit the rSVS_Species database
            sys.exit( "audited rSVS_Species.csv" )

        if( SOPT.tc ):                                                          # Compare TreeForm file against rSVS_Species.csv
            Compare_TreeForm_To_rSVS_Species( SOPT.C[0], SOPT.v )               # pass TreeForm file basename to function
            sys.exit("performed audit")

        if( SOPT.CF ):                                                          # create FIA.trf from rSVS_Species.xlsx
            Create_FIA_TreeForm_File()
            sys.exit( "created FIA.trf" )

        if( SOPT.t ):                                                           # hook for testing
            print( "No testing functionality currently defined!")
            #CMDLINE = ".\\inst\\bin\\SVS\\winsvs.exe ./inst/bin/SouthernPine.svs"
            #print( "StandViz.py: CMDLINE={}".format(CMDLINE) )
            #os.system(CMDLINE)
            sys.exit()
        
        if( (nFile==0) | SOPT.h ):                                              # if no files or -h specified
            SARG.print_help()                                                   # print help screen
            sys.exit( "help printed" )

        ExpandCoord = SOPT.e                                                    # if specified, copy for commandline, otherwise copy default value of 2.0

        if( SOPT.o ):                                                           # if -o specified, copy to OutFormat
            OutFormat = SOPT.o[0].lower()                                           # store as lower case string
            if( OutFormat in ['b','bit','bitmap','bmp','png','bmp'] ): OutFormat = 'bitmap'     # output BMP file and convert to .PNG
            if( OutFormat in ['c','csv','x','e','excel'] ): OutFormat = 'csv'       # output .csv file (excel) format
            if( OutFormat in ['s','svs'] ): OutFormat = 'svs'                       # output .svs for loading SVS program
            if( OutFormat in ['w','web','h','html'] ): OutFormat = 'web'            # output html web page with imbedded .png graphics
        else: OutFormat = 'svs'                                                 # else, default OutFormat='svs'

        if( SOPT.g ):                                                           # if -g specified, copy to GenMethod
            GenMethod = SOPT.g[0].lower()                                           # store as lower case string
            if( GenMethod in ['c','clump','clumped'] ): GenMethod = 'clumped'       # generate clumped coordinates (see -cf and -cr)
            if( GenMethod in ['f','fix','fixed'] ): GenMethod = 'fixed'             # generate fixed (rows) coordintes
            if( GenMethod in ['r','ran','random'] ): GenMethod = 'random'           # generate random coordinates
            if( GenMethod in ['u','unif','uniform'] ): GenMethod = 'uniform'        # generate uniform coordinates
        else: GenMethod = 'random'                                              # else, default GenMethod='random'

        if( DEBUG ): print( "OutFormat={}, GenMethod={}, ExpandCoord={}".format(OutFormat, GenMethod, ExpandCoord) )

        #if( (SOPT.gc==0) & (SOPT.gf==0) & (SOPT.gr==0) & (SOPT.gu==0) ): SOPT.gr = True   # random is default coordinate generation
        #if( (SOPT.ob==0) & (SOPT.ow==0) & (SOPT.os==0) & (SOPT.ox==0) ): SOPT.os = 1   # SVS is default output

        if( NOTIFY ): print( 'StandViz.py - Python implementation of Stand Visualization Addin for Excel' )
        if( DEBUG ): print(sys.argv)
        #if( DEBUG ): print( 'nFile={}, FILELIST={}' % (nFile, SOPT.FILELIST) )
        #if( DEBUG ): print( 'Using Python {} on {} from {}'.format(sys.version, sys.platform, sys.prefix) )

        for FILE in SOPT.FILELIST:
            #D = {}              # create data dictionary
            (dirname, filename) = os.path.split( FILE )                         # get path and filename for file from command line
            (basename, ext) = os.path.splitext( filename )                      # get filebase and extension
            if( DEBUG ): print( "File: {}, dirname={} filename={} basename={} ext={}".format(FILE, dirname, filename, basename, ext) )
            #DataSet = 'None'
            FileType = 'unknown'                                                # start with FileType='unknown'
            # determine file format from filename provided on command line
            if( re.search( '.csv', filename ) != None ):                        # have .csv extension
                DataSet = re.sub( '.csv', '', filename )                        # name dataset from base filename
                FileType = Determine_CSV_Format( FILE )                         # determine file format
                if( DEBUG ): print( "{}: FileType={}".format(FILE,FileType) )
            elif( re.search( '.svs', filename ) != None ):                      # have .svs extension, just pass through to winsvs.exe if the file exists
                CMDLINE = "{} {}".format(SVSPath, FILE)                         # build command line
                if( VERBOSE ): print(CMDLINE)                                   # echo command line
                os.system(CMDLINE)                                              # execute command line
                return                                                          # exit program
            elif( re.search( '.xlsx', filename ) != None ):                     # have .xlsx extension
                DataSet = re.sub( '.xlsx', '', filename )                       # name dataset from filename
            elif( re.search( '.xls', filename ) != None ):                      # have .xls extension
                DataSet = re.sub( '.xls', '', filename )                        # name dataset from filename

            if( FileType=='FMDObject' ):
                print( "Creating FDM visualizations..." )
                D = pd.read_csv( FILE )     # read .csv file
                # now process into pieces by PlotKey and MeasDate
                # PlotKey, TreeKey, Species, MeasDate, MeaseAge, Status, Condition, Damage, Screen, DBH, Height, CrownRatio, TPA
                # need to accumulate data into dictionary by PlotKey and MeasDate
                TD = {}
                FileNames = []
                for d in D.itertuples():
                    (Plot,Tree,Spp,MeasDate,DBH,Ht,CRat,Status,Cond,Dam,TPA) = (d.PlotKey,d.TreeKey,d.Species,d.MeasDate,d.DBH,d.Height,
                                                                                d.CrownRatio,d.Status,d.Condition,d.Damage,d.TPA)
                    if( not Plot in TD ): TD[Plot] = {}
                    if( not MeasDate in TD[Plot] ): TD[Plot][MeasDate] = {}
                    TD[Plot][MeasDate][Tree] = (Spp,DBH,Ht,CRat,Status,Cond,Dam,TPA)
                print( "Plots={}".format(TD.keys()) )
                for P in sorted(TD.keys()):
                    print( "Plot={}, Years={}".format(P,sorted(TD[P].keys())) )
                    for Y in sorted(TD[P].keys()):
                        OutFilename = "{}/{}-{}.asc".format(dirname,P,Y)
                        SvsFilename = "{}/{}-{}.svs".format(dirname,P,Y)
                        print( "Creating {} and {}".format(OutFilename, SvsFilename))
                        FileNames.append(SvsFilename)
                        OUT = open( OutFilename, 'w' )
                        OUT.write( ";species dbh height crat crad status pclass cclass tpa\n")
                        for T in sorted(TD[P][Y].keys()):
                            (Spp,DBH,Ht,CRat,Status,Cond,Dam,TPA) = TD[P][Y][T]
                            if( pd.isna(DBH) ): DBH = 0.01
                            if( pd.isna(Ht) ): Ht = DBH * 6
                            if( pd.isna(CRat) | (CRat==0) ): CRat = 0.33
                            if( Status == 'Live' ): Status = 1
                            else: Status = 2
                            PClass = 0
                            CRad = CRat * Ht / 4
                            CClass = 0
                            OUT.write( "{} {} {} {} {} {} {} {} {}\n".format(Spp,DBH,Ht,CRat,CRad,Status,PClass,CClass,TPA) )
                        OUT.close()
                        OPT = open("{}/{}-{}.opt".format(dirname,P,Y), 'w' )
                        OPT.write( "-P1 -N 0 -H 0.33 -T..\\inst\\bin\\SVS\\FIA.trf {} {}".format(OutFilename,SvsFilename) )
                        OPT.close()
                        cmdline = "{} -G -X{}/{}-{}.opt {}".format(SVSPath,dirname,P,Y,SvsFilename)
                        print( "cmdline={}".format(cmdline) )
                        os.system(cmdline)
                print( FileNames )
            elif( FileType=='LMSObject' ):
                print( "Creating LMS visualizations..." )
                D = pd.read_csv( FILE )     # read .csv file
                # now process into pieces by PlotKey and MeasDate
                # PlotKey, TreeKey, Species, MeasDate, MeaseAge, Status, Condition, Damage, Screen, DBH, Height, CrownRatio, TPA
                # need to accumulate data into dictionary by PlotKey and MeasDate
                TD = {}
                FileNames = []
                for d in D.itertuples():
                    (Stand,Year,Tree,Spp,DBH,Height,CRat,Status,PC,CC,TPA) = (d.STANDNAME,d.year,d.OBJECTID,d.SPECIES,d.QDBH,d.HEIGHT,d.cr,d.status,d.pc,d.cc,d.TPA)
                    if( not Stand in TD ): TD[Stand] = {}
                    if( not Year in TD[Stand] ): TD[Stand][Year] = {}
                    TD[Stand][Year][Tree] = (Spp,DBH,Height,CRat,Status,PC,CC,TPA)
                for P in sorted(TD.keys()):
                    for Y in sorted(TD[P].keys()):
                        OutFilename = "{}/{}-{}.asc".format(dirname,P,Y)
                        SvsFilename = "{}/{}-{}.svs".format(dirname,P,Y)
                        FileNames.append(SvsFilename)
                        OUT = open( OutFilename, 'w' )
                        OUT.write( ";species dbh height crat crad status pclass cclass tpa\n" )
                        for T in sorted(TD[P][Y].keys()):
                            (Spp,DBH,Ht,CRat,Status,PC,CC,TPA) = TD[P][Y][T]
                            if( pd.isna(DBH) ): DBH = 0.01
                            if( pd.isna(Ht) ): Ht = DBH * 6
                            if( pd.isna(CRat) | (CRat==0) ): CRat = 0.45
                            if( Status=='Live'): Status = 1
                            #else: Status = 2
                            PClass = PC
                            CRad = CRat * Ht / 4
                            CClass = 0
                            OUT.write( "{} {} {} {} {} {} {} {} {}\n".format(Spp,DBH,Ht,CRat,CRad,Status,PClass,CClass,TPA ) )
                        OUT.close()
                        OPT = open( "{}/{}-{}.opt".format(dirname,P,Y), 'w' )
                        OPT.write( "-P1 -N 0 -H 0.33 -T..\\inst\\bin\\SVS\\FIA.trf {} {}".format(OutFilename,SvsFilename) )
                        OPT.close()
                        cmdline = "{} -G -X{}/{}-{}.opt {}".format(SVSPath,dirname,P,Y,SvsFilename)
                        print( "cmdline={}".format(cmdline) )
                        os.system(cmdline)
            elif( FileType=='PosTex' ):
                # data loader for Postex plots: Plot, Plot_Radius, Nr, Tree_Spc, Tree_Dia(.1 in), Tree_Hgt(ft), Tree_Postex1, Tree_Poste2, Tree_Postx3,
                # Tree_Local_x, Tree_Local_y, Tree_Local_Dist, Tree_Local_Angle, Tree_Angle_ToPlotCenter, Latitude, Longitude, Tree_Nr
                # TreeSpc: 1=Unforked pine, 2=hardwood, 3=dead tree (pine or hardwood), 4=forked pine
                CsvFileName = "{}".format(FILE)
                # split path from filename and create .svs in svsfiles folder
                print( "Processing {}...".format(CsvFileName))
                # check that it exists
                D = pd.read_csv( CsvFileName )
                DataSet = re.sub( '.csv', '', filename )
                SVS = StandViz( DataSet )
                Year = 2020
                SvsFilename = "{}_{}.svs".format(DataSet,Year)
                SVS.SVF = open( SvsFilename, 'w' )
                SVS.SVS_Write_Header()
                for L in D.itertuples():
                    standname = L.Plot
                    SVS.Data.Stand[standname] = StandData(standname)
                    (TreeNo, Species, DBH, Ht, X, Y) = (L.Nr, L.Tree_Spc, L.Tree_Dia, L.Tree_Hgt, L.Tree_Local_x, L.Tree_Local_y)
                    X = (208.71 / 2.0 ) + ((float(X)*3.28084)) * ExpandCoord
                    Y = (208.71 / 2.0 ) + ((float(Y)*3.28084)) * ExpandCoord
                    if( Species == 1 ): (Species, Status) = ('PITA', 1)
                    elif( Species == 2 ): (Species, Status) = ('HARDWOOD', 1)
                    elif( Species == 3 ): (Species, Status) = ('SNAG', 2)
                    elif( Species == 4 ): (Species, Status) = ('PITA', 2)
                    nTree = len(SVS.Data.Stand[standname].Tree) + 1
                    SVS.Data.Stand[standname].Tree[nTree] = TreeData(Species,TreeNumber=TreeNo, X=X, Y=Y)
                    SVS.Data.Stand[standname].Tree[nTree].Year[Year] = MeasurementData( DBH, Ht, '', 1, 0, Status )
                    #print("{},{},{},{},{},{},{},{}".format(standname,TreeNo,Species,DBH,Ht,X,Y,nTree))
                    (LAng, Bearing, EDia, Mark, Z) = (0,0,0,0,0)
                    TPA = 1
                    DBH /= 10
                    PClass = 0
                    CClass = 1
                    CR = 0.45
                    CW = 10
                    SVS.SVS_Write_Tree_Live( Species, TreeNo, PClass, CClass, Status, DBH, Ht, LAng, Bearing, EDia, CW, CR, TPA, Mark, X,Y, Z)
                #print("Stand.keys()={}".format(SVS.Data.Stand.keys()))
                SVS.SVS_Write_Footer()
                SVSEXE = "inst\\bin\SVS\winsvs.exe"
                CMDLINE = "{} -A 180 -D 325 {}".format(SVSEXE, SvsFilename)
                print(CMDLINE)
                os.system(CMDLINE)
            elif( FileType=='StandObject' ):
                print( "visualizing {}".format(FILE))
                D = pd.read_csv( FILE )
                print( "{} lines read".format(len(D.index)))
                OutFilename = "{}/{}.asc".format(dirname,basename)
                SvsFilename = "{}/{}.svs".format(dirname,basename)
                OptFilename = "{}/{}.opt".format(dirname,basename)
                print( "OutFilename={}".format(OutFilename))
                OUT = open( OutFilename, 'w' )
                OUT.write( ";species dbh height crat crad status pclass cclass tpa\n")
                for d in D.itertuples():
                    OUT.write( "{} {} {} {} {} {} {} {} {}\n".format(d.species,d.dbh,d.height,d.cr,d.crad,d.status,d.pc,d.cc,d.tpa))
                OUT.close()
                OUT = open( OptFilename, 'w')
                if( SOPT.N ): OUT.write( "-P1 -N 0 -H 0.33 -T..\\inst\\bin\SVS\\NRCS.trf {} {}".format(OutFilename,SvsFilename) )
                else: OUT.write( "-P1 -N 0 -H 0.33 -T..\\inst\\bin\SVS\\FIA.trf {} {}".format(OutFilename,SvsFilename) )
                OUT.close()
                #SVS = StandViz( basename )
                SVSEXE = "inst\\bin\SVS\winsvs.exe"
                if( not os.path.exists( SVSEXE ) ): print( "This command will fail!: {}".format(SVSEXE))
                cmdline = "{} -G -X{} {}".format(SVSEXE,OptFilename,SvsFilename)
                print( "cmdline={}".format(cmdline) )
                os.system(cmdline)
            elif( FileType=='StandViz' ):
                D = pd.read_csv( FILE )
                print( "{} lines read".format(len(D.index)))
                OutFilename = "{}/{}.asc".format(dirname,basename)
                SvsFilename = "{}/{}.svs".format(dirname,basename)
            elif( FileType=='TBL2SVSObject' ):
                print( "visualizing {}".format(FILE))
                D = pd.read_csv( FILE )
                print( "{} lines read".format(len(D.index)))
                OutFilename = "{}/{}.asc".format(dirname,basename)
                SvsFilename = "{}/{}.svs".format(dirname,basename)
                OptFilename = "{}/{}.opt".format(dirname,basename)
                print( "OutFilename={}".format(OutFilename))
                OUT = open( OutFilename, 'w' )
                OUT.write( ";species dbh height crat crad status pclass cclass tpa\n")
                for d in D.itertuples():
                    if( pd.isna(d.DBH) ): DBH = 0.01
                    else: DBH = d.DBH
                    if( pd.isna(d.Height) | (d.Height <= 0) ): Height = d.DBH * 6
                    else: Height = d.Height
                    if( pd.isna(d.CrownRatio) | (d.CrownRatio <= 0) ): CRat = 0.33
                    else: CRat = d.CrownRatio
                    if( pd.isna(d.CrownRadius) ): CRad = CRat * Height / 4.0
                    else: CRad = d.CrownRadius
                    OUT.write( "{} {} {} {} {} {} {} {} {}\n".format(d.Species,DBH,Height,CRat,CRad,d.Status,d.PlantClass,d.CrownClass,d.TPA))
                OUT.close()
                OUT = open( OptFilename, 'w')
                if( SOPT.N ): OUT.write( "-P1 -N 0 -H 0.33 -T..\\inst\\bin\SVS\\NRCS.trf {} {}".format(OutFilename,SvsFilename) )
                else: OUT.write( "-P1 -N 0 -H 0.33 -T..\\inst\\bin\SVS\\FIA.trf {} {}".format(OutFilename,SvsFilename) )
                OUT.close()
                if( not os.path.exists( SVSPath ) ): print( "This command will fail!: {}".format(SVSEXE))
                cmdline = "{} -G -X{} {}".format(SVSPath,OptFilename,SvsFilename)
                print( "cmdline={}".format(cmdline) )
                os.system(cmdline)
            else:
                print( "Error, Sorry I don't know how to handle this kind of data yet!")

            #print 'DataSet = %s' % (DataSet)
            #SVS = StandViz( DataSet )                 # create class/dataset for input file

            # if extension is .xls or xlsx file then need to determine if we are a SvsAddin format or TIR format file
            #if( ext in ['.xls', '.xlsx' ] ):            # test eXcel file for type
            #    FileFormat = 'Excel'
            #    FileFormat = Test_Excel_Format( f )
            #    #raw_input( "Paused: After Test_Excel_Format(): %s is %s" % (f, FileFormat) )
            #    else:       # unknown excel file format
            #        sys.exit()
            #elif( ext in [ '.csv' ] ):
            #    SVS.CSV_Load_File( f )                  # load SvsAddin format .csv file
            #    #print 'D.Stand.keys()=%s' % (D.Stand.keys())
            #    #raw_input( "Processing .csv file, press return to continue: " )
            #else:
            #    raw_input( 'unknown file type, press return to exit' )
            #    sys.exit()

            #raw_input( "Pause" )

            #if( OPT['c'] ):                             # generate tree coordinates based on requested pattern
            #    SVS.Generate_Clumped( 15, 40 )          # generate clummped coordinates
            #    # should be using cLumpiness and clumPration parameters
            #elif( OPT['f'] ):
            #    SVS.Generate_Fixed()                    # generate fixed coordinates
            #elif( OPT['r'] ):
            #    SVS.Generate_Random()                   # generate random coordinates
            #    # should be using the randomness factor
            #elif( OPT['u'] ):
            #    SVS.Generate_Uniform( Variation=2.0 )   # generate uniform coordinates

            #if( OPT['s'] ):                             # output to SVS
            #    if( DEBUG ): print( 'output SVS' )
            #    SVS.SVS_Create_Files( dirname )
            #    SVS.SVS_Show_Files( dirname )
            #elif( OPT['x'] ):                           # output to Excel .csv file
            #    if( DEBUG ): print( 'output csv file' )
            #    SVS.CSV_Write_File( '%s/%s.csv' % (dirname, DataSet) )
            #elif( OPT['h'] ):                           # output to html page
            #    if( DEBUG ): print( 'output html' )
            #    SVS.SVS_Create_Files( dirname )
            #    SVS.SVS_Webpage_Create( dirname )
            #    if( OPT['z'] ): SVS.SVS_Webpage_Zip( dirname )   # if -z then zip the website for download
            #elif( OPT['b'] ):                           # output to bitmaps (.PNG)
            #    if( DEBUG ): print( 'ouptut bmp file' )
            #    SVS.SVS_Create_Files( dirname )
            #    SVS.SVS_Create_Bitmaps( dirname )

    except SystemExit:
        pass
    except:
        StandViz_ReportError( sys.exc_info(), sys.argv, Header='StandViz.py\n' )

##############################
# Begin Function Definitions #
##############################

def StandViz_ReportError( errorobj, args, Header = None ):              # error reporting and traceback function
    """ReportError( sys.exec_info(), errorfilename, sys.argv )"""
    (MyPath, MyFile) = os.path.split( args[0] )                         # retrieve filename and path of running python script
    (MyBaseName, MyExt) = os.path.splitext( MyFile )                    # separate basefilename from extension
    errorfilename = "{}.txt".format(MyBaseName)                         # create new error filename based on base of script filename
    ERRFILE = open( errorfilename, 'w' )                                # open text file for writting
    if( Header != None ): ERRFILE.write( '%s\n' % Header )              # if Header defined, write Header to file
    ERRFILE.write( "Error running '{}'\n".format(MyFile) )              # write error message with filename
    MyTrace = errorobj[2]                                               # retrieve error object
    while( MyTrace != None ):                                           # loop through stack trace
        (line, file, name) = ( MyTrace.tb_lineno, MyTrace.tb_frame.f_code.co_filename, MyTrace.tb_frame.f_code.co_name )    # extract line, file, and error name
        F = open( file, 'r' )                                           # open source file of Python script
        L = F.readlines()                                               # read scripot source into memory
        F.close()                                                       # close script file
        code = L[line-1].strip()                                        # extract line of source code that caused error
        ERRFILE.write( "  File '{}', line {}, in {}\n    {}\n".format(file, line, name, code) )     # write filename, source code line, error name, and error code
        MyTrace = MyTrace.tb_next                                       # step to next level of call stack trace
    ERRFILE.write( "errorobj: {}\n".format(errorobj) )                  # write error object and arguments for call
    ERRFILE.write( "Calling Argument Vector: {}\n".format(args) )       # write calling arguments
    ERRFILE.close()                                                     # close text file with error stack trace
    os.system( "notepad.exe {}".format(errorfilename) )                 # display error log file with notepad.exe

def Audit_rSVS_Species_File( ScriptPath ):
    print( "Performing audit of rSVS_Species.xlsx file" )
    SppXlsFile = "{}/rSVS_Species.xlsx".format(os.path.realpath("{}/bin".format(os.path.split(ScriptPath)[0])))
    print( "ScriptPath={}, SppXlsFile={}".format(ScriptPath,SppXlsFile) )
    print(SppXlsFile)
    SPPXLS = pd.ExcelFile( SppXlsFile )
    SPP = SPPXLS.parse( 'rSVS_Species' )                                                    # get species list from rSVS_Species sheet
    print( "Read {} lines from {}".format(len(SPP.index), SppXlsFile) )
    DUP = {'FIA':{}, 'NRCS':{}}
    (nMissF, nMissN, nMissG, nMissS, nMissC) = (0, 0, 0, 0, 0)          # set missing counters
    (nDupF, nDupN) = (0,0)                                              # set duplicate counters
    for S in SPP.itertuples():                                          # loop across rows in file
        (FIA, NRCS, Genus, Species, Common, Comment, NRCSTRF, FVSVar, FVSSpp) = (S.FIA, S.NRCS, S.Genus, S.Species, S.Common, S.Comment, S._7, S._8, S.SpCode)
        if( pd.isna(FIA) ): nMissF += 1                                 # FIA # missing
        else:
            FIA = int(FIA)                                              # have FIA, convert to integer
            if( not FIA in DUP['FIA'] ): DUP['FIA'][FIA] = 1            # if not seen before, store number
            else:                                                       # else already have, duplicate
                print( "Duplicate FIA #{}".format(FIA) )                # output message
                nDupF += 1                                              # increment duplicate counter for FIA
        if( pd.isna(NRCS) ): nMissN += 1                                # NRCS code missing
        else:
            if( not NRCS in DUP['NRCS'] ): DUP['NRCS'][NRCS] = 1        # if not seen before, store code
            else:                                                       # else already have, duplicate
                print( "Duplcate NRCS code: {}".format(NRCS) )          # output message
                nDupN += 1                                              # increment duplicate counter for NRCS
        if( pd.isna(Genus) ): nMissG += 1                               # Genus missing
        if( pd.isna(Species) ): nMissS += 1                             # Species missing
        if( pd.isna(Common) ): nMissC += 1                              # Common name missing
        #print( "{}, {}, {} {}, {}, {}, {}, {}, {}".format(FIA, NRCS, Genus, Species, Common, Comment, NRCSTRF, FVSVar, FVSSpp) )
    print( "Total Species {}: ".format( len(SPP.index) ) )              # output audit
    print( "    FIA: Have {}, Missing {}, Dup {}".format( len(DUP['FIA'].keys()), nMissF, nDupF) )
    print( "    NRCS Have {}, Missing {}, Dup {}".format( len(DUP['NRCS'].keys()), nMissN, nDupN ) )
    print( "    Genus (Missing {}), Species (Missing  {}), Common (Missing {})".format( nMissG, nMissS, nMissC) )
    #os.chdir( OriginalWindowsPath )

def Compare_TreeForm_To_rSVS_Species( SppCodes, Verbose=False ):
    print( "Compare audit of {}.trf against rSVS_Species.xlsx".format(SppCodes) )
    SppXlsFile = "../bin/rSVS_Species.xlsx"                             # set path to rSVS_Species.xlsx
    SPPXLS = pd.ExcelFile( SppXlsFile )                                 # open excel file
    SPP = SPPXLS.parse( 'rSVS_Species' )                                # parse rSVS_Species worksheet
    print( "Read {} lines from {}".format(len(SPP.index), SppXlsFile) ) # output message
    TreeFormFile = "../bin/SVS/{}.trf".format(SppCodes)                 # make path to appropriate TreeForm file
    # should test for existance of file
    SVSTF = SVS_TreeForm()
    (SpecialForm, SppForm) = SVSTF.SVS_LoadTreeFormFile( TreeFormFile )       # load TreeFormFile
    print( "Read {} species and {} species treeforms from {}".format(len(SppForm.keys()), len(SpecialForm.keys()), TreeFormFile) )  # output status message
    AUDIT = {}                                                          # create dictionary to compare
    for S in SPP.itertuples():                                          # loop across rows in spreadsheet (rSVS_Species)
        (FIA, NRCS, Genus, Species) = (S.FIA, S.NRCS, S.Genus, S.Species)   # get columns of interest
        if( SppCodes == 'NRCS' ):                                       # if NRCS
            if( not NRCS in AUDIT ): AUDIT[NRCS] = 1
        elif( SppCodes == 'FIA' ):                                      # if FIA
            if( pd.isna(FIA) ): continue                                # if missing, skip
            else: FIA = "{}".format(int(FIA))
            if( not FIA in AUDIT ): AUDIT[FIA] = 1
    #print(sorted(AUDIT.keys()))
    (Have, Missing) = (0, 0)                                            # initialize counters
    if( Verbose ): print( "Species Codes: " )
    for S in sorted(SppForm.keys()):                                    # loop across TreeForms
        #print("'{}'".format(S))
        if( not S in AUDIT ): Missing += 1
        else: 
            Have += 1
            if( Verbose ): print( " {}".format(S), end='' )
    if( Verbose ): print("")
    print( "{}: Has {}, Missing {}".format(TreeFormFile, Have, len(SPP.index)-Have) )   # report results

def Determine_CSV_Format( FileName, debug=False ):
    FileType = 'Unknown'
    # first make sure file exists
    if( os.path.exists( FileName ) ):
        if( debug ): print( "Determine_CSV_Format(): {} exists".format(FileName) )
        F = pd.read_csv( FileName )                                             # read with pd.read_csv()
        #print( list(F.columns) )         
        #PosTex: Plot,Plot_Radius,Nr,Tree_Spc,Tree_Dia,Tree_Hgt,Tree_PosTex1,Tree_PosTex2,Tree_PosTex3,Tree_Local_x,Tree_Local_y,Tree_Local_Dist,Tree_Local_Angle,
        #Tree_Angle_ToPlotCenter,Latitude,Longitude,Tree_Nr
        if( 'Tree_PosTex1' in F.columns ): FileType = 'PosTex'
        elif( ('PlotKey' in F.columns) & ('TreeKey' in F.columns) & ('CrownRatio' in F.columns) ): FileType = 'FMDObject'
        elif( ('STANDNAME' in F.columns) & ('SPECIES' in F.columns) & ('QDBH' in F.columns) ): FileType = 'LMSObject'
        elif( ("species" in F.columns) & ("dbh" in F.columns) ): FileType = 'StandObject'
        # StandViz: Stand,Year/Age,Species,TreeNo,Live/Dead,TreeStat,CrwnClass,DBH,Height,Cradius,Cratio,TPA,X,Y
        # StandVizExtended: Stand,Year/Age,Species,TreeNo,Live/Dead,Status,Condition,DBH,Height,CrownRatio,CrownRadius,TPA,BrokenHt,Offset,Bearing,Lean,RootWad,X,Y
        elif( ('Year/Age' in F.columns) | ('Year.Age' in F.columns) ):
            if( 'RootWad' in F.columns ): FileType = 'StandVizExtended'
            else: FileType = 'StandViz'
        # SVScsv: Species,TreeNo,PlntClass,CrwnClass,Status,DBH,Height,LAng,FAng,SDia,CRad1,CRat1,CRad2,CRat2,CRad3,CRat3,CRad,CRat4,ExpFactor,MarkCode,X,Y,Z
        elif( ('PlntClass' in F.columns) & ('CRat1' in F.columns) ): FileType = 'SVScsv'
        elif( ('Species' in F.columns) & ('PlantClass' in F.columns) & ('CrownClass' in F.columns) ): FileType = 'TBL2SVSObject'
        else: print( "Unknown filetype: columns = {}".format(F.columns) )
    else:
        print( "Error, file '{}' does not exist!".format(FileName) )
    return( FileType )

# create SVSTreeForm class and move the SVS_LoadTreeFormFile(), SVS_Write_TreeFormFile(), SVS_WriteHeader() and other appropriate functions into class



# -B# clump ratio # clumps = (0.01 - 0.5) * TPA
# -G# clumpiness factor = 1.5-1.4*clumpiness factor)*clump spacing
# -R# Randomness Factor (0 = perfect rows and columns; 0.4-0.8 aproximate planted stands; > 0.8 some clumps of 2-3 trees

# not used: egijknoqy

###########################
# Begin Class Definitions #
###########################

##################################################################################################################
# TreeData, StandData, and ForestData classes store tree information stored by Forest (dataset), stand, and tree
##################################################################################################################
#
# D = ForestData( ForestName )
# D.Stand[StandName] = StandData( StandName )
# D.Stand[StandName].Plot[PlotName] = PlotData( 0, Size=1.0 )
# D.Stand[StandName].Plot[PlotName].Tree[TreeNo] = TreeData( Species, TreeNumber, X, Y )
# D.Stand[StandName].Plot[PlotName].Tree[TreeNo].Year[Year] = (DBH, Height, CrownRatio, TPA, Live, Status, Condition, ... )
#
##########################################################################################
class ForestData:
    """class for containing forest/data set/project/file level inventory information"""
    def __init__( self, Name ):
        self.Name = Name                    # name for forest/data set
        self.Stand = {}                     # dictionary for StandData objects

##########################################################################################
class StandData:
    """class for holding stand level information"""
    def __init__( self, Name, Plots=False ):
        self.Name = Name                    # name for stand
        if( Plots ):
            self.Plot = {}                      # dictionary to hold PlotData objects
        else:
            self.Tree = {}                      # dictionary for TreeData objects
            self.Year = {}                      # dictionary to hold stand summary information

##########################################################################################
class PlotData:
    """class for holding plot level information within a stand and dictionary of TreeData for tree information"""
    def __init__ (self, Name, Size=1.0 ):
        self.Name = Name
        self.Size = Size
        self.Tree = {}                      # dictionary to hold TreddData objects

##########################################################################################
class TreeData:
    """class for holding tree information (what does nto change and a dictionary of MeasurementData by year/age (changes with time)"""
    def __init__( self, Species=None, TreeNumber=None, X=None, Y=None ):
        self.Species = Species              # species
        self.TreeNumber = TreeNumber        # tree numbers
        self.X = X                          # tree X coordinate
        self.Y = Y                          # tree Y coordinate
        self.Year = {}                      # dictionary for holding MeasurementData objects

##########################################################################################
class MeasurementData:
    """The MeasurementData class holds tree measurement information"""
    def __init__( self, DBH=None, Height=None, CrownRatio=None, TPA=None, Live=None, Status=None, Condition=None,
                  Bearing=None, BrokenHeight=None, BrokenOffset=None, CrownRadius=None, DMR=None, LeanAngle=None, RootWad=None ):
        self.DBH = DBH                      # Diameter at Breat Height
        self.Height = Height                # Height
        self.CrownRatio = CrownRatio        # Crown Ratio
        if( CrownRatio == None ): self.CrownRatio = 0.45
        if( CrownRatio.strip() == '' ): self.CrownRatio = 0.45
        self.TPA = TPA                      # Trees Per Acre
        self.Live = Live                    # Live/Dead status (d|dead, dying, l|live, s|stump)
        if( Live == None ): self.Live = 'l'
        self.Status = Status                # Status = b|broken, brokentop, deadtop, d|down, s|standing
        self.svsStatus = None
        self.Condition = Condition          # Condition = Live:1|d|dominant,2|c|codominant,3|i|intermediate,4|s|suppressed; Dying:1-3; Dead:1-5
        self.Bearing = Bearing              # Bearing for down trees or broken tops
        self.BrokenHeight = BrokenHeight    # Height a which tree is broken
        self.BrokenOffset = BrokenOffset    # Distance top is from tree (in Bearing direction)
        self.CrownRadius = CrownRadius      # Crown Radius
        if( CrownRadius == None ): self.CrownRadius = (float( self.Height ) * 0.33) / 2.0
        self.DMR = DMR                      # Hawksworth Dwarf Mistelto Rating
        self.LeanAngle = LeanAngle          # Angle tree leaning (not implemented yet)
        self.RootWad = RootWad              # Radius of root wad


# D = ForestData( "DataName" )
# D.Stand[1] = StandData( ... )
# D.Stand[1].Plot[1] = PlotData()
# D.Stand[1].Plot[1].Tree[1] = TreeData( Species='DF', DBH=5.2, TPA=10.2 )
# for s in D.Stand.keys()
#     for p in D.Stand[s].Plot.keys()
#         for t in D.Stand[s].Plot[p].Tree.keys()
#             (Spp, Dbh, Ht, Cr, TPA) = D.Stand[s].Plot[p].Tree[t]

##############################################################################
# SVS_TreeForm class to abstract and provide interface to SVS treeform files #
##############################################################################
class SVS_TreeForm:
    """class to abstract and provide interface to SVS treeform files"""
    def __init__( self ):
        '''Initialze SVS_TreeForm class'''
        pass

    def __del__( self ):
        '''Destrop SVS_TreeForm class'''
        pass

    def Create_FIA_TreeForm_File():
        """Create FIA.trf file from rSVS_Species.xlsx file"""
        print( "Creating FIA.trf..." )
        SppXlsFile = "../bin/rSVS_Species.xlsx"
        SPPXLS = pd.ExcelFile( SppXlsFile )
        SPP = SPPXLS.parse( 'rSVS_Species' )                                                    # get tree records from Blackrock worksheet
        print( "Read {} lines from {}".format(len(SPP.index), SppXlsFile) )
        TRANSLATE = {}
        for S in SPP.itertuples():
            TRANSLATE[S.NRCS] = S.FIA
        TreeFormFile = "../bin/SVS/{}.trf".format('NRCS')
        (SpecialForm, SppForm) = SVS_LoadTreeFormFile( TreeFormFile )
        FIAForm = {}
        print( "Read {} lines from {}".format(len(SppForm.keys()), TreeFormFile) )
        # loop through SppForm.keys() and change species to FIA # from SPP
        for S in sorted(SppForm.keys()):
            if( not S in TRANSLATE ): print( "No FIA # for {}, skipping.".format(S) )
            else: 
                #print("Need to translate {} to {}".format(S,int(TRANSLATE[S])))
                FIAForm[int(TRANSLATE[S])] = SppForm[S]
                #input(FIAForm[int(TRANSLATE[S])])
        NewTreeFormFile = '../bin/SVS/FIA.trf'
        SVS_Write_TreeFormFile( NewTreeFormFile, SpecialForm, FIAForm )

    def SVS_LoadTreeFormFile( self, TreeFormFile ):
        SppForm = {}
        SpecialForm = {}
        SpecialList = [ '--','@flame.eob','CAR','CRANEBOOM','CRANETOWER','CONIFER','CUBE','DEFAULT','DMBROOM','HARDWOOD','MARKER','MISTBROOM','OTHER','PALM',
                        'R6CLUMP','R6SHRUB','R6SNAG','RANGEPOLE','ROCK','ROOTWAD','SEEDLING','SHRUB','SNAG','SNAG2','SNAG3','SNAG4','SNAG5','TETRAHEDRON','TRUCK' ]
        TRF = open( TreeFormFile, 'r' )
        for L in TRF:
            if( re.search( "^;", L ) != None ): pass    # skip comment/header lines
            else:
                (Spp, PlantClass, CrownClass, PlantForm, NoBranch, NoWhorl, BrBase, BrAngle, LowX, LowY, HighX, HighY, BaseUp, TopUp, StemCol, BrCol,
                 Fol1, Fol2, SampHt, SampRat, SampRad, Scale) = L.split()
                if( Spp in SpecialList ):               # if Spp code in SpedialList save to SpecialForm
                    SpecialForm[Spp] = {}
                    if( not PlantClass in SpecialForm[Spp] ): SpecialForm[Spp][PlantClass] = {}
                    if( not PlantForm in SpecialForm[Spp][PlantClass] ): SpecialForm[Spp][PlantClass][PlantForm] = {}
                    if( not CrownClass in SpecialForm[Spp][PlantClass][PlantForm] ):
                        SpecialForm[Spp][PlantClass][PlantForm][CrownClass] = (NoBranch,NoWhorl,BrBase,BrAngle,LowX,LowY,HighX,HighY,BaseUp,TopUp,StemCol,BrCol,
                                                                               Fol1,Fol2,SampHt,SampRat,SampRad,Scale)
                else:                                   # otherwide handle normal species treeforms
                    if( not Spp in SppForm ): SppForm[Spp] = {}
                    if( not PlantClass in SppForm[Spp] ): SppForm[Spp][PlantClass] = {}
                    if( not PlantForm in SppForm[Spp][PlantClass] ): SppForm[Spp][PlantClass][PlantForm] = {}
                    if( not CrownClass in SppForm[Spp][PlantClass][PlantForm] ):
                        SppForm[Spp][PlantClass][PlantForm][CrownClass] = (NoBranch,NoWhorl,BrBase,BrAngle,LowX,LowY,HighX,HighY,BaseUp,TopUp,StemCol,BrCol,
                                                                           Fol1,Fol2,SampHt,SampRat,SampRad,Scale)
        TRF.close()
        return( SpecialForm, SppForm )

    def SVS_Write_Header( self, OUT ):
        OUT.write( ";Stand Visualization System\n" )
        OUT.write( ";Plant form definition file\n" )
        OUT.write( ";File produced by SVS version: 3.24\n" )
        OUT.write( ";\n" )
        OUT.write( ";DO NOT EDIT THIS FILE BY HAND!!!!!\n" )
        OUT.write( ";SVS does not perform rigorous validation of the parameters\n" )
        OUT.write( ";in this file so any mistakes could cause SVS to crash\n" )
        OUT.write( ";NOTE: Allow 15 characters for the species code!!\n" )
        OUT.write( ";              | this marks column 16\n" )
        OUT.write( ";              |\n" )
        OUT.write( ";Species       | Plant  Crown  Plant     #        #     Branch  Branch  Low pt  Low pt  High pt  High pt    Base    Top    Stem   Branch  Foliage  " )
        OUT.write( "Foliage  Sample    Sample    Sample    Scale\n" )
        OUT.write( "; code         | class  class  form   branches  whorls   base   angle     X       Y        X        Y      uptilt  uptilt  color  " )
        OUT.write( "color   color 1  color 2  height    cratio    cradius\n" )
        OUT.write( ";---------------------------------------------------------------------------------------------------------------------------------" )
        OUT.write( "-------------------------------------------------------------\n" )
        #OUT.write( "--                99     99      0       190      13     0.00    49     1.00    0.15     0.83     0.55    -2.40    5.00     10     10      " )
        #OUT.write( "18       18      120.0     0.50      13.00      0\n" )
        #OUT.write( "@flame.eob        99     99     15       100       0     0.00    47     1.00    0.10     0.60     0.80     0.05    0.05      0      1       " )
        #OUT.write( "0        0       38.0     0.40      18.00      0\n" )

    def SVS_Write_TreeForm( self, OUT, Spp, PlantClass, CrownClass, PlantForm, NoBranch, NoWhorl, BranchBase, BranchAngle, LowX, LowY, HighX, HighY, BaseUp, TopUp,
                            StemColor, BranchColor, Foliage1, Foliage2, SampHt, SampRat, SampRad, Scale ):

        #print( "{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}\n".format(Spp,
        #           PlantClass,CrownClass,PlantForm,NoBranch,NoWhorl,BranchBase,BranchAngle,LowX,LowY,HighX,HighY,BaseUp,TopUp,StemColor,BranchColor,Foliage1,Foliage2,
        #           SampHt,SampRat,SampRad,Scale) )
        #Species = "{:15d}".format(Spp)
        OUT.write( "{:15s}{:>5s}{:>7s}{:>7s}{:>10s}{:>8s}{:>9s}{:>6s}{:>9s}{:>8s}{:>9s}{:>9s}{:>9s}{:>8s}{:>7s}{:>7s}{:>8s}{:>9s}{:>11s}{:>9s}{:>11s}{:>7s}\n".format(
                   str(Spp),PlantClass,CrownClass,PlantForm,NoBranch,NoWhorl,BranchBase,BranchAngle,LowX,LowY,HighX,HighY,BaseUp,TopUp,StemColor,BranchColor,
                   Foliage1,Foliage2,SampHt,SampRat,SampRad,Scale) )

    def SVS_Write_TreeFormFile( self, TreeFormFile, SpecialForm, SppForm ):
        TFM = open( TreeFormFile, 'w' )
        SVS_Write_Header( TFM )
        for TF in sorted(SpecialForm.keys()):
            for PC in sorted(SpecialForm[TF].keys()):
                for PF in sorted(SpecialForm[TF][PC].keys()):
                    for CC in sorted(SpecialForm[TF][PC][PF]):
                        (NoBranch,NoWhorl,BrBase,BrAngle,LowX,LowY,HighX,HighY,BaseUp,TopUp,StemCol,BrCol,Fol1,Fol2,SampHt,SampRat,SampRad,Scale) = SpecialForm[TF][PC][PF][CC]
                        SVS_Write_TreeForm( TFM, TF, PC, CC, PF, NoBranch, NoWhorl, BrBase, BrAngle, LowX, LowY, HighX, HighY, BaseUp, TopUp, StemCol, BrCol,
                                            Fol1, Fol2, SampHt, SampRat, SampRad, Scale )
        nRec = 0
        for TF in sorted(SppForm.keys()):
            for PC in sorted(SppForm[TF].keys()):
                for PF in sorted(SppForm[TF][PC].keys()):
                    for CC in sorted(SppForm[TF][PC][PF]):
                        (NoBranch,NoWhorl,BrBase,BrAngle,LowX,LowY,HighX,HighY,BaseUp,TopUp,StemCol,BrCol,Fol1,Fol2,SampHt,SampRat,SampRad,Scale) = SppForm[TF][PC][PF][CC]
                        SVS_Write_TreeForm( TFM, TF, PC, CC, PF, NoBranch, NoWhorl, BrBase, BrAngle, LowX, LowY, HighX, HighY, BaseUp, TopUp, StemCol, BrCol,
                                            Fol1, Fol2, SampHt, SampRat, SampRad, Scale )
                        nRec += 1
        TFM.close()
        print( "Should have written {} lines".format(nRec) )

#########################################################################################
# Class to abstract interface to Windows SVS program
#########################################################################################
class SVS:
    '''Class to abstract Stand Visualization System (SVS) files'''
    def __init__( self ):
        '''Contructor/Initialize SVS class'''
        pass

    def __del__( self ):
        '''Destructor for SVS class'''
        pass

#########################################################################################
# Class to provide interface to creating sVS visualizations
#########################################################################################
class StandViz:
    """class to handle interface to Stand Visualization System (SVS)"""
    def __init__( self, DataSetName ):
        self.ResolutionLow = '1024x768'
        self.ResolutionHigh = '2048x1536'
        self.FocalLength = 150
        self.RandSeed = -5000
        self.RangePole = ''
        self.Season = 'Summer'
        self.SpeciesCase = 'Upper'
        self.TPAScale = 1
        self.TreeFormFile = '%s/NRCS.trf' % (OWNPATH)
        #self.PaletteFile = '%s/TIR-BLUE.pal' % (OWNPATH)
        self.PaletteFile = None
        self.ViewpointDist = 1000
        self.ViewpointElev = 1000
        self.SvsExe = '{}\\SVS\\winsvs.exe'.format(OWNPATH)
        self.Data = ForestData( DataSetName )
        self.SVF = None                         # variable for file handle object
        random.seed( self.RandSeed )            # initialize random seed generator to common starting point

    def BMP_To_PNG( bmpfilename, pngfilename ):
        """"""
        cmdline = '{}/bmp2png.exe -E "{}"'.format(OWNPATH, bmpfilename)

    def Compute_Offset( self, Bearing, Distance ):
        """compute random offset distance"""
        xoff = random.uniform( 0, Distance )
        yoff = random.uniform( 0, Distance )
        if( (Bearing >= 0) & (Bearing <= 90 ) ):
            return(0+xoff, 0+yoff )
        elif( (Bearing > 90) & (Bearing <= 180) ):
            return(0+xoff, 0-yoff )
        elif( (Bearing > 180) & (Bearing <= 270) ):
            return(0-xoff, 0-yoff )
        elif( (Bearing > 270) & (Bearing <= 360) ):
            return(0-xoff, 0+yoff )
        else:
            return(0, 0)

    def CSV_Load_File( self, infilename ):
        """Load .cvs format file into class structures"""
        print( 'Loading "{}"'.format(infilename) )
        IN = open( infilename, 'r' )
        standname = None
        laststand = None
        for L in IN:
            if( L[0:9] == 'Site/Plot' ): continue
            col = L.split( ',' )
            standname = col[0]
            year = int(col[1])
            #if( re.search( '-', standname ) != None ):
            #    loc = re.search( '-', standname )
            #    year = int(standname[loc.start()+1:])
            #    standname = standname[0:loc.start()]
            #print standname, year
            if( (standname != None ) & (standname != laststand) ): self.Data.Stand[standname] = StandData( standname )
            (treeno, species, dbh, ht, live, status, cclass, tpa) = \
                (int(col[2]), col[3], float(col[4]), float(col[5]), col[6], col[7], int(float(col[8])), float(col[9]))
            if( OPT['d'] ):
                if( dbh > 10.0 ): dbh *= 1.25
                if( dbh > 15.0 ): dbh *= 1.50
            for t in range( 1, int( math.ceil( tpa ))+1, 1 ):
                ntree = len( self.Data.Stand[standname].Tree ) + 1
                self.Data.Stand[standname].Tree[ntree] = TreeData( species, TreeNumber=treeno )
                self.Data.Stand[standname].Tree[ntree].Year[year] = MeasurementData( dbh, ht, '', 1, live, status, cclass )
            laststand = standname
        IN.close()

    def CSV_Write_File( self, cvsfilename ):
        """write data from D to csv file format"""
        self.SVF = open( cvsfilename, 'w' )
        self.SVF.write( 'Site/Plot, Age, Tree#, OrigTree#, Species, Dia, Ht, Live/Dead, Status, Condition, TPA, CR, Crad, ' )
        self.SVF.write( 'BrokenHt, BrokenOff, Bearing, DMR, LeanAngle, RootWad, X, Y\n' )
        stands = self.Data.Stand.keys()
        stands.sort()
        for s in stands:
            print( s )
            ymin = 9999
            ymax = 0
            trees = self.Data.Stand[s].Tree.keys()
            for t in trees:
                years = self.Data.Stand[s].Tree[t].Year.keys()
                for y in years:
                    if( y < ymin ): ymin = y
                    if( y > ymax ): ymax = y
            years = range( ymin, ymax+1, 5 )
            for y in years:
                trees = self.Data.Stand[s].Tree.keys()
                trees.sort()
                for t in trees:
                    if( not self.Data.Stand[s].Tree.has_key(t) ): continue
                    if( not self.Data.Stand[s].Tree[t].Year.has_key(y) ): continue
                    (species, dbh, ht, tpa, treeno, live, cclass, status) = ( self.Data.Stand[s].Tree[t].Species,
                            self.Data.Stand[s].Tree[t].Year[y].DBH, self.Data.Stand[s].Tree[t].Year[y].Height,
                            self.Data.Stand[s].Tree[t].Year[y].TPA, self.Data.Stand[s].Tree[t].TreeNumber,
                            self.Data.Stand[s].Tree[t].Year[y].Live, self.Data.Stand[s].Tree[t].Year[y].Condition,
                            self.Data.Stand[s].Tree[t].Year[y].Status )
                    self.SVF.write( '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s\n' % \
                               (s, y, t, treeno, species, dbh, ht, live, status, cclass, tpa ) )
        self.SVF.close()

    def Determin_Excel_Format( self, ExcelFileName ):
        # determine if StandViz, StandVizExtended, OTIS, or other
        pass
    

    def Excel_Load_Data( self, ExcelFilename ):
        """Load tree records from Excel file"""
        pass

    def Generate_Clumped( self, nClump, nAttract ):
        """generate tree coordinates in a clumped pattern with nClump clumps and nAttract attraction to the clumps"""
        #nClump = 10
        #nAttract = 10
        CLUMP = {}
        for c in range( 1, nClump+1, 1 ):
            CLUMP[c] = (random.uniform( 0, 208.71), random.uniform( 0, 208.71) )
        stands = self.Data.Stand.keys()
        for s in stands:
            trees = self.Data.Stand[s].Tree.keys()
            trees.sort()
            for t in trees:
                c = int(random.uniform( 1, nClump ))
                b = random.uniform( 0, 360 )
                v = random.uniform( 0, nAttract )
                (vx, vy) = self.Compute_Offset( b, v )
                (x, y) = CLUMP[c]
                if( (x + vx) > 208.71 ): x -= vx
                else: x += vx
                if( (y + vy) > 208.71 ): y -= vy
                else: y += vy
                if( y < 0 ): y *= -1.0
                self.Data.Stand[s].Tree[t].X = x
                self.Data.Stand[s].Tree[t].Y = y


    def Generate_Fixed( self ):
        """"""
        stands = self.Data.Stand.keys()
        stands.sort()
        for s in stands:
            trees = self.Data.Stand[s].Tree.keys()
            trees.sort()
            stpa = 0
            for t in trees:
                years = self.Data.Stand[s].Tree[t].Year.keys()
                years.sort()
                y = years[0]
                tpa = self.Data.Stand[s].Tree[t].Year[y].TPA
                raw_input( "Stand={}, Tree={}, Year={}, TPA={}".format(s, t, y, tpa) )


    def Generate_Random( self ):
        """generate tree coordinates using uniform random numbers for x,y locations"""
        print( 'Generating Random coordinates' )
        stands = self.Data.Stand.keys()
        stands.sort()
        for s in stands:
            trees = self.Data.Stand[s].Tree.keys()
            trees.sort()
            for t in trees:
                self.Data.Stand[s].Tree[t].X = random.uniform( 0, 208.71 )
                self.Data.Stand[s].Tree[t].Y = random.uniform( 0, 208.71 )

    def Generate_Uniform( self, Spacing=None, Variation=0.75 ):
        """generate tree coordinates using uniform grid pattern with even spacing"""
        stands = self.Data.Stand.keys()
        stands.sort()
        for s in stands:
            trees = self.Data.Stand[s].Tree.keys()
            tpa = 0.0
            for t in trees:
                years = self.Data.Stand[s].Tree[t].Year.keys()
                tpa += self.Data.Stand[s].Tree[t].Year[years[0]].TPA
            #print tpa
            if( Spacing==None ):
                #tpa = self.Data.Stand[s].Year[15].TPA
                rows = math.floor( math.sqrt( 43560 ) / math.sqrt( 43560 / math.ceil( tpa ) ) )
                spacing = 208.71 / rows
            else:
                spacing = Spacing
            print( tpa, spacing )
            GRID = {}
            x = 5
            y = 5
            trees = self.Data.Stand[s].Tree.keys()
            trees.sort()
            for t in trees:
                if( x > 208.71 ):
                    x = 5
                    y += spacing
                if( y > 208.71 ):
                    x = 5
                    y = 5
                GRID[t] = (x,y)
                x += spacing
            for t in trees:
                g = int(random.uniform( 1, tpa))
                var = random.uniform( 0, Variation)
                ang = random.uniform( 0, 360 )
                (ox,oy) = self.Compute_Offset( ang, var)
                #print ox, oy
                (x,y) = GRID[g]
                self.Data.Stand[s].Tree[t].X = x+ox
                self.Data.Stand[s].Tree[t].Y = y+oy

    def CopyFile( self, fromfile, tofile ):
        """"""
        try:
            f1 = open( fromfile, 'rb' )
            f2 = open( tofile, 'wb' )
            while 1:
                line = f1.readline()
                if not line: break
                f2.write( line )
            f1.close()
            f2.close()
        except IOError: pass

    def Update_WinSVS_IniFile( self, Update ):
        """"""
        inifile = '{}\winsvs.ini'.format(OWNPATH)
        bckfile = '{}\winsvs-SvsAddin-backup.ini'.format(OWNPATH)
        if( not os.path.exists( bckfile ) ):
            self.CopyFile( inifile, bckfile )       # make backup copy of file
        CFG = ConfigParser.RawConfigParser()
        CFG.read( '%s\winsvs.ini' % (OWNPATH) )
        if( Update == 'Restore' ):                  # restore to default values
            CFG.set( 'Preferences', 'DefaultLayout', 'perspective')
            CFG.set( 'Preferences', 'Imagesave', '1024, 768, 0, 0, 1')
            CFG.set( 'Preferences', 'FormFilter', '0')
            CFG.set( 'Preferences', 'FormRows', '4')
            CFG.set( 'Preferences', 'FormCols', '8')
        elif( Update == 'NormalRes' ):                  # setup for HiRes image capture
            CFG.set( 'Preferences', 'DefaultLayout', 'perspective')
            CFG.set( 'Preferences', 'Imagesave', '1024, 768, 0, 0, 1')
        elif( Update == 'HighRes' ):                  # setup for HiRes image capture
            CFG.set( 'Preferences', 'DefaultLayout', 'perspective')
            CFG.set( 'Preferences', 'Imagesave', '2048, 1536, 0, 0, 1')
        elif( Update == 'Legend' ):                 # setup for Legend capture
            CFG.set( 'Preferences', 'Imagesave', '1024, 768, 0, 0, 1')
            CFG.set( 'Preferences', 'DefaultLayout', 'TreeFormLegend')
            CFG.set( 'Preferences', 'FormFilter', '1')
            CFG.set( 'Preferences', 'FormRows', '2')
            CFG.set( 'Preferences', 'FormCols', '12')
        else:                                       # unkown, error
            print( 'Error, unknonw Winsvs.ini update requested: %s' % (Update) )
        CFGOUT = open( '{}\winsvs.ini'.format(OWNPATH), 'w' )
        CFG.write( CFGOUT )
        CFGOUT.close()

    def SVS_Show_Files( self, dirname ):
        """"""
        ProjectName = self.Data.Name
        stands = self.Data.Stand.keys()
        stands.sort()
        batfilename = '{}\\RunSvs.bat'.format(dirname)
        print( batfilename )
        if( os.path.exists( batfilename ) ): os.unlink( batfilename )
        BAT = open( batfilename, 'w' )
        for s in stands:
            ymin = 9999
            ymax = 0
            trees = self.Data.Stand[s].Tree.keys()
            for t in trees:
                years = self.Data.Stand[s].Tree[t].Year.keys()
                for y in years:
                    if( y < ymin ): ymin = y
                    if( y > ymax ): ymax = y
            years = range( ymin, ymax+1, 5 )
            for y in years:
                SvsFilename = '{}/svsfiles/{}/{}-{}.svs'.format(dirname, ProjectName, s, y)
                SvsTitle = '{} : {}-{}'.format(ProjectName, s, y)
                SvsCmdLine = '-E{} -D{} -L{} -T"{}"'.format(self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle)
                BAT.write( '"{}" {} "{}"\n'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
        BAT.close()
        #raw_input( "paused after RunSvs.bat created")
        os.system( '"{}"'.format(batfilename) )
        #os.unlink( batfilename )

    def SVS_Create_Bitmaps( self, dirname ):
        """"""
        ProjectName = self.Data.Name
        stands = self.Data.Stand.keys()
        stands.sort()
        #BAT = open( 'SvsBat.bat', 'w' )
        for s in stands:
            ymin = 9999
            ymax = 0
            trees = self.Data.Stand[s].Tree.keys()
            for t in trees:
                years = self.Data.Stand[s].Tree[t].Year.keys()
                for y in years:
                    if( y < ymin ): ymin = y
                    if( y > ymax ): ymax = y
            years = range( ymin, ymax+1, 5 )
            for y in years:
                SvsFilename = '{}svsfiles/{}/{}-{}.svs'.format(dirname, ProjectName, s, y)
                BmpFilename = '{}svsfiles/{}/{}-{}.bmp'.format(dirname, ProjectName, s, y)
                HrBmpFilename = '{}/svsfiles/{}/{}-{}_HiRes.bmp'.format(dirname, ProjectName, s, y)
                SvsTitle = '{} : {}-{}'.format(ProjectName, s, y)
                SvsCmdLine = '-E{} -D{} -L{} -T"{}" -C"{}"'.format(self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, BmpFilename)
                #BAT.write( '{} {} {}\n'.format((self.SvsExe, SvsCmdLine, SvsFilename) )
                os.system( '{} {} {}'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                os.system( '{}\\bmp2png.exe -E "{}"'.format(OWNPATH, BmpFilename) )
                # update winini file
                #self.Update_WinSVS_IniFile( 'HighRes' )
                SvsCmdLine = '-E{} -D{} -L{} -T"{}" -C"{}"'.format(self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, HrBmpFilename)
                #BAT.write( '{} {} {}\n'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                os.system( '{} {} {}'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                #self.Update_WinSVS_IniFile( 'NormalRes' )
                #BAT.write( '{} -E "{}"\n'.format('.\SvsAddin\\bmp2png', BmpFilename) )
                os.system( '{}\\bmp2png.exe -E "{}"'.format(OWNPATH, HrBmpFilename) )
                #self.Update_WinSVS_IniFile( 'Restore' )
                #os.system( '.\SvsAddin\\bmp2png -E svsfiles/{}/*.bmp'.format(ProjectName) )
        #BAT.close()
        #os.system( 'SvsBat.bat' )


    def SVS_Create_Files( self, dirname ):
        """ """
        ProjectName = self.Data.Name
        #if( not os.path.exists( 'svsfiles' ) ): os.mkdir( 'svsfiles' )
        target_path = '{}\\svsfiles'.format(dirname)
        if( not os.path.exists( target_path ) ):
            print( 'did not find {}, create it'.format(target_path) )
            os.mkdir( target_path )
        #else: print '{} exists already'.format(target_path)
        target_path = '{}\\svsfiles\\{}'.format(dirname, ProjectName)
        print( 'Creating SVS files in {}'.format(target_path) )
        if( not os.path.exists( target_path ) ):
            print( 'did not find {}, create it'.format(target_path) )
            os.mkdir( target_path )
        #else: print '{} exists already' % (target_path)
        # copy .pal and .trf file to directory
        #print 'Copying {}\\TIR-BLUE.pal to %s'.format(OWNPATH)
        self.CopyFile( '{}\\SVS\\TIR-BLUE.pal'.format(OWNPATH), '{}\\TIR-BLUE.pal'.format(target_path) )
        self.CopyFile( '{}\\SVS\\TIR-SvAddin.trf'.format(OWNPATH), '{}\\TIR-SvAddin.trf'.format(target_path) )
        #self.SVF.write( '#PALETTE {}\SvsAddin\TIR-BLUE.pal\n'.format(OWNPATH) )
        #self.SVF.write( '#TREEFORM {}\SvsAddin\TIR-SvAddin.trf\n'.format(OWNPATH) )
        stands = self.Data.Stand.keys()
        stands.sort()
        print( stands )
        #raw_input( "paused" )
        for s in stands:
            ymin = 9999
            ymax = 0
            trees = self.Data.Stand[s].Tree.keys()
            for t in trees:
                years = self.Data.Stand[s].Tree[t].Year.keys()
                for y in years:
                    if( y < ymin ): ymin = y
                    if( y > ymax ): ymax = y
            years = range( ymin, ymax+1, 5 )
            #print 'years=%s' % (years)
            for y in years:
                SvsFilename = '{}/svsfiles/{}/{}-{}.svs'.format(dirname, ProjectName, s, y)
                print( SvsFilename )
                self.SVF = open( SvsFilename, 'w' )
                self.SVS_Write_Header()
                trees = self.Data.Stand[s].Tree.keys()
                trees.sort()
                for t in trees:
                    if( not self.Data.Stand[s].Tree.has_key(t) ): continue
                    if( not self.Data.Stand[s].Tree[t].Year.has_key(y) ): continue
                    (species, dbh, ht, tpa, treeno, live, cclass, status) = ( self.Data.Stand[s].Tree[t].Species,
                             self.Data.Stand[s].Tree[t].Year[y].DBH, self.Data.Stand[s].Tree[t].Year[y].Height,
                             self.Data.Stand[s].Tree[t].Year[y].TPA, self.Data.Stand[s].Tree[t].TreeNumber,
                             self.Data.Stand[s].Tree[t].Year[y].Live, self.Data.Stand[s].Tree[t].Year[y].Condition,
                             self.Data.Stand[s].Tree[t].Year[y].Status )
                    #print species, dbh, ht
                    (cw, cr) = (self.Data.Stand[s].Tree[t].Year[y].CrownRadius, self.Data.Stand[s].Tree[t].Year[y].CrownRatio)
                    (X, Y, z) = (self.Data.Stand[s].Tree[t].X, self.Data.Stand[s].Tree[t].Y, 0)
                    if( cclass in ( 1, 'D' ) ):
                        cclass = 1
                    elif( cclass in ( 2, 'C' ) ):
                        cclass = 2
                    elif( cclass in ( 3, 'I' ) ):
                        cclass = 3
                    elif( cclass in ( 4, 'S' ) ):
                        cclass = 4
                    else:
                        cclass = 1

                    pclass = 0
                    if( live in ( 1, '', 'l', 'live' ) ):
                        pclass = 0
                    elif( live in ( 'dying' ) ):
                        pclass = 1
                    elif( live in ( 'd', 'dead' ) ):
                        pclass = 2
                        if( cr == 0 ): cr = 0.5
                    elif( live in ( 's', 'stump' ) ):
                        pclass = 0
                        status = 'stump'
                    bearing = 0
                    edia = 0
                    lang = 0
                    mark = 0
                    #X = int(treeno)*2
                    #Y = int(treeno)*1.5
                    #status = 1
                    # Spp, T#, PClass, CClass, Status, DBH, Ht, LAng, Bearing, EDia, CW1, CR1, CW2, CR2, CW3, CR3, CW4, CR4, TPA, Mark, X, Y, Z
                    if( status in ( 1, '', 's', 'standing' ) ):
                        if( live in ( 'd', 'dead' ) ):
                            self.SVS_Write_Tree_Dead( species,treeno,pclass,cclass,status,dbh,ht,lang,bearing,edia,cw,cr,tpa,mark,X,Y,z )
                        else:
                            #print species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, X, Y, z
                            self.SVS_Write_Tree_Live( species,treeno,pclass,cclass,status,dbh,ht,lang,bearing,edia,cw,cr,tpa,mark,X,Y,z )
                    elif( status in ( 'b', 'broken' ) ):
                        self.SVS_Write_Tree_Broken( species,treeno,pclass,cclass,status,dbh,ht,lang,bearing,edia,cw,cr,tpa,mark,X,Y,z )
                    elif( status in ( 'brokentop' ) ):
                        self.SVS_Write_Tree_BrokenTop( species,treeno,pclass,cclass,status,dbh,ht,lang,bearing,edia,cw,cr,tpa,mark,X,Y,z )
                    elif( status in ( 'deadtop' ) ):
                        self.SVS_Write_Tree_DeadTop( species,treeno,pclass,cclass,status,dbh,ht,lang,bearing,edia,cw,cr,tpa,mark,X,Y,z )
                    elif( status in ( 'd', 'down' ) ):
                        self.SVS_Write_Tree_Down( species,treeno,pclass,cclass,status,dbh,ht,lang,bearing,edia,cw,cr,tpa,mark,X,Y,z )
                    elif( status in ( 'stump' ) ):
                        self.SVS_Write_Tree_Stump( species,treeno,pclass,cclass,status,dbh,ht,lang,bearing,edia,cw,cr,tpa,mark,X,Y,z )
                self.SVS_Write_Footer()
                self.SVF.close()

    def SVS_Write_Tree( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        fmtstr = '%-15s %-14s %4s %4s %4s %6s %6s %5s %5s %5s %6.2f %4.2f %6.2f %4.2f %6.2f %4.2f %6.2f %4.2f %6s %3s %9.2f %8.2f %8s\n'
        self.SVF.write( fmtstr % (species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, cw, cr, cw, cr,
                   cw, cr, tpa, mark, x, y, z ) )

    def SVS_Write_Tree_Broken( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        #self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )

    def SVS_Write_Tree_BrokenTop( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        #self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )

    def SVS_Write_Tree_Dead( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        #self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )

    def SVS_Write_Tree_DeadTop( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        #self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )

    def SVS_Write_Tree_Down( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        status = 1
        lang = 90
        z = 1
        self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )

    def SVS_Write_Tree_Live( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        if( edia == '' ): edia = 0
        if( mark == '' ): mark = 0
        status = 1
        self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )
        #if( DMR > 0 ):
        #    self.SVS_Write_Tree_DMR()

    def SVS_Write_Tree_Stump( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        status = 1
        height = 3
        self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )

    def SVS_Write_Tree_RootWad( self, species, treeno, pclass, cclass, status, dbh, ht,lang, bearing, edia, cw, cr, tpa, mark, x, y, z ):
        """"""
        #self.SVS_Write_Tree( species, treeno, pclass, cclass, status, dbh, ht, lang, bearing, edia, cw, cr, tpa, mark, x, y, z )

    def SVS_Generate_Pictures( self ):
        """ """
        # get list of plots
        # for p in plots:
        SvsFileName = '{}/{}.svs'.format(SVSTempPath, PlotName)
        BmpFileName = '{}/{}.bmp'.format(SvsTempPath, PlotName)
        BmpLegendFile = '{}/{}.legend'.format(SvsTempPath, PlotName)
        SVS_Create_File( SvsFileName )
        SvsTitle = '{} : {}-{}'.format(Workbook_Name, Worksheet_Name, PlotName)
        SvsCmdLine = '-E {} -D {} -L {} -T "{}" -C"{}"'.format(ViewpointElev, ViewpointDist, FocalLength, SvsTitle, BmpFileName)
        SVS_Run( SvsCmdLine, SvsFileName )
        SVS_Generate_Legend( SVSFileName, '{}.bmp'.format(BmpLegendFile) )

    def SVS_Run( self, SvsOpts, SvsFilename ):
        """ """
        BAT = open( 'SvBat.bat', 'w' )
        if( SvsOpts == '' ):
            CmdLine = '{}'.format(self.SvsExe)
            BAT.write( '{}\n'.format(CmdLine) )
        else:
            CmdLine = '{} {} {}'.format(self.SvsExe, SvsOpts, SvsFilename)
        BAT.close()

    def SVS_Webpage_Create( self, dirname ):
        """"""
        ProjectName = self.Data.Name
        stands = self.Data.Stand.keys()
        stands.sort()
        batfilename = '{}\\RunSvs.bat'.format(dirname)
        if( os.path.exists( batfilename ) ): os.unlink( batfilename )
        print( 'creating {}'.format(batfilename) )
        BAT = open( batfilename, 'w' )
        BAT.write( 'del {}\\svsfiles\\{}\\*.png\n'.format(dirname, ProjectName) )
        htmlfilename = '{}/svsfiles/{}/{}.html'.format(dirname, ProjectName, ProjectName)
        HTML = open( htmlfilename, 'w' )
        HTML.write( '<html>\n<head>\n<title>{}</title>\n'.format(ProjectName) )
        HTML.write( '<meta http-equiv="Content-Type" content="text/html; chartset=iso-8859-1" />\n' )
        HTML.write( '</head>\n' )
        HTML.write( '<a name="Top">\n' )
        HTML.write( '<center><h1>Visualizations for {}</h1></center>'.format(ProjectName) )
        HTML.write( '<p>The {} project contains the following plots.  Click on the plot name below to jump to the '.format(ProjectName) )
        HTML.write( 'visualization for that plot.  Click on the main visualization image to load a higher resolution version of the image.' )
        for s in stands:
            ymin = 9999
            ymax = 0
            trees = self.Data.Stand[s].Tree.keys()
            for t in trees:
                years = self.Data.Stand[s].Tree[t].Year.keys()
                for y in years:
                    if( y < ymin ): ymin = y
                    if( y > ymax ): ymax = y
            years = range( ymin, ymax+1, 5 )
            for y in years:
                SvsFilename = '{}/svsfiles/{}/{}-{}.svs'.format(dirname, ProjectName, s, y)
                BmpFilename = '{}/svsfiles/{}/{}-{}.bmp'.format(dirname, ProjectName, s, y)
                HrBmpFilename = '{}/svsfiles/{}/{}-{}_HiRes.bmp'.format(dirname, ProjectName, s, y)
                SvsTitle = '{} : {}-{}' % (ProjectName, s, y)
                SvsCmdLine = '-E{} -D{} -L{} -T"{}" -C"{}"'.format(self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, BmpFilename)
                #self.Update_WinSVS_IniFile( 'NormalRes' )
                BAT.write( '"{}\\python.exe" "{}\\Update_WinSVS_IniFile.py" NormalRes\n'.format(OWNPATH, OWNPATH) )
                #print '"{}" {} "{}"'.format(self.SvsExe, SvsCmdLine, SvsFilename)
                #os.system( '"{}" {} "{}"'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                BAT.write( '"{}" {} "{}"\n'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                #print '"{}\\bmd2png.exe" -E "{}"'.format(OWNPATH, BmpFilename)
                #os.system( '"{}\\bmp2png.exe" -E "{}"'.format(OWNPATH, BmpFilename) )
                BAT.write( '"{}\\bmp2png.exe" -E "{}"\n'.format(OWNPATH, BmpFilename) )
                #self.Update_WinSVS_IniFile( 'HighRes' )
                BAT.write( '"{}\\python.exe" "{}\\Update_WinSVS_IniFile.py" HighRes\n'.format(OWNPATH, OWNPATH) )
                SvsCmdLine = '-E{} -D{} -L{} -T"{}" -C"{}"'.format(self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, HrBmpFilename)
                #os.system( '"{}" {} "{}"'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                BAT.write( '"{}" {} "{}"\n'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                #os.system( '"{}\\bmp2png.exe" -E "{}"'.format( OWNPATH, HrBmpFilename) )
                BAT.write( '"{}\\bmp2png.exe" -E "{}"\n'.format( OWNPATH, HrBmpFilename) )
                #self.Update_WinSVS_IniFile( 'Legend' )
                BAT.write( '"%s\\python.exe" "{}\\Update_WinSVS_IniFile.py" Legend\n'.format(OWNPATH, OWNPATH) )
                BmpLegendFilename = '{}/svsfiles/{}/{}-{}_legend.bmp'.format(dirname, ProjectName, s, y)
                SvsCmdLine = '-C"{}"'.format(BmpLegendFilename)
                #os.system( '"{}" {} "{}"'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                BAT.write( '"{}" {} "{}"\n'.format(self.SvsExe, SvsCmdLine, SvsFilename) )
                #os.system( '"{}\\bmp2png.exe" -E "{}"'.format( OWNPATH, BmpLegendFilename) )
                BAT.write( '"{}\\bmp2png.exe" -E "{}"\n'.format( OWNPATH, BmpLegendFilename) )
                #self.Update_WinSVS_IniFile( 'Restore' )
                BAT.write( '"{}\\python.exe" "{}\\Update_WinSVS_IniFile.py" Restore\n'.format(OWNPATH, OWNPATH) )
        HTML.write( '<ul>\n' )
        BAT.close()
        raw_input( 'RunSvs.bat created...pause before run' )
        os.system( '"%s"'.format(batfilename) )
        #os.unlink( batfilename )
        for s in stands:
            for y in years:
                HTML.write( '<li><a href="#{}-{}">{}-{}</a></li>\n'.format(s, y, s, y) )
        HTML.write( '</ul><hr>\n' )
        if( OPT['z'] ):
            HTML.write( 'Download zip archive of this webpage: <a href="{}.zip">{}.zip</a>\n<hr>\n'.format(ProjectName, ProjectName) )
        for s in stands:
            for y in years:
                PngFilename = '{}-{}.png'.format(s, y)
                PngHighResFilename = '{}-{}_HiRes.png'.format(s, y)
                PngLegendFilename = '{}-{}_legend.png'.format(s, y)
                HTML.write( '<a name="{}-{}"><h1>File: {} - {}-{}</h1></a>\n'.format(s, y, ProjectName, s, y))
                HTML.write( '<a href="{}"><img src="{}" boarder="0"></a>\n'.format(PngHighResFilename, PngFilename))
                HTML.write( '<img src="{}">\n'.format(PngLegendFilename) )
                HTML.write( '<p><a href="#Top">Top</a><hr>\n')
        HTML.write( '<p>Visualizations generated: {}.'.format(time.asctime()) )
        HTML.write( '</html>\n' )
        HTML.close()
        os.system( '"{}\\svsfiles\\{}\\{}.html"'.format(dirname, ProjectName, ProjectName) )

    def SVS_Webpage_Zip( self, dirname ):
        """ """
        ProjectName = self.Data.Name
        zipcmd = '{}\\zip.exe {}.zip *.html *.pal *.trf *.svs *.png'.format(OWNPATH, ProjectName)
        os.chdir( '{}\\svsfiles\{}'.format(dirname, ProjectName) )
        os.system( zipcmd )
        os.chdir( OWNPATH )

    def SVS_Write_Footer( self ):
        self.SVF.write( '; SVS file created by StandViz.py {}\n'.format(__file_version__[11:len(__file_version__)-1]) )

    def SVS_Write_Header( self ):
        """ """
        self.SVF.write( '#PLOTORIGIN  0.00 0.00\n' )
        self.SVF.write( '#PLOTSIZE    208.71 208.71\n' )
        self.SVF.write( '#FORMAT      2\n' )
        self.SVF.write( '#UNITS       ENGLISH\n' )
        #self.SVF.write( '#TREEFORM PLANTS-SvAddin.trf\n' )
        #self.SVF.write( '#PALETTE %s\SvsAddin\TIR-BLUE.pal\n' % (OWNPATH) )
        #self.SVF.write( '#TREEFORM %s\SvsAddin\TIR-SvAddin.trf\n' % (OWNPATH) )
        #self.SVF.write( '#PALETTE TIR-BLUE.pal\n' )
        self.SVF.write( '#TREEFORM inst\\bin\\SVS\\NRCS.trf\n' )
        self.SVF.write( ';              |                                                           Crown       Crown       Crown       Crown\n' )
        self.SVF.write( ';              |    Plant        Class   Tree                            end  RadiusRatio RadiusRatio RadiusRatio ' )
        self.SVF.write( 'RadiusRatio' )
        self.SVF.write( 'Expans Mark  X Coor-  Y Coor-\n' )
        self.SVF.write( '; Species      |     ID       |Plnt|Crwn|Stat|  dbh |height| lang| fang| dia |   1  |  1 |   2  |  2 |   3  |  3 |   ' )
        self.SVF.write( '4  |  4 |' )
        self.SVF.write( 'Factor|Code| dinate | dinate |     Z  \n' )
        self.SVF.write( ';--------------------------------------------------------------------------------------------------------------------' )
        self.SVF.write( '---------' )
        self.SVF.write( '-----------------------------------------\n' )
        #self.SVF.write( ';PITA           01-01             0    0    1    1.5     10     0     0     0    3.3 0.85    3.3 0.85    3.3 0.85    3.3 ' )
        #self.SVF.write( '0.85    1.0   0       4.8   201.91        0\n' )
        #self.SVF.write( ';PITA           01-02             0    0    1    2.5     14     0     0     0   4.62 0.88   4.62 0.88   4.62 0.88   4.62 ' )
        #self.SVF.write( '0.88    1.0   0      12.8   201.91        0\n' )

def Test_Excel_Format( filename ):
    FileFormat = 'Unknown'
    XL = win32com.client.Dispatch( 'Excel.Application' )
    XL.Visible = False
    WB = XL.Workbooks.Open( filename )
    nSheets = WB.Sheets.Count
    #print '%s: found %s worksheets' % (filename, nSheets)
    bHaveConfig = False
    for sheet in WB.Sheets:
        #print 'Worksheet: %s' % (sheet.Name)
        XLS = XL.Worksheets( sheet.Name )
        if( sheet.Name == 'Configuration' ):
            A1 = XLS.Range( "A1:A1" ).Value
            B1 = XLS.Range( "B1:B1" ).Value
            if( (A1 == 'Worksheet') & (B1 == 'TreeformFile') ): bHaveConfig = True
        else:
            A1 = XLS.Range( "A1:A1" ).Value
            B1 = XLS.Range( "B1:B1" ).Value
            C1 = XLS.Range( "C1:C1" ).Value
            J1 = XLS.Range( "J1:J1" ).Value
            if( (A1 == 'Site/Plot') & (B1 == 'Tree#') & (C1 == 'Species') ): FileFormat = 'SvsAddin'
            elif( re.search( 'Stand Development Report:', J1 ) != None ): FileFormat = 'TIR'
    #print 'File Format: %s' % (FileFormat)
    XL.ActiveWorkbook.Close(False)
    WB = None
    del WB
    XL.Quit()
    XL = None
    del XL
    return( FileFormat )


########################
# Begin Execution Here #
########################

if( __name__ == "__main__" ):
    main()

"""

Miscelaneous notes and programming ideas

#DataSet = 'LCP_TQF_SI70'
#DataSetNew = '%s-Fixed' % (DataSet)
#infile = '%s.csv' % (DataSet)
#outfile = '%s.csv' % (DataSetNew)
## load data
#D = ForestData( DataSet )
#IN = open( infile, 'r' )
#standname = None
#laststand = None
#for L in IN:
#    if( L[0:9] == "Site/Plot" ): continue
#    col = L.split(',')
#    standname = col[0]
#    if( (standname != None) & (standname != laststand) ): D.Stand[standname] = StandData( standname )
#    (treeno, species, dbh, ht, live, status, cclass, tpa) = (int(col[1]), col[2], float(col[3]), float(col[4]), col[5], col[6], int(float(col[7])), float(col[8]))
#    for t in range( 1, int( math.ceil( tpa ))+1, 1 ):
#        ntree = len( D.Stand[standname].Tree ) + 1
#        D.Stand[standname].Tree[ntree] = TreeData( species, dbh, ht, 1, live, status, cclass, TreeNumber=treeno )
#    laststand = standname
#IN.close()

# generate coordinates

# load .csv file of tree lists and generate spatial coordinates to simulate thinnings

# gen coordinates in 208.7103 x 208.7103 grid  - 208.71 = SQRT( 43560 )
# 1 acre = 43560 sq ft
#   0       thin         is 484.95 tpa
#  75 sq ft thin results in 296.29 tpa
# 100 sq ft thin results in 221.23 tpa
# 125 sq ft thin results in 146.18 tpa
# tpa = 43560 / ( x * y )
# spacing = SQRT( 43560 / tpa )
# rows = SQRT( 43560 ) / spacing
# rows = math.floor( math.sqrt( 43560 ) / math.sqrt( 43560/math.ceil(tpa) ) )
# need 22x22 grid for 484 trees / acre

#GRID = {}

# random.random()  # rand float between 0 and 1
# RS = random.getstate()
# random.setstate( RS)
# random.seed( longint )

# if need python modules
# python.exe pip -m pip install module_name

TQI = 1-2
Tree[tqi][drank]

"""
