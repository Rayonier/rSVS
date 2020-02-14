#!/usr/bin/python
#
# Filename:...: StandViz.py
# Description.: Stand level visualzation package using the Stand Visualization System (SVS)
# Author......: James B. McCarter
# Copyright...: 2020, Rayonier, Inc
# Requirements:
#

import argparse             # ArgumentParser()
import math
import os                   # os.path.split(), os.path.splitext()
import pandas as pd         # Pandas DataFrame, .read_csv()
import platform             # platform.system()
import re                   # re.search()
import sys
import time

(__file_version__, __file_date__) = ( '$Revision: 1.0.1 $', '$Date: 2020/02/12 07:19:00 $' )
(_MyPath, _MyFile, _MyOS) = (os.path.split(sys.argv[0])[0], os.path.split(sys.argv[0])[1], platform.system())
(_MyVersion, _MyDate) = ( __file_version__.split()[1], '{} - {}'.format(__file_date__.split()[1], __file_date__.split()[2]) )

#
# PySvsAddin.py - stand visualization adding implementation in Python
#

#import ConfigParser, math, os, random, re, sys, time, win32com.client, pythoncom, _winreg

##################################
# Begin Global Data Declarations #
##################################

#OWNPATH = os.getcwd()
(OWNPATH, file) = os.path.split(sys.argv[0])
#global OWNPATH
VERBOSE = 0

########################################################
# Begin Main Program - implement command line inferface
########################################################
def main():     # implement main scope for handling of command line execution of script
    global DEBUG, NOTIFY, VERBOSE
    try:
        (DEBUG, NOTIFY, VERBOSE) = (False, False, False)
        # ..C.E.GHIJKLMNOPQR.TU...YZ .b..e.g.ijk.m.o.q.st..wxyz 0123456789
        SARG = argparse.ArgumentParser( add_help=False, usage=" %(prog)s [-B|-W|-S|-X] [-c|-f|-r|-u] [dDhnv] [-a #] [-l #] [-p #] [-w worksheet] file [file [...]]\n" +
                                                              "\t%(prog)s [-v] [-A [FIA|NRCS]] [-F]" )

        SARGO = SARG.add_argument_group( "Output arguments" )
        SARGO.add_argument( "-B", action="store_true", help="output to Bitmap (capture .bmp, convert to .png)" )
        SARGO.add_argument( "-D", action="store_true", help="Debug output" )
        SARGO.add_argument( "-W", action="store_true", help="output to HTML (create .png, generate .html page)" )
        SARGO.add_argument( "-S", action="store_true", help="output to SVS (default)" )
        SARGO.add_argument( "-X", action="store_true", help="output to eXcel (.csv file)" )

        SARGC = SARG.add_argument_group( "Coordiante arguments" )
        SARGC.add_argument( "-c", action="store_true", help="generate clumped coordinates" )
        SARGC.add_argument( "-f", action="store_true", help="generate fixed coordinates" )
        SARGC.add_argument( "-r", action="store_true", help="generate random coordinates" )
        SARGC.add_argument( "-u", action="store_true", help="generate Uniform coordinates" )
        SARGC.add_argument( "-a", action="store", nargs=1, metavar="#", help="rAndomness factor (0=perf rows, 0.4-0.8=plantation, <.8=clumps)" )
        SARGC.add_argument( "-l", action="store", nargs=1, metavar="#", help="cLumpiness factor (default 0.75)" )
        #SARGC.add_argument( "-w", action="store", nargs=1, metavar="name", help="worksheet name for Excel input" )
        SARGC.add_argument( "-p", action="store", nargs=1, metavar="#", help="clumP ratio (n clumps = (0.01-0.5)*TPA)" )

        SARGG = SARG.add_argument_group( "General arguments" )
        SARGG.add_argument( "-d", action="store_true", help="scale Diameter: dbh>10*1.25; dbh>15*1.50" )
        SARGG.add_argument( "-h", action="store_true", help="display help" )
        SARGG.add_argument( "-n", action="store_true", help="Notify progress in DOS window" )
        SARGG.add_argument( "-v", action="store_true", help="Verbose output" )
        #SARGG.add_argument( "-z", action="store_true", help="zip file for transfer" )
        #SARGG.add_argument( "-m", action="store_true", help="Mechanical (row) thinning (only for Fixed coordinates)" )
        #SARGG.add_argument( "-t", action="store_true", help="TIR format files" )

        SARGT = SARG.add_argument_group( "Treeform arguments" )
        SARGT.add_argument( "-A", action="store", nargs=1, metavar="TRFile", help="Audit treeform file versions rSVS_Species.csv" )
        SARGT.add_argument( "-F", action="store_true", help="create FIA.TRF from NRCS.trf" )
        SARGT.add_argument( "-V", action="store_true", help="Validate rSVS_Species.csv file" )


        SARG.add_argument( "FILELIST", nargs="*", help="Files [File [...]]")
        SOPT = SARG.parse_args()
        nFile = len( SOPT.FILELIST )

        if( SOPT.D ): DEBUG = 1
        if( SOPT.v ): VERBOSE = 1
        if( SOPT.n ): NOTIFY = 1

        if( SOPT.A ):           # audit TreeForm file against rSVS_Species.csv
            print( "Performing audit of {}.trf against rSVS_Species.csv".format(SOPT.A[0]) )
            SppFile = "../bin/rSVS_Species.csv"
            SPP = pd.read_csv( SppFile )
            print( "Read {} lines from {}".format(len(SPP.index), SppFile) )
            SppCodes = SOPT.A[0]
            TreeFormFile = "../bin/SVS/{}.trf".format(SppCodes)
            (SpecialForm, SppForm) = SVS_LoadTreeFormFile( TreeFormFile )
            print( "Read {} lines from {}".format(len(SppForm.keys()), TreeFormFile) )
            AUDIT = {}
            for S in SPP.itertuples():
                (FIA, NRCS, Genus, Species) = (S.FIA, S.NRCS, S.Genus, S.Species)
                if( SppCodes == 'NRCS' ):
                    if( not NRCS in AUDIT ): AUDIT[NRCS] = 1
                elif( SppCodes == 'FIA' ):
                    if( not FIA in AUDIT ): AUDIT[FIA] = 1
            (Have, Missing) = (0, 0)
            for S in sorted(SppForm.keys()):
                if( not S in AUDIT ): Missing += 1
                else: Have += 1
            print( "{}: Has {}, Missing {}".format(TreeFormFile, Have, Missing) )
            sys.exit("performed audit")

        if( SOPT.F ):       # create FIA.trf from NCRS.trf
            print( "Creating FIA.trf..." )
            SppFile = "../bin/rSVS_Species.csv"
            SPP = pd.read_csv( SppFile )
            print( "Read {} lines from {}".format(len(SPP.index), SppFile) )
            TreeFormFile = "../bin/SVS/{}.trf".format('NRCS')
            (SpecialForm, SppForm) = SVS_LoadTreeFormFile( TreeFormFile )
            print( "Read {} lines from {}".format(len(SppForm.keys()), TreeFormFile) )
            # loop through SppForm.keys() and change species to FIA # from SPP
            NewTreeFormFile = 'TEST.trf'
            # SVS_Write_TreeFormFile( NewTreeFormFile, SpecialForm, SppForm )
            sys.exit( "created FIA.trf" )

        if( SOPT.V ):
            print( "Performing audit of rSVS_Species.csv file" )
            SppFile = "../bin/rSVS_Species.csv"
            SPP = pd.read_csv( SppFile )
            sys.exit( "audited rSVS_Species.csv" )

        #nfiles = len(cmdline)
        if( (nFile==0) | SOPT.h ):
            SARG.print_help()
            sys.exit( "help printed" )

        if( (SOPT.c==0) & (SOPT.f==0) & (SOPT.r==0) & (SOPT.u==0) ): SOPT.r = True   # random is default coordinate generation
        if( (SOPT.B==0) & (SOPT.W==0) & (SOPT.S==0) & (SOPT.X==0) ): SOPT.S = 1   # SVS is default output

        if( NOTIFY ): print( 'StandViz.py - Python implementation of Stand Visualization Addin for Excel' )
        if( DEBUG ): print( 'len(cmdline)=%d, cmdline="%s", OPT=%s' % (nfiles, cmdline, OPT) )


        for FILE in FILELIST:
        #for f in cmdline:
            #D = {}              # create data dictionary
            if( DEBUG ): print( 'File: %s' % (f) )
            (dirname, filename) = os.path.split( f )
            if( DEBUG ): print( 'dirname=%s, filename=%s' % (dirname, filename) )
            (basename, ext) = os.path.splitext( filename )
            if( DEBUG ): print( 'dirname=%s, filename=%s, basename=%s, ext=%s' % (dirname, filename, basename, ext) )
            #DataSet = 'None'
            if( re.search( '.csv', filename ) != None ):        # create DataSet name from filename
                DataSet = re.sub( '.csv', '', filename )
            elif( re.search( '.xlsx', filename ) != None ):
                DataSet = re.sub( '.xlsx', '', filename )
            elif( re.search( '.xls', filename ) != None ):
                DataSet = re.sub( '.xls', '', filename )

            #print 'DataSet = %s' % (DataSet)
            #SVS = StandViz( DataSet )                 # create class/dataset for input file

            # if extension is .xls or xlsx file then need to determine if we are a SvsAddin format or TIR format file
            #if( ext in ['.xls', '.xlsx' ] ):            # test eXcel file for type
            #    FileFormat = 'Excel'
            #    FileFormat = Test_Excel_Format( f )
            #    #raw_input( "Paused: After Test_Excel_Format(): %s is %s" % (f, FileFormat) )
            #    if( FileFormat == 'TIR' ):      # TIR format excel files, run PyTIRData.py to extract
            #        BAT = open( '%s\\RunPyTIRData.bat' % (OWNPATH), 'w' )
            #        cmd = '"%s\\python.exe" "%s\\PyTIRData.py" "%s"\n' % (OWNPATH, OWNPATH, f)
            #        BAT.write( cmd )
            #        BAT.close()
            #        cmd = '%s\\RunPyTIRData.bat' % (OWNPATH)
            #        #raw_input( 'Running: %s' % (cmd) )
            #        os.system( cmd )
            #        #os.unlink( cmd )
            #        (standname, treatment, age) = filename.split('_')
            #        #dirname = '%s\\%s_%s' % (dirname, standname, treatment)
            #        csvfilename = '%s\\%s_%s.csv' % (dirname, standname, treatment)
            #        SVS.CSV_Load_File( csvfilename )
            #    else:       # unknown excel file format
            #        sys.exit()
            #elif( ext in [ '.csv' ] ):
            #    #SVS.CSV_Load_File( f )                      # load the data from .cvs file
            #    if( OPT['t'] ):
            #        D = TIR_Load_Data( DataSet, f )         # load TIR format file
            #        TIR_Expand_Treelist( D, SVS )           # expand the treelist
            #    else:
            #        SVS.CSV_Load_File( f )                  # load SvsAddin format .csv file
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
        ERRFILE.write( "  File '{}', line {}, in {}\n    %s\n".format(file, line, name, code) )     # write filename, source code line, error name, and error code
        MyTrace = MyTrace.tb_next                                       # step to next level of call stack trace
    ERRFILE.write( "errorobj: {}\n".format(errorobj) )                  # write error object and arguments for call
    ERRFILE.write( "Calling Argument Vector: {}\n".format(args) )       # write calling arguments
    ERRFILE.close()                                                     # close text file with error stack trace
    os.system( "notepad.exe {}".format(errorfilename) )                 # display error log file with notepad.exe

def SVS_LoadTreeFormFile( TreeFormFile ):
    SppForm = {}
    SpecialForm = {}
    SpecialList = [ '--', '@flame.eob', 'CAR', 'CRANEBOOM', 'CRANETOWER', 'CONIFER', 'DEFAULT', 'DMBROOM', 'HARDWOOD', 'MARKER', 'MISTBROOM', 'R6CLUMP',
                    'R6SHRUB', 'R6SNAG', 'RANGEPOLE', 'ROCK', 'ROOTWAD', 'SEEDLING', 'SHRUB', 'SNAG', 'SNAG2', 'SNAG3', 'SNAG4', 'SNAG5', 'TETRAHEDRON', 'TRUCK' ]
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

def SVS_Write_TreeFormFile( TreeFormFile, SpecialForm, SppForm ):
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


def SVS_Write_Header( OUT ):
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

def SVS_Write_TreeForm( OUT, Spp, PlantClass, CrownClass, PlantForm, NoBranch, NoWhorl, BranchBase, BranchAngle, LowX, LowY, HighX, HighY, BaseUp, TopUp,
                        StemColor, BranchColor, Foliage1, Foliage2, SampHt, SampRat, SampRad, Scale ):
    OUT.write( "{:15s}{:>5s}{:>7s}{:>7s}{:>10s}{:>8s}{:>9s}{:>6s}{:>9s}{:>8s}{:>9s}{:>9s}{:>9s}{:>8s}{:>7s}{:>7s}{:>8s}{:>9s}{:>11s}{:>9s}{:>11s}{:>7s}\n".format(Spp,
               PlantClass,CrownClass,PlantForm,NoBranch,NoWhorl,BranchBase,BranchAngle,LowX,LowY,HighX,HighY,BaseUp,TopUp,StemColor,BranchColor,Foliage1,Foliage2,
               SampHt,SampRat,SampRad,Scale) )

# -B# clump ratio # clumps = (0.01 - 0.5) * TPA
# -G# clumpiness factor = 1.5-1.4*clumpiness factor)*clump spacing
# -R# Randomness Factor (0 = perfect rows and columns; 0.4-0.8 aproximate planted stands; > 0.8 some clumps of 2-3 trees

# not used: egijknoqy

###########################
# Begin Class Definitions #
###########################

#
# TreeData, StandData, and ForestData classes store tree information stored by Forest (dataset), stand, and tree
#

#
# The TreeData class holds tree information (does not change) and measurements (change with time)
#
class MeasurementData:
    """class to hold tree measurement information organized by year"""
    #D = ForestData( 'ForestName' )
    #D.Stand['StandName'] = StandData( 'StandName' )
    # D.STand['StandName'].Plot[0] = PlotData( 0, Size=1.0 )
    #D.Stand['StandName'].Plot[0].Tree[1] = TreeData( Species, TreeNumber, X, Y )
    #D.Stand['StandName'].Plot[0].Tree[1].Year[1] = (DBH, Height, CrownRatio, TPA, Live, Status, Condition, ... )
    def __init__( self, DBH=None, Height=None, CrownRatio=None, TPA=None, Live=None, Status=None,
                  Condition=None, Bearing=None, BrokenHeight=None, BrokenOffset=None,
                  CrownRadius=None, DMR=None, LeanAngle=None, RootWad=None ):
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

class TreeData:
    """class for holding tree information"""
    # D = ForestData( 'Forest' )
    # D.Stand['StandName'] = StandData( 'StandName' )
    # D.STand['StandName'].Plot[0] = PlotData( 0, Size=1.0 )
    # D.Stand['StandName'].Plot[0].Tree[1] = TreeData( Species, TreeNumber, X, Y )
    def __init__( self, Species=None, TreeNumber=None, X=None, Y=None ):
        self.Species = Species              # species
        self.TreeNumber = TreeNumber        # tree numbers
        self.X = X                          # tree X coordinate
        self.Y = Y                          # tree Y coordinate
        self.Year = {}                      # dictionary for holding MeasurementData objects

class PlotData:
    """class for holding plot level information within a stand"""
    # D = ForestData( 'ForestName' )
    # D.Stand['StandName'] = StandData( 'StandName' )
    # D.Stand['StandName'].Plot['PlotName'] = PlotData( 'PlotName' )
    def __init__ (self, Name, Size=1.0 ):
        self.Name = Name
        self.Size = Size
        self.Tree = {}                      # dictionary to hold TreddData objects

# StandData class holds tree information by stand
#
class StandData:
    """class for holding stand level information"""
    #D = ForestData( 'DatasetName' )
    #D.Stand[1] = StandData( 'StandName' )
    def __init__( self, Name, Plots=False ):
        self.Name = Name                    # name for stand
        if( Plots ):
            self.Plot = {}                      # dictionary to hold PlotData objects
        else:
            self.Tree = {}                      # dictionary for TreeData objects
            self.Year = {}                      # dictionary to hold stand summary information

#
# ForestData class holds stand information for dataset
#
class ForestData:
    """class for containing forest/data set/project/file level inventory information"""
    #D = ForestData( 'DatasetName' )
    def __init__( self, Name ):
        self.Name = Name                    # name for forest/data set
        self.Stand = {}                     # dictionary for StandData objects

# D = ForestData( "DataName" )
# D.Stand[1] = StandData( ... )
# D.Stand[1].Plot[1] = PlotData()
# D.Stand[1].Plot[1].Tree[1] = TreeData( Species='DF', DBH=5.2, TPA=10.2 )
# for s in D.Stand.keys()
#     for p in D.Stand[s].Plot.keys()
#         for t in D.Stand[s].Plot[p].Tree.keys()
#             (Spp, Dbh, Ht, Cr, TPA) = D.Stand[s].Plot[p].Tree[t]

class SVS:
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
        self.TreeFormFile = '%s/TIR.trf' % (OWNPATH)
        self.PaletteFile = '%s/TIR-BLUE.pal' % (OWNPATH)
        self.ViewpointDist = 1000
        self.ViewpointElev = 1000
        #self.SvsExe = '%s\winsvs.exe' % ('C:\ProgramData\SVS')
        #self.SvsExe = '%s\winsvs.exe' % ('C:\App\TIRViz\SVS')
        self.SvsExe = '%s\\SVS\\winsvs.exe' % (OWNPATH)
        self.Data = ForestData( DataSetName )
        self.SVF = None                         # variable for file handle object
        random.seed( self.RandSeed )            # initialize random seed generator to common starting point

    def BMP_To_PNG( bmpfilename, pngfilename ):
        """"""
        cmdline = '%s/bmp2png.exe -E "%s"' % (OWNPATH, bmpfilename)

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
        print( 'Loading "%s"' % (infilename) )
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
                raw_input( "Stand=%s, Tree=%s, Year=%s, TPA=%s" % (s, t, y, tpa))


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
        inifile = '%s\winsvs.ini' % (OWNPATH)
        bckfile = '%s\winsvs-SvsAddin-backup.ini' % (OWNPATH)
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
        CFGOUT = open( '%s.\winsvs.ini' % (OWNPATH), 'w' )
        CFG.write( CFGOUT )
        CFGOUT.close()

    def SVS_Show_Files( self, dirname ):
        """"""
        ProjectName = self.Data.Name
        stands = self.Data.Stand.keys()
        stands.sort()
        batfilename = '%s\\RunSvs.bat' % (dirname)
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
                SvsFilename = '%s/svsfiles/%s/%s-%s.svs' % (dirname, ProjectName, s, y)
                SvsTitle = '%s : %s-%s' % (ProjectName, s, y)
                SvsCmdLine = '-E%s -D%s -L%s -T"%s"' % (self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle)
                BAT.write( '"%s" %s "%s"\n' % (self.SvsExe, SvsCmdLine, SvsFilename) )
        BAT.close()
        #raw_input( "paused after RunSvs.bat created")
        os.system( '"%s"' % (batfilename) )
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
                SvsFilename = '%s/svsfiles/%s/%s-%s.svs' % (dirname, ProjectName, s, y)
                BmpFilename = '%s/svsfiles/%s/%s-%s.bmp' % (dirname, ProjectName, s, y)
                HrBmpFilename = '%s/svsfiles/%s/%s-%s_HiRes.bmp' % (dirname, ProjectName, s, y)
                SvsTitle = '%s : %s-%s' % (ProjectName, s, y)
                SvsCmdLine = '-E%s -D%s -L%s -T"%s" -C"%s"' % (self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, BmpFilename)
                #BAT.write( '%s %s %s\n' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                os.system( '%s %s %s' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                os.system( '%s\\bmp2png.exe -E "%s"' % (OWNPATH, BmpFilename) )
                # update winini file
                #self.Update_WinSVS_IniFile( 'HighRes' )
                SvsCmdLine = '-E%s -D%s -L%s -T"%s" -C"%s"' % (self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, HrBmpFilename)
                #BAT.write( '%s %s %s\n' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                os.system( '%s %s %s' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                #self.Update_WinSVS_IniFile( 'NormalRes' )
                #BAT.write( '%s -E "%s"\n' % ('.\SvsAddin\\bmp2png', BmpFilename) )
                os.system( '%s\\bmp2png.exe -E "%s"' % (OWNPATH, HrBmpFilename) )
                #self.Update_WinSVS_IniFile( 'Restore' )
                #os.system( '.\SvsAddin\\bmp2png -E svsfiles/%s/*.bmp' % (ProjectName) )
        #BAT.close()
        #os.system( 'SvsBat.bat' )


    def SVS_Create_Files( self, dirname ):
        """ """
        ProjectName = self.Data.Name
        #if( not os.path.exists( 'svsfiles' ) ): os.mkdir( 'svsfiles' )
        target_path = '%s\\svsfiles' % (dirname)
        if( not os.path.exists( target_path ) ):
            print( 'did not find %s, create it' % (target_path) )
            os.mkdir( target_path )
        #else: print '%s exists already' % (target_path)
        target_path = '%s\\svsfiles\\%s' % (dirname, ProjectName)
        print( 'Creating SVS files in %s' % (target_path) )
        if( not os.path.exists( target_path ) ):
            print( 'did not find %s, create it' % (target_path) )
            os.mkdir( target_path )
        #else: print '%s exists already' % (target_path)
        # copy .pal and .trf file to directory
        #print 'Copying %s\\TIR-BLUE.pal to %s' % (OWNPATH)
        self.CopyFile( '%s\\SVS\\TIR-BLUE.pal' % (OWNPATH), '%s\\TIR-BLUE.pal' % (target_path) )
        self.CopyFile( '%s\\SVS\\TIR-SvAddin.trf' % (OWNPATH), '%s\\TIR-SvAddin.trf' % (target_path) )
        #self.SVF.write( '#PALETTE %s\SvsAddin\TIR-BLUE.pal\n' % (OWNPATH) )
        #self.SVF.write( '#TREEFORM %s\SvsAddin\TIR-SvAddin.trf\n' % (OWNPATH) )
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
                SvsFilename = '%s/svsfiles/%s/%s-%s.svs' % (dirname, ProjectName, s, y)
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
        SvsFileName = '%s/%s.svs' % (SVSTempPath, PlotName)
        BmpFileName = '%s/%s.bmp' % (SvsTempPath, PlotName)
        BmpLegendFile = '%s/%s.legend' % (SvsTempPath, PlotName)
        SVS_Create_File( SvsFileName )
        SvsTitle = '%s : % s- %s' % (Workbook_Name, Worksheet_Name, PlotName)
        SvsCmdLine = '-E %s -D %s -L %s -T "%s" -C"%s"' % (ViewpointElev, ViewpointDist, FocalLength, SvsTitle, BmpFileName)
        SVS_Run( SvsCmdLine, SvsFileName )
        SVS_Generate_Legend( SVSFileName, '%s.bmp' % (BmpLegendFile) )

    def SVS_Run( self, SvsOpts, SvsFilename ):
        """ """
        BAT = open( 'SvBat.bat', 'w' )
        if( SvsOpts == '' ):
            CmdLine = '%s' % (self.SvsExe)
            BAT.write( '%s\n' % (CmdLine) )
        else:
            CmdLine = '%s %s %s' % (self.SvsExe, SvsOpts, SvsFilename)
        BAT.close()

    def SVS_Webpage_Create( self, dirname ):
        """"""
        ProjectName = self.Data.Name
        stands = self.Data.Stand.keys()
        stands.sort()
        batfilename = '%s\\RunSvs.bat' % (dirname)
        if( os.path.exists( batfilename ) ): os.unlink( batfilename )
        print( 'creating %s' % (batfilename) )
        BAT = open( batfilename, 'w' )
        BAT.write( 'del %s\\svsfiles\\%s\\*.png\n' % (dirname, ProjectName) )
        htmlfilename = '%s/svsfiles/%s/%s.html' % (dirname, ProjectName, ProjectName)
        HTML = open( htmlfilename, 'w' )
        HTML.write( '<html>\n<head>\n<title>%s</title>\n' % (ProjectName) )
        HTML.write( '<meta http-equiv="Content-Type" content="text/html; chartset=iso-8859-1" />\n' )
        HTML.write( '</head>\n' )
        HTML.write( '<a name="Top">\n' )
        HTML.write( '<center><h1>Visualizations for %s</h1></center>' % (ProjectName) )
        HTML.write( '<p>The %s project contains the following plots.  Click on the plot name below to jump to the ' % (ProjectName) )
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
                SvsFilename = '%s/svsfiles/%s/%s-%s.svs' % (dirname, ProjectName, s, y)
                BmpFilename = '%s/svsfiles/%s/%s-%s.bmp' % (dirname, ProjectName, s, y)
                HrBmpFilename = '%s/svsfiles/%s/%s-%s_HiRes.bmp' % (dirname, ProjectName, s, y)
                SvsTitle = '%s : %s-%s' % (ProjectName, s, y)
                SvsCmdLine = '-E%s -D%s -L%s -T"%s" -C"%s"' % (self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, BmpFilename)
                #self.Update_WinSVS_IniFile( 'NormalRes' )
                BAT.write( '"%s\\python.exe" "%s\\Update_WinSVS_IniFile.py" NormalRes\n' % (OWNPATH, OWNPATH) )
                #print '"%s" %s "%s"' % (self.SvsExe, SvsCmdLine, SvsFilename)
                #os.system( '"%s" %s "%s"' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                BAT.write( '"%s" %s "%s"\n' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                #print '"%s\\bmd2png.exe" -E "%s"' % (OWNPATH, BmpFilename)
                #os.system( '"%s\\bmp2png.exe" -E "%s"' % (OWNPATH, BmpFilename) )
                BAT.write( '"%s\\bmp2png.exe" -E "%s"\n' % (OWNPATH, BmpFilename) )
                #self.Update_WinSVS_IniFile( 'HighRes' )
                BAT.write( '"%s\\python.exe" "%s\\Update_WinSVS_IniFile.py" HighRes\n' % (OWNPATH, OWNPATH) )
                SvsCmdLine = '-E%s -D%s -L%s -T"%s" -C"%s"' % (self.ViewpointElev, self.ViewpointDist, self.FocalLength, SvsTitle, HrBmpFilename)
                #os.system( '"%s" %s "%s"' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                BAT.write( '"%s" %s "%s"\n' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                #os.system( '"%s\\bmp2png.exe" -E "%s"' % ( OWNPATH, HrBmpFilename) )
                BAT.write( '"%s\\bmp2png.exe" -E "%s"\n' % ( OWNPATH, HrBmpFilename) )
                #self.Update_WinSVS_IniFile( 'Legend' )
                BAT.write( '"%s\\python.exe" "%s\\Update_WinSVS_IniFile.py" Legend\n' % (OWNPATH, OWNPATH) )
                BmpLegendFilename = '%s/svsfiles/%s/%s-%s_legend.bmp' % (dirname, ProjectName, s, y)
                SvsCmdLine = '-C"%s"' % (BmpLegendFilename)
                #os.system( '"%s" %s "%s"' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                BAT.write( '"%s" %s "%s"\n' % (self.SvsExe, SvsCmdLine, SvsFilename) )
                #os.system( '"%s\\bmp2png.exe" -E "%s"' % ( OWNPATH, BmpLegendFilename) )
                BAT.write( '"%s\\bmp2png.exe" -E "%s"\n' % ( OWNPATH, BmpLegendFilename) )
                #self.Update_WinSVS_IniFile( 'Restore' )
                BAT.write( '"%s\\python.exe" "%s\\Update_WinSVS_IniFile.py" Restore\n' % (OWNPATH, OWNPATH) )
        HTML.write( '<ul>\n' )
        BAT.close()
        raw_input( 'RunSvs.bat created...pause before run' )
        os.system( '"%s"' % (batfilename) )
        #os.unlink( batfilename )
        for s in stands:
            for y in years:
                HTML.write( '<li><a href="#%s-%s">%s-%s</a></li>\n' % (s, y, s, y) )
        HTML.write( '</ul><hr>\n' )
        if( OPT['z'] ):
            HTML.write( 'Download zip archive of this webpage: <a href="%s.zip">%s.zip</a>\n<hr>\n' % (ProjectName, ProjectName) )
        for s in stands:
            for y in years:
                PngFilename = '%s-%s.png' % (s, y)
                PngHighResFilename = '%s-%s_HiRes.png' % (s, y)
                PngLegendFilename = '%s-%s_legend.png' % (s, y)
                HTML.write( '<a name="%s-%s"><h1>File: %s - %s-%s</h1></a>\n' % (s, y, ProjectName, s, y))
                HTML.write( '<a href="%s"><img src="%s" boarder="0"></a>\n' % (PngHighResFilename, PngFilename))
                HTML.write( '<img src="%s">\n' % (PngLegendFilename) )
                HTML.write( '<p><a href="#Top">Top</a><hr>\n')
        HTML.write( '<p>Visualizations generated: %s.' % (time.asctime()) )
        HTML.write( '</html>\n' )
        HTML.close()
        os.system( '"%s\\svsfiles\\%s\\%s.html"' % (dirname, ProjectName, ProjectName) )

    def SVS_Webpage_Zip( self, dirname ):
        """ """
        ProjectName = self.Data.Name
        zipcmd = '%s\\zip.exe %s.zip *.html *.pal *.trf *.svs *.png' % (OWNPATH, ProjectName)
        os.chdir( '%s\\svsfiles\%s' % (dirname, ProjectName) )
        os.system( zipcmd )
        os.chdir( OWNPATH )

    def SVS_Write_Footer( self ):
        self.SVF.write( '; SVS file created by PyStandViz 1.0\n' )

    def SVS_Write_Header( self ):
        """ """
        self.SVF.write( '#PLOTORIGIN  0.00 0.00\n' )
        self.SVF.write( '#PLOTSIZE    208.71 208.71\n' )
        self.SVF.write( '#FORMAT      2\n' )
        self.SVF.write( '#UNITS       ENGLISH\n' )
        #self.SVF.write( '#TREEFORM PLANTS-SvAddin.trf\n' )
        #self.SVF.write( '#PALETTE %s\SvsAddin\TIR-BLUE.pal\n' % (OWNPATH) )
        #self.SVF.write( '#TREEFORM %s\SvsAddin\TIR-SvAddin.trf\n' % (OWNPATH) )
        self.SVF.write( '#PALETTE TIR-BLUE.pal\n' )
        self.SVF.write( '#TREEFORM TIR-SvAddin.trf\n' )
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


def TIR_Expand_Treelist( D, SVS ):
    """expand treelist to whole tree records"""
    #
    stands = D.Stand.keys()
    stands.sort()
    for s in stands:
        mtpa = 0
        ymin = 9999
        ymax = 0
        trees = D.Stand[s].Tree.keys()
        cyears = []
        for t in trees:
            years = D.Stand[s].Tree[t].Year.keys()
            for y in years:
                if( y < ymin ): ymin = y
                if( y > ymax ): ymax = y

        # need maxtpa and tpa after thinning to compute spacing after thinning

        years = range( ymin, ymax+1, 5 )
        if( not SVS.Data.Stand.has_key( s ) ): SVS.Data.Stand[s] = StandData( s )
        trees = D.Stand[s].Tree.keys()
        trees.sort()
        for t in trees:
            y = years[0]
            #print s, y, t
            if( not D.Stand[s].Tree.has_key(t) ): continue
            if( not D.Stand[s].Tree[t].Year.has_key(y) ): continue
            #print 'Looking for %s, %s, %s' % (s,t,y)
            (species, dbh, ht, live, status, cclass, tpa) = (D.Stand[s].Tree[t].Species, D.Stand[s].Tree[t].Year[y].DBH,
                D.Stand[s].Tree[t].Year[y].Height, D.Stand[s].Tree[t].Year[y].Live, D.Stand[s].Tree[t].Year[y].Status,
                D.Stand[s].Tree[t].Year[y].Condition, D.Stand[s].Tree[t].Year[y].TPA)
            for n in range( 1, int( math.ceil( tpa ) )+1, 1 ):
                ntree = len( SVS.Data.Stand[s].Tree ) + 1
                SVS.Data.Stand[s].Tree[ntree] = TreeData( species, TreeNumber=t )
                SVS.Data.Stand[s].Tree[ntree].Year[y] = MeasurementData( dbh, ht, '', 1, live, status, cclass )
        #print 'years=%s' % (years)
        #for y in years:
        #    SVS.Data.Stand[s].Year[y] = D.Stand[s].Year[y]
        for y in years[1:]:
            for t in trees:
                if( not D.Stand[s].Tree.has_key(t) ): continue
                if( not D.Stand[s].Tree[t].Year.has_key(y) ): continue
                #print y, t
                (species, dbh, ht, live, status, cclass, tpa) = (D.Stand[s].Tree[t].Species, D.Stand[s].Tree[t].Year[y].DBH,
                    D.Stand[s].Tree[t].Year[y].Height, D.Stand[s].Tree[t].Year[y].Live, D.Stand[s].Tree[t].Year[y].Status,
                    D.Stand[s].Tree[t].Year[y].Condition, D.Stand[s].Tree[t].Year[y].TPA)
                ntree = 0
                for n in range( 1, int( math.ceil( tpa ) )+1, 1 ):
                    ntree = len( SVS.Data.Stand[s].Tree ) + 1
                    #ntree += n
                    SVS.Data.Stand[s].Tree[ntree] = TreeData( species, TreeNumber=t )
                    SVS.Data.Stand[s].Tree[ntree].Year[y] = MeasurementData( dbh, ht, '', 1, live, status, cclass )

def TIR_Load_Data( DataSet, filename ):
    """"""
    IN = open( filename, 'r' )
    (dirname, file) = os.path.split( filename )
    laststand = None
    TD = {}                         # create temporary dictionary for storing tree records
    D = ForestData( DataSet )       # create data structure for final treelist
    for L in IN:
        if( L[0:9] == 'Site/Plot' ): continue
        col = L.split( ',' )
        standname = col[0]
        year = int(col[1])
        (treeno, species, dbh, ht, live, status, cclass, tpa) = \
            (int(col[2]), col[3], float(col[4]), float(col[5]), col[6], col[7], int(float(col[8])), float(col[9]))
        if( status == '' ): status = 's'
        #print standname, year, treeno, species
        if( not TD.has_key( standname ) ): TD[standname] = {}
        if( not TD[standname].has_key( year ) ): TD[standname][year] = {}
        if( not TD[standname][year].has_key( treeno ) ): TD[standname][year][treeno] = {}
        #print 'status="%s"' % (status)
        if( status == 's' ):
            if( not TD[standname][year][treeno].has_key( 'Live' ) ):
                TD[standname][year][treeno]['Live'] = (species, dbh, ht, tpa, live, status, cclass)
            else: print( 'error, already have tree record' )
        elif( status == 'Cut' ):
            if( not TD[standname][year][treeno].has_key( 'Cut' ) ):
                TD[standname][year][treeno]['Cut'] = (species, dbh, ht, tpa, live, status, cclass)
            else: print( 'error, already have cut record' )
        else:
            print( 'error storing %s' % (L) )
    IN.close()

    stands = TD.keys()
    stands.sort()

    for s in stands:
        years = TD[s].keys()
        years.sort()
        cyears = []     # list of years with thinnings
        TPA = {}
        CTPA = {}
        for y in years:
            raw_input( "Looking at year %s" % (y) )
            #trees = TD[s][y]['Live'].keys()


    for s in stands:
        if( not D.Stand.has_key( s ) ):
            D.Stand[s] = StandData( s)     # initialize data structure for treelist
            D.Stand[s].Cut = {}                     # add dictionary for cut trees
        years = TD[s].keys()
        years.sort()
        print( 'Stand %s has inventory for years %s' % (s, years) )
        for y in years:
            trees = TD[s][y].keys()
            trees.sort()
            trees.reverse()
            print( '%s at %s: %s' % (s, y, trees) )
            for t in trees:
                #print 'Tree %s, Live=%s' % (t, TD[s][y][t]['Live'])
                if( TD[s][y][t].has_key('Live')):
                    (species, dbh, ht, tpa, live, status, cclass) = TD[s][y][t]['Live']
                    if( not D.Stand[s].Tree.has_key(t) ): D.Stand[s].Tree[t] = TreeData( species, TreeNumber=t )
                    D.Stand[s].Tree[t].Year[y] = MeasurementData( dbh, ht, '', tpa, live, status, cclass)
                if( TD[s][y][t].has_key('Cut')):
                    (species, dbh, ht, tpa, live, status, cclass) = TD[s][y][t]['Cut']
                    if( not D.Stand[s].Cut.has_key(t) ): D.Stand[s].Cut[t] = TreeData( species, TreeNumber=t )
                    D.Stand[s].Cut[t].Year[y] = MeasurementData( dbh, ht, '', tpa, live, status, cclass)
    #raw_input("paused")


    #stands = D.Stand.keys()
    #print 'stands=%s' % (stands)
    #stands.sort()
    #for s in stands:
    #    tyears = D.Stand[s].Tree.keys()
    #    tyears.sort()
    #    cyears = D.Stand[s].Cut.keys()
    #    cyears.sort()
    #    print '%s has Trees for %s and Cut for %s' % (s, tyears, cyears)

    #raw_input( "Pause after D creation")

    #sumtpa = sumctpa = 0
    #stands = D.Stand.keys()
    #stands.sort()
    #for s in stands:
    #    ymin = 9999
    #    ymax = 0
    #    trees = D.Stand[s].Tree.keys()
    #    for t in trees:
    #        years = D.Stand[s].Tree[t].Year.keys()
    #        for y in years:
    #            if( y < ymin ): ymin = y
    #            if( y > ymax ): ymax = y
    #        years = range( ymin, ymax+1, 5 )
    #        sumpta = sumctpa = 0
    #    for y in years:
    #        trees = D.Stand[s].Tree.keys()
    #        trees.sort()
    #        for t in trees:
    #            if( not D.Stand[s].Tree.has_key(t) ): continue
    #            if( not D.Stand[s].Tree[t].Year.has_key(y) ): continue
    #            sumtpa += D.Stand[s].Tree[t].Year[y].TPA
    #        ctrees = D.Stand[s].Cut.keys()
    #        ctrees.sort()
    #        for c in ctrees:
    #            if( not D.Stand[s].Cut.has_key(c) ): continue
    #            if( not D.Stand[s].Cut[c].Year.has_key(y) ): continue
    #            sumctpa += D.Stand[s].Cut[c].Year[y].TPA
    #        D.Stand[s].Year[y] = MeasurementData( 0.0, 0.0, '', sumtpa, '', '', '' )
    #        D.Stand[s].Year[y].CutTPA = sumctpa
    #        sumtpa = sumctpa = 0
    return( D )

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

"""
TQI = 1-2
Tree[tqi][drank]
"""
