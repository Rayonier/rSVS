# rSVS

This rSVS package provides an interface to perform Stand Visualization System (SVS, Robert J. McGaughey, USDA Forest Service, PNW Research Station) visualiations from R.

You can either download the binary package installation file from the dist folder or download the entire repository and build and install local.

## Installing binary distribution

Check the dist folder in the rSVS repository for the most recent version, click on the link and then use the download button.

The easiest way to install the binary distribution is to use the Install button on the Packages tab in RStudio.  On the Install Packages dialog change Install from: to "Packagee Archive File (.zip; .tar.gz)" and browse to where you saved the download .zip file.

Use "Install package(s) from local files" on the Packages menu if you are using RGui.

## Installing from downloaded repository

NOTE: You will need devtools and dependencies to build the package from the repository.

Open RStudio and navigate to the repository directory using Session/Set Working Directory.
```R
require(devtools)
devtools::build(binary=TRUE)
```

## Getting Started with rSVS

Once installed you can get an overview of the package by looking at the package documentation.
```R
help("rSVS-package")
```

NOTE: The first time you run many of the fuctions in rSVS the package will check to see if you have a Python distribution available.  If not, the package will prompt you to allow it to install (unzip) a package internal copy of Python so that it's "back-end" can do it's work.  Example visulizations from the SVS_Example() function will work without this step so that you can demo the package.

