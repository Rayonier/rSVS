% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/rSVS.R
\name{svsdata}
\alias{svsdata}
\title{Create SVS example data}
\usage{
svsdata(
  type = "TBL2SVS",
  species,
  dbh,
  tpa,
  scale = 4,
  shape = 2,
  n = 30,
  dmin = 0.001,
  dmax = 100,
  incr = 0.1,
  hd = 7
)
}
\arguments{
\item{type}{Type of example data to create}

\item{species}{individual or list of species for stand}

\item{dbh}{average diameter breast height for each species}

\item{tpa}{trees per acre for each species}

\item{scale}{weibull scale parameter (default=4)}

\item{shape}{weibull shape parameter (default=2)}

\item{dmin}{minimum diameter for distribution (default 0.001")}

\item{dmax}{maximum diameter for distribution (default 100")}

\item{incr}{diameter increment between dmin and dmax (default 0.1")}

\item{hd}{height to diameter ratio for height dubbing (default 7)}
}
\value{
data frame of list with example data for use with SVS()
}
\description{
Create different kinds of example data understood by this package.
}
\details{
This function will return data data frame in one of the following formats:
\itemize{
  \item StandViz format
  \item TBL2SVS format
}
}
\examples{
SV <- svsdata( species=c(202,263), tpa=c(300,275), dbh=c(7.8,4.7) )
svsdata( datatype=NULL )     # force function to return known data types
}
\author{
James McCarter \email{jim.mccarter@rayonier.com}
}
