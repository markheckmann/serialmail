\name{serialmail}
\alias{serialmail}
\title{Creates a series of emails from a textfile template.}
\usage{
  serialmail(tfile, dfile, output = "email",
    subject = "subject line", intag = "<<", outtag = ">>")
}
\arguments{
  \item{tfile}{Path to template file. Must be a \code{.txt}
  file.  If no file is supplied a file selection dialog is
  opened.}

  \item{dfile}{Path to data file. Must be a \code{.xlsx}
  (Excel 2007 or newer) or \code{.csv} (German type csv,
  i.e. semicolon as separator) file.  If no file is
  supplied a file selection dialog is opened.}

  \item{output}{String. Type of output. Default is to
  generate an \code{"email"}.  Currently not used.}

  \item{subject}{Subject line of email, default is
  \code{"subject line"}.}

  \item{intag}{String. The opening signs for tags. The
  default is \code{<<}. Be careful not to use sign that
  have a special meaning in regular expressions.}

  \item{outtag}{String. The closing sign for tags. The
  default is \code{>>}. Be careful not to use signs that
  have a special meaning in regular expressions.}
}
\description{
  A plain textfile (\code{.txt}) is used as a template to
  create emails. The textfile may contain any number of
  arbitrary tags of the form \code{<<tag>>}. The tenplate
  is parsed and the tags are replaced with values taken
  from a data file. The data is contained in an Excel
  (\code{.xlsx}) or comma seperated value \code{.csv}.
  Based on these files an email for each line of data in
  the data file is created using the default email program.
}
\section{Template textfile}{
  The template file contains plain text and tags wrapped in
  braces <<tag>>. Any tag can be used as long as it appears
  in the data file. An example:

  \code{Hello <<name>>, your exam mark is <<mark>>. Best,
  Mark.}
}

\section{Data file}{
  The data file can be an Excel file. Then the tags without
  sourrounding braces are the headers, i.e. contained in
  the first row of the Excel spreadsheet. Only the first
  spreadsheet is read.

  For comma seperated value files the case is the same. The
  tags without braces are used as headers. Semicolons are
  used as seperators An example (\code{.csv}):

  \code{name; surname; mark}

  \code{Mark; Heckmann; 1,3}

  \code{Holger; Doering; 1,3}
}
\examples{
\dontrun{

 # download the files template.txt and data.csv from github dowload page
 serialmail("template.txt", "data.csv")
}
}
\author{
  Mark Heckmann
}

