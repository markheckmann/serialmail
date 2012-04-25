# generate serial mail from .xlsx or .csv file
# parse a document and replace placeholder tags
#
# package depends:
# xlsx, xlsxjars, rJava, stringr, tcltk
# library(xlsx)  
# library(tcltk)
# library(stringr) 

# 1) read template
# 2) get template tags
# 3) read data file
# 4) check if all tags have values
# 5) generate new text strings
# 6) output text to file, prepare email etc.

# get file suffix 
#
get_file_type <- function(file){  
  x <- unlist(str_split(file, "[.]"))
  tail(x, 1)
}


# read in .csv or .xlsx file
# use file selection dialog if no file is supplied 
#
read_data_file <- function(file) 
{
  if (missing(file)) {                                           # open file selection menu if no file argument is supplied
    Filters <- matrix(c("Excel > 2007 files", ".xlsx",
                        "csv files", ".csv"),
                        ncol=2, byrow = TRUE)
    file <- tk_choose.files(filters = Filters, caption="Choose data file",
                            multi=FALSE)    # returns complete path                    
  }
  #basename("r/text.txt")
  filetype <- get_file_type(file)
  if (filetype == "csv") {               # read csv, german version, i.e. ";" as seperator
    x <- read.csv2(file)
  } else if (filetype == "xlsx") {       # read excel spreadsheet 
    x <- read.xlsx(file, 1)
  } else {
     stop("only files of type .csv and .xlxs are allowed")  
  }    
  x
} 

# read in .txt template file
# use file selection dialog if no file is supplied
#
read_template_file <- function(file){
  if (missing(file)) {                                           # open file selection menu if no file argument is supplied
    Filters <- matrix(c("txt file", ".txt"),
                        ncol=2, byrow = TRUE)
    file <- tk_choose.files(filters = Filters, caption="Choose template file",
                            multi=FALSE)    # returns complete path                    
  }  
  con <- file(file, "rt") 
  template <- readLines(con)  
  close(con)
  template                  
} 

# strip of sourrounding tags (<<tag>>)
#
strip_tags <- function(x, intag="<<", outtag=">>"){
  x <- str_replace_all(x, intag, "")
  str_replace_all(x, outtag, "")  
}


# extract all tags used in template file
#
get_template_tags <- function(txt, intag="<<", outtag=">>")
{ 
  pattern <- paste(intag, "(.+?)", outtag, sep="")
  tags <- str_extract_all(txt, pattern)  # retrieve all tags in template 
  unlist(tags)
}


# check if all tags in template file are
# contained in data file
#
check_values_for_tags_in_template <- function(tags, data, intag="<<", outtag=">>"){
  tags.inner <- strip_tags(tags, intag, outtag)
  tags.defined <- tags.inner %in% names(data) 
  if (any(!tags.defined))
    warning("no data for the following tags: ", 
            paste(tags[!tags.defined], collapse=", "), 
            call. = FALSE)
}


# prepare small data frames containing one line of data 
# from data file and names of the tags as colnames
#
make_pairs_list <- function(x, intag="<<", outtag=">>"){
  names(x)  <- paste(intag, names(x), outtag, sep="")  
  lapply(1L:nrow(x), function(i) x[i, ])
}


# parse template file and replace tags by corresponding
# tag data from data file
#
replace_tags_by_value <- function(pair, template, intag="<<", outtag=">>"){
  tags <- get_template_tags(template, intag, outtag)
  avail.tags <- names(pair)
  tags.exist <- tags %in% avail.tags
  txt <- template  
  for (tag in tags[tags.exist]){ 
    i <- match(tag, avail.tags)
    txt <- str_replace_all(txt, tag, pair[1, i])
  }
  txt 
}


# extract mail adress from data pair
#
get_email <- function(pair, intag="<<", outtag=">>"){
  emailcolumn <- paste(intag, "email", outtag, sep="")
  x <- pair[1, emailcolumn]
  if (is.null(x))
    message("email adress is missing")
  x
}


# prepare a single email
#
prepare_single_email <- function(msg, email, subject= "course feedback"){
  create.post.new(msg, subject = subject, address = email)
} 


# prepare all emails using a list of parsed templates and email adresses
# 
prepare_all_emails <- function(msgs, emails, subject= "course feedback"){
  message("Preparing email(s)")  
  msgs.vec <- sapply(msgs, function(x) paste(x, collapse="\n"))   
  emails.vec <- as.vector(unlist(emails))  
  mapply(prepare_single_email, msg=msgs.vec, email=emails.vec, 
         MoreArgs=list(subject=subject))  
  invisible()
}


#' Creates a series of emails from a textfile template.
#'
#' A plain textfile (\code{.txt}) is used as a template to create emails.
#' The textfile may contain any number of arbitrary tags of the form
#' \code{<<tag>>}. The template is parsed and the tags are replaced
#' by values taken from a data file. The data is contained in 
#' an Excel (\code{.xlsx}) or comma seperated value \code{.csv} file. 
#' Based on these two files an email for each line of the data in the data
#' file is created using the default email program. 
#'
#' @section Template textfile: 
#' The template file contains plain text and tags using double braces as indication 
#' for a tag<<tag>>.
#' Any tag name can be used as long as it appears in the header of the data file. 
#' An example:
#'
#' \code{Hello <<name>>, your exam mark is <<mark>>. Best, Mark.}  
#'
#' @section Data file:
#' The data file can be an Excel file. Then the tags without 
#' sourrounding braces are the headers,
#' i.e. contained in the first row of the Excel spreadsheet. Only the
#' first spreadsheet is read.
#'
#' For comma seperated value files the case is the same. The tags 
#' without braces are used as headers. Semicolons are used as seperators
#' An example (\code{.csv}):
#'
#' \code{name; surname; mark}   
#'
#' \code{Mark; Heckmann; 1,3}   
#'
#' \code{Holger; Doering; 1,3}  
#'  
#' @param tfile     Path to template file. Must be a \code{.txt} file. 
#'                  If no file is supplied a file selection dialog
#'                  is opened.
#' @param dfile     Path to data file. Must be a \code{.xlsx} (Excel 2007 or newer) 
#'                  or \code{.csv} (German type csv, i.e. semicolon as separator) file. 
#'                  If no file is supplied a file selection dialog
#'                  is opened.
#' @param output    String. Type of output. Default is to generate an \code{"email"}.
#'                  Currently not used.
#' @param subject   Subject line of email, default is \code{"subject line"}.
#' @param intag     String. The opening signs for tags. The default is \code{<<}. Be careful 
#'                  not to use signs that have a special meaning in regular expressions.
#' @param outtag    String. The closing sign for tags. The default is \code{>>}. Be careful 
#'                  not to use signs that have a special meaning in regular expressions.
#' @author          Mark Heckmann 
#' @export  
#' @examples \dontrun{
#'
#'  # download the files template.txt and data.csv from github dowload page
#'  serialmail("template.txt", "data.csv")
#' }
#' 
# opening sins for parsed tags 
# closing sins for parsed tags
serialmail <- function(tfile, dfile, output="email", subject="subject line",
                       intag="<<", outtag=">>")
{ 
  template <- read_template_file(tfile)                   # 1) read template 
  tags <- get_template_tags(template, intag, outtag)      # 2) get template tags   
  data <- read_data_file(dfile)                           # 3) read data file  
  check_values_for_tags_in_template(tags, data, 
                                    intag, outtag)        # 4) check if all tags have values  
  pairlist <- make_pairs_list(data, intag, outtag)                  
  text.parsed <- lapply(pairlist, replace_tags_by_value, 
                        template, intag, outtag)          # 5) parse text and overwrite tags 
  if (output == "email") {                                # 6) prepare emails etc.    
    emails <- lapply(pairlist, get_email, intag, outtag)  
    prepare_all_emails(text.parsed, emails, subject= subject)
  }
  invisible()
}



