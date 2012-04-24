# serialmail

**serialmail** is an R package to create a series of emails from a text file template `(.txt)` and a data file `(.csv)` or `(.xlsx)`.
                                                  
To install the latest development version from github you can use the `devtools` package.
    
    library(devtools)
    install_github("OpenRepGrid", "markheckmann", dependencies=TRUE) 

You may need to install the following packages before you install `serialmail`:
`xlsx, xlsxjars, rJava, stringr`.

To get started download the files from the [downloads](http://github.com/markheckmann/serialmail/downloads) page.
Start to generate emails from the files: 

    library(serialmail)
    serialmail("template.txt", "data.csv")     
    # or   
    serialmail("template.txt", "data.xlsx") 
    
