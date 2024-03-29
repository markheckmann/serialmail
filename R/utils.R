### slightly modified version of utils::create.post

create.post.new <- function (instructions = character(), description = "post", subject = "", 
    method = getOption("mailer"), address = "the relevant mailing list", 
    ccaddress = getOption("ccaddress", ""), filename = "R.post", 
    info = character()) 
{
    method <- if (is.null(method)) 
        "none"
    else match.arg(method, c("mailto", "mailx", "gnudoit", "none", 
        "ess"))
    open_prog <- if (grepl("-apple-darwin", R.version$platform)) 
        "open"
    else "xdg-open"
    if (method == "mailto") 
        if (!nzchar(Sys.which(open_prog))) {
            browser <- Sys.getenv("R_BROWSER", "")
            if (!nzchar(browser)) {
                warning("cannot find program to open 'mailto:' URIs: reverting to 'method=\"none\"'")
                flush.console()
                Sys.sleep(5)
            }
            else {
                message("Using the browser to open a mailto: URI")
                open_prog <- browser
            }
        }
    body <- c(instructions, info)
    none_method <- function() {
        disclaimer <- paste("# Your mailer is set to \"none\",\n", 
            "# hence we cannot send the, ", description, " directly from R.\n", 
            "# Please copy the ", description, " (after finishing it) to\n", 
            "# your favorite email program and send it to\n#\n", 
            "#       ", address, "\n#\n", "######################################################\n", 
            "\n\n", sep = "")
        cat(c(disclaimer, body), file = filename, sep = "\n")
        cat("The", description, "is being opened for you to edit.\n")
        flush.console()
        file.edit(filename)
        cat("The unsent ", description, " can be found in file ", 
            sQuote(filename), "\n", sep = "")
    }
    if (method == "none") {
        none_method()
    }
    else if (method == "mailx") {
        if (missing(address)) 
            stop("must specify 'address'")
        if (!nzchar(subject)) 
            stop("'subject' is missing")
        if (length(ccaddress) != 1L) 
            stop("'ccaddress' must be of length 1")
        cat(body, file = filename, sep = "\n")
        cat("The", description, "is being opened for you to edit.\n")
        file.edit(filename)
        if (is.character(ccaddress) && nzchar(ccaddress)) {
            cmdargs <- paste("-s", shQuote(subject), "-c", shQuote(ccaddress), 
                shQuote(address), "<", filename, "2>/dev/null")
        }
        else cmdargs <- paste("-s", shQuote(subject), shQuote(address), 
            "<", filename, "2>/dev/null")
        status <- 1L
        answer <- readline(paste("Email the ", description, " now? (yes/no) ", 
            sep = ""))
        answer <- grep("yes", answer, ignore.case = TRUE)
        if (length(answer)) {
            cat("Sending email ...\n")
            status <- system(paste("mailx", cmdargs), , TRUE, 
                TRUE)
            if (status) 
                status <- system(paste("Mail", cmdargs), , TRUE, 
                  TRUE)
            if (status) 
                status <- system(paste("/usr/ucb/mail", cmdargs), 
                  , TRUE, TRUE)
            if (status == 0L) 
                unlink(filename)
            else {
                cat("Sending email failed!\n")
                cat("The unsent", description, "can be found in file", 
                  sQuote(filename), "\n")
            }
        }
        else cat("The unsent", description, "can be found in file", 
            filename, "\n")
    }
    else if (method == "ess") {
        cat(body, sep = "\n")
    }
    else if (method == "gnudoit") {
        cmd <- paste("gnudoit -q '", "(mail nil \"", address, 
            "\")", "(insert \"", paste(body, collapse = "\\n"), 
            "\")", "(search-backward \"Subject:\")", "(end-of-line)'", 
            sep = "")
        system(cmd)
    }
    else if (method == "mailto") {
        if (missing(address)) 
            stop("must specify 'address'")
        if (!nzchar(subject)) 
            subject <- "<<Enter Meaningful Subject>>"
        if (length(ccaddress) != 1L) 
            stop("'ccaddress' must be of length 1")
        #cat("The", description, "is being opened in your default mail program\nfor you to complete and send.\n")
        arg <- paste("mailto:", address, "?subject=", subject, 
            if (is.character(ccaddress) && nzchar(ccaddress)) 
                paste("&cc=", ccaddress, sep = ""), "&body=", 
            paste(body, collapse = "\r\n"), sep = "")
        if (system2(open_prog, shQuote(URLencode(arg)), FALSE, 
            FALSE)) {
            cat("opening the mailer failed, so reverting to 'mailer=\"none\"'\n")
            flush.console()
            none_method()
        }
    }
    invisible()
}

