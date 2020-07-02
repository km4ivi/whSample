#' Generate Sample Lists from Excel or CSV Files
#'
#' \code{sampler} generates Simple Random or Stratified samples
#'
#' @return Writes samples to an Excel workbook and generates a report summary.
#' @importFrom magrittr "%>%"
#' @importFrom tools file_ext
#' @importFrom purrr map2_dfr
#' @import openxlsx
#' @importFrom data.table fread
#' @import dplyr
#' @import glue
#' @importFrom stats qnorm
#' @importFrom utils install.packages, installed.packages
#' @section Details:
#' \code{sampler} lets users select an Excel or CSV data file and the type of sample they prefer (Simple Random Sample, Stratified Random Sample, or Tabbed Stratified Sample with each stratum in a different Excel worksheet).
#' @examples
#' sampler()
#' sampler(backups=0, p=0.6)
#' @export
#'

sampler <- function(ci=0.95, me=0.07, p=0.50, backups=0, seed=NULL) {

  # install necessary packages
  is_installed <- function(mypkg) is.element(mypkg, installed.packages()[,1])
  whInstall <- function(pkgNames){
    for(pkg in pkgNames){
      if(!is_installed(pkg)){
        install.packages(pkg, repos="http://lib.stat.cmu.edu/R/CRAN")
      }
      library(pkg, character.only=T, quietly=T, verbose=F)
    }
  }
  whInstall(c("magrittr","tools","purrr","openxlsx","data.table","dplyr","glue"))

  # set up the Excel style
  hdrStyle <- createStyle(halign="center", valign="center",
                          borderColour="black", textDecoration="bold",
                          border="TopBottomLeftRight", wrapText=F,
                          borderStyle="thin", fgFill="#e7e6e6") # lt gray

  pctStyle <- createStyle(halign="center", numFmt="0.0%")

  ifelse(!is.numeric(seed), rns <- as.integer(Sys.time()), rns <- seed)
  set.seed(rns)

  # choose the source file
  dataName <- file.choose()

  # save the path to it so we can write to the same place
  wb.path <- dirname(dataName)

  wb <- createWorkbook(dataName)

  if(file_ext(dataName)=='xlsx') {
    data <- read.xlsx(dataName)
  } else if(file_ext(dataName)=='csv') {
    data <- fread(dataName)
  } else paste("Not a valid data file")

  N <- nrow(data)

  sampleSize <- whSample::ssize(N, ci, me, p)

  numBackups <- utils::menu(c(0, 5, 10,"Custom"),
                         graphics=T, title="Number of backups")
  backups <- switch(numBackups, 0, 5, 10, backups)

  sampleType <- utils::menu(c("Simple Random Sample",
                              "Stratified Random Sample",
                              "Tabbed Stratified Sample"),
                            graphics=T,
                            title="Sample Type")

  sampleTypeName <- switch(sampleType,
                           "Simple Random Sample",
                           "Stratified Random Sample",
                           "Tabbed Stratified Sample")

  # create a new output workbook
  new.wb <- createWorkbook()
  new.wb.name <- glue('{wb.path}/{file_path_sans_ext(dataName) %>%
                      basename()} Sample.xlsx')

  # include the original worksheet for reference
  addWorksheet(new.wb, "Original")
  writeDataTable(new.wb, "Original", data, tableName="Data")
  setColWidths(new.wb, "Original", cols=1:ncol(data), widths="auto")

  if(sampleType == 1L) {
    numStrata <- 1
    numSamples <- sampleSize+(numStrata*backups)
    addWorksheet(new.wb, "Simple Random Sample")
    writeDataTable(new.wb, "Simple Random Sample", data[sample(numSamples),],
              tableName="SRS")
    setColWidths(new.wb, "Simple Random Sample", cols=1:ncol(data),
                 widths="auto")
    # saveWorkbook(new.wb, new.wb.name, overwrite=T)

  } else {

    stratifyOn <- names(data)[utils::menu(names(data), graphics=T,
                                          title="Stratify on")]

    dataSamples <- data %>% group_by_at(stratifyOn) %>% count() %>%
      data.table::setDT() %>% mutate(prop = prop.table(n)) %>%
      mutate(numSamples = ceiling(ifelse(
        backups +
        prop * sampleSize < 1, 1, backups + prop * sampleSize
      )))

    numStrata <- nrow(dataSamples)
    numSamples <- sampleSize+(numStrata*backups)

    data.table::setDF(data)
    dataList <- split(data,data[stratifyOn])

    data2 <- map2_dfr(dataList, dataSamples$numSamples,
                      ~sample_n(.x, .y))

    if(sampleType == 2L){

      addWorksheet(new.wb,"Stratified Random Sample")
      writeDataTable(new.wb,"Stratified Random Sample", data2,
                     tableName="Stratified")
      setColWidths(new.wb, "Stratified Random Sample", cols=1:ncol(data),
                   widths="auto")
      # addStyle(new.wb, "Stratified Random Sample", hdrStyle, rows=1,
      #          cols=1:ncol(data))


    } else if(sampleType == 3L){
      dataTabs <- split(data2, data2[stratifyOn])

      Map(function(data, name){
        addWorksheet(new.wb, name)
        writeDataTable(new.wb, name, data)
        setColWidths(new.wb, name, cols=1:ncol(data),
                     widths="auto")
        # addStyle(new.wb, name, hdrStyle, rows=1,
        #          cols=1:ncol(data))
      }, dataTabs, names(dataTabs))

    }
  }
    addWorksheet(new.wb,"Report")
    writeData(new.wb,"Report", withFilter=F, borders="all", x=
                data.frame("Variable"=c("Source","Source Size","Sample Type",
                                        "Sample Size",
                                        "Strata","Backups per Stratum",
                                        "Random Number Seed", "Created"),
                           "Value"=c(dataName, N, sampleTypeName, sampleSize,
                                     numStrata, backups, rns,
                                     as.character(
                                       Sys.time()))))
    addStyle(new.wb, "Report", hdrStyle, rows=1, cols=1:2)
    setColWidths(new.wb, "Report", cols=1:2, widths="auto")

    writeData(new.wb,"Report", startRow=1, startCol=4, borders="all",
              withFilter=F, x=format.data.frame(dataSamples,digits=2))

    addStyle(new.wb, "Report", hdrStyle, rows=1, cols=4:7)
    addStyle(new.wb, "Report", pctStyle, rows=2:nrow(dataSamples)+1, cols=6,
             stack=T)
    setColWidths(new.wb, "Report", cols=4:(ncol(dataSamples)+4),
                 widths="auto")

    saveWorkbook(new.wb,new.wb.name,overwrite=T)

}
