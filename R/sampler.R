#' Generate Sample Lists from Excel or CSV Files
#'
#' \code{sampler} generates Simple Random or Stratified samples
#'
#' @return Writes samples to an Excel workbook and generates a report summary.
#' @name sampler
#' @importFrom magrittr "%>%"
#' @import tools
#' @importFrom purrr map2_dfr
#' @import openxlsx
#' @importFrom data.table fread setDF setDT
#' @import dplyr
#' @importFrom glue glue
#' @importFrom stats qnorm
#' @importFrom utils head tail
#' @param ci the required confidence level
#' @param me the margin of error
#' @param p the expected probability of occurrence
#' @param backups the number of available replacements
#' @param seed the random number seed
#' @export
#' @section Details:
#' \code{sampler} lets users select an Excel or CSV data file and the type of sample they prefer (Simple Random Sample, Stratified Random Sample, or Tabbed Stratified Sample with each stratum in a different Excel worksheet).
#' @examples
#' \dontrun{
#' sampler()
#' sampler(backups=5, p=0.6)
#' }

utils::globalVariables(c("prop", "."))

sampler <- function(ci=0.95, me=0.07, p=0.50, backups=5, seed=NULL) {

  # set up the Excel style
  hdrStyle <- createStyle(halign="center", valign="center",
                          borderColour="black", textDecoration="bold",
                          border="TopBottomLeftRight", wrapText=F,
                          borderStyle="thin", fgFill="#e7e6e6") # lt gray

  tableStyle <- createStyle(halign="center")

  pctStyle <- createStyle(halign="center", numFmt="0.0%")

  ifelse(!is.numeric(seed), rns <- as.integer(Sys.time()), rns <- seed)
  set.seed(rns)

  # choose the source file
  dataName <- file.choose()

  # save the path to it so we can write to the same place
  wb.path <- dirname(dataName)

  wb <- createWorkbook(dataName)

  if(file_ext(dataName)=='xlsx') {
    sheetNames <- (getSheetNames(dataName))
    tabMenu <- utils::menu(sheetNames, graphics=T,
                            title="Use sheet")
    SrcTab <- sheetNames[tabMenu]
    data <- read.xlsx(dataName, sheet=SrcTab)
  } else if(file_ext(dataName)=='csv') {
    data <- fread(dataName)
  } else paste("Not a valid data file")

  N <- nrow(data)

  sampleSize <- whSample::ssize(N, ci, me, p)

  backupMenu <- utils::menu(c("0", "5", "10"),
                         graphics=T, title="Number of backups")
  backups <- ifelse(backupMenu==0, backups,
                    switch(backupMenu, 0, 5, 10))

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
  writeDataTable(new.wb, "Original", data, tableName="Data", withFilter=F)
  setColWidths(new.wb, "Original", cols=1:ncol(data), widths="auto")
  addStyle(new.wb, "Original", tableStyle,
           rows=1, cols=1:length(data))

  if(sampleType == 1L) {
    numStrata <- 1
    numSamples <- sampleSize+backups
    addWorksheet(new.wb, "Simple Random Sample")

    allSamples <- sample_n(data, numSamples)
    primarySamples <- head(allSamples, sampleSize)
    backupSamples <- tail(allSamples, backups)

    writeDataTable(new.wb, "Simple Random Sample", primarySamples,
              tableName="primarySRS", withFilter=F)
    addStyle(new.wb, "Simple Random Sample", tableStyle,
             rows=1, cols=1:length(primarySamples))

    mergeCells(new.wb, "Simple Random Sample", cols=1:length(primarySamples),
               rows=nrow(primarySamples)+3)
    writeData(new.wb, "Simple Random Sample", "Backup Samples",
              startRow=nrow(primarySamples)+3)
    addStyle(new.wb,"Simple Random Sample", hdrStyle,
             rows=nrow(primarySamples)+3,
             cols=1:length(primarySamples))

    writeDataTable(new.wb, "Simple Random Sample", backupSamples,
                   startRow=nrow(primarySamples)+4)

    setColWidths(new.wb, "Simple Random Sample", cols=1:ncol(data),
                 widths="auto")
  } else {

    stratifyOn <- names(data)[utils::menu(names(data), graphics=T,
                                          title="Stratify on")]

    dataSamples <- data %>% group_by_at(stratifyOn) %>% count() %>%
      data.table::setDT() %>% mutate(prop = prop.table(n)) %>%
      mutate(numSamples = ceiling(ifelse(
        backups +
        prop * sampleSize < 1, 1, prop * sampleSize
      ))) %>%
      mutate(numBackups=ifelse(
        numSamples+backups > n, (n-numSamples), backups))

    numStrata <- nrow(dataSamples)

    setDF(data)
    primarySamples <- split(data, data[stratifyOn]) %>%
      map2_dfr(., dataSamples$numSamples,
               ~head(.x, .y))

    backupSamples <- split(data, data[stratifyOn]) %>%
      map2_dfr(., dataSamples$numBackups,
               ~tail(.x, .y))

    if(sampleType == 2L){

      addWorksheet(new.wb,"Stratified Random Sample")
      writeDataTable(new.wb,"Stratified Random Sample", primarySamples,
                     tableName="Stratified", withFilter=F)
      addStyle(new.wb, "Stratified Random Sample", tableStyle,
               rows=1, cols=1:length(primarySamples))

      mergeCells(new.wb, "Stratified Random Sample",
                 cols=1:length(primarySamples),
                 rows=nrow(primarySamples)+3)
      writeData(new.wb, "Stratified Random Sample", "Backup Samples",
                startRow=nrow(primarySamples)+3)
      addStyle(new.wb,"Stratified Random Sample", hdrStyle,
               rows=nrow(primarySamples)+3,
               cols=1:length(primarySamples))

      writeDataTable(new.wb, "Stratified Random Sample", backupSamples,
                     startRow=nrow(primarySamples)+4,
                     tableName="Stratified_Backups", withFilter=F)
      addStyle(new.wb, "Stratified Random Sample", tableStyle,
               rows=nrow(primarySamples)+4, cols=1:length(primarySamples))

      setColWidths(new.wb, "Stratified Random Sample", cols=1:ncol(data),
                   widths="auto")

    } else if(sampleType == 3L){
      primaryTabs <- split(primarySamples, primarySamples[stratifyOn])
      backupTabs <- split(backupSamples, backupSamples[stratifyOn])

      invisible(Map(function(primary, backup, name) {
        addWorksheet(new.wb, name)
        writeDataTable(new.wb, name, primary, withFilter=F)
        addStyle(new.wb, name, tableStyle, rows=1, cols=1:length(primary))

        mergeCells(new.wb, name, cols=1:length(primary),
                   rows=nrow(primary)+3)
        writeData(new.wb, name, "Backup Samples",
                  startRow=nrow(primary)+3)
        addStyle(new.wb,name, hdrStyle,
                 rows=nrow(primary)+3,
                 cols=1:length(primary))

        writeDataTable(new.wb, name, backup,
                       startRow=nrow(primary)+4, withFilter=F)
        addStyle(new.wb, name, tableStyle,
                 rows=nrow(primary)+4, cols=1:length(primary))

        setColWidths(new.wb, name, cols=1:ncol(primary),
                     widths="auto")
      }, primaryTabs, backupTabs, names(primaryTabs)))

    }
  }
    addWorksheet(new.wb,"Report")
    writeDataTable(new.wb,"Report", withFilter=F, x=
                data.frame("Variable"=c("Source","Source Size","Sample Type",
                                        "Sample Size",
                                        "Strata","Backups per Stratum",
                                        "Random Number Seed", "Created"),
                           "Value"=c(dataName, N, sampleTypeName, sampleSize,
                                     numStrata, backups, rns,
                                     as.character(
                                       Sys.time()))))
    # addStyle(new.wb, "Report", hdrStyle, rows=1, cols=1:2)
    setColWidths(new.wb, "Report", cols=1:2, widths="auto")

    if(numStrata!=1L) {
      writeDataTable(new.wb,"Report", startRow=1, startCol=4,
                withFilter=F, x=data.frame(dataSamples))

      # addStyle(new.wb, "Report", hdrStyle, rows=1, cols=4:8)
      addStyle(new.wb, "Report", pctStyle, rows=1:nrow(dataSamples)+1, cols=6,
               stack=T)
      setColWidths(new.wb, "Report", cols=4:(ncol(dataSamples)+4),
                   widths="auto")
    }

    saveWorkbook(new.wb,new.wb.name,overwrite=T)

}
