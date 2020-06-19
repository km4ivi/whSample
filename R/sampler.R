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
#' @section Details:
#' \code{sampler} lets users select an Excel or CSV data file and the type of sample they prefer (Simple Random Sample, Stratified Random Sample, or Tabbed Stratified Sample with each stratum in a different Excel worksheet).
#' @examples
#' sampler()
#' sampler(backups=0, p=0.6)
#' @export
#'

sampler <- function(ci=0.95, me=0.07, p=0.50, backups=0, seed=NULL) {

  # set up the Excel style
  headerStyle <- createStyle(halign="center", valign="center",
                             borderColour="black", textDecoration="bold",
                             border="TopBottomLeftRight", wrapText=F,
                             borderStyle="thin", fgFill="#e7e6e6") # lt gray

  ifelse(!is.numeric(seed), rns <- as.integer(Sys.time()), rns <- seed)
  set.seed(rns)

  # wb <- createWorkbook()
  # dataName <- file.choose()

  # choose the source file
  dataName <- file.choose()

  # save the path to it so we can write to the same place
  wb.path <- dirname(dataName)

  wb <- loadWorkbook(dataName)

  if(file_ext(dataName)=='xlsx') {
    data <- read.xlsx(wb)
  } else if(file_ext(dataName)=='csv') {
    data <- fread(wb)
  } else paste("Not a valid data file")

  N <- nrow(data)

  sampleSize <- whSample::ssize(N, ci, me, p)

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
  writeData(new.wb, "Original", data)
  setColWidths(new.wb, "Original", cols=1:ncol(data), widths="auto")
  addStyle(new.wb, "Original", headerStyle, rows=1, cols=1:ncol(data))

  if(sampleType == 1L) {
    numSamples <- sampleSize+backups
    addWorksheet(new.wb, "Simple Random Sample")
    writeData(new.wb, "Simple Random Sample", data[sample(numSamples),])
    setColWidths(new.wb, "Simple Random Sample", cols=1:ncol(data),
                 widths="auto")
    addStyle(new.wb, "Simple Random Sample", headerStyle, rows=1,
             cols=1:ncol(data))
    saveWorkbook(new.wb, new.wb.name, overwrite=T)

  } else {

    stratifyOn <- names(data)[utils::menu(names(data), graphics=T,
                                          title="Stratify on")]

    dataSamples <- data %>% group_by_at(stratifyOn) %>% count() %>%
      data.table::setDT() %>% mutate(prop = prop.table(n)) %>%
      mutate(numSamples = ceiling(ifelse(
        backups +
        prop * sampleSize < 1, 1, backups + prop * sampleSize
      )))

    data.table::setDF(data)
    dataList <- split(data,data[stratifyOn])

    data2 <- map2_dfr(dataList, dataSamples$numSamples,
                      ~sample_n(.x, .y))

    if(sampleType == 2L){
      addWorksheet(wb,"Stratified Random Sample")
      setColWidths(new.wb, "Stratified Random Sample", cols=1:ncol(data),
                   widths="auto")
      addStyle(new.wb, "Stratified Random Sample", headerStyle, rows=1,
               cols=1:ncol(data))
      writeData(wb,"Stratified Random Sample", data2)

    } else if(sampleType == 3L){
      dataTabs <- split(data2, data2[stratifyOn])

      Map(function(data, name){
        addWorksheet(wb, name)
        writeData(wb, name, data)
        setColWidths(new.wb, "Tabbed Random Sample", cols=1:ncol(data),
                     widths="auto")
        addStyle(new.wb, "Tabbed Random Sample", headerStyle, rows=1,
                 cols=1:ncol(data))
      }, dataTabs, names(dataTabs))

    }
    saveWorkbook(wb, new.wb.name, overwrite=T)
  }
  report <- t(data.frame("Source"=dataName,"Source Size"=N,
                         "Sample Type"=sampleTypeName,
                         "Sample Size"=sampleSize,'Backups'=backups,
                         "Random Number Seed"=rns,
                         "Created"=file.info("samples.xlsx")$ctime))

  write.table(report,"Sampling Report.csv",col.names=F,sep=",")
}
