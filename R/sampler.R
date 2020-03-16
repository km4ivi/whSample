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
#' @section Details:
#' \code{sampler} lets users select an Excel or CSV data file and the type of sample they prefer (Simple Random Sample, Stratified Random Sample, or Tabbed Stratified Sample with each stratum in a different Excel worksheet).
#' @examples
#' sampler()
#' sampler(backups=0, p=0.6)
#' @export
#'
sampler <- function(ci=0.95, me=0.07, p=0.50, backups=5, seed=NULL) {

  ifelse(!is.numeric(seed), rns <- as.integer(Sys.time()), rns <- seed)
  set.seed(rns)

  wb <- createWorkbook()

  dataName <- file.choose()

  if(file_ext(dataName)=='xlsx') {
    data <- read.xlsx(dataName)
  } else if(file_ext(dataName)=='csv') {
    data <- fread(dataName)
  } else paste("Not a valid data file")

  N <- nrow(data)

 # sampleSize <- whSample::ssize(nrow(data))
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

  if(sampleType == 1L) {
    numSamples <- sampleSize+backups
    addWorksheet(wb, "Simple Random Sample")
    writeData(wb, "Simple Random Sample", data[sample(numSamples),])
#    writeData(wb, "Simple Random Sample", data[sample(nrow(data)),])
    saveWorkbook(wb, file="Samples.xlsx", overwrite=T)

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
      writeData(wb,"Stratified Random Sample", data2)

    } else if(sampleType == 3L){
      dataTabs <- split(data2, data2[stratifyOn])

      Map(function(data, name){
        addWorksheet(wb, name)
        writeData(wb, name, data)
      }, dataTabs, names(dataTabs))

    }
    saveWorkbook(wb, file = "Samples.xlsx", overwrite=T)
  }
  report <- t(data.frame("Source"=dataName,"Source Size"=N,
                         "Sample Type"=sampleTypeName,
                         "Sample Size"=sampleSize,'Backups'=backups,
                         "Random Number Seed"=rns,
                         "Created"=file.info("samples.xlsx")$ctime))

  write.table(report,"Sampling Report.csv",col.names=F,sep=",")
}
