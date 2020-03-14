#' Generate Sample Lists from Excel or CSV Files
#'
#' \code{sampler} generates Simple Random or Stratified samples
#'
#' @return Writes samples to an Excel workbook
#' @importFrom magrittr "%>%"
#' @importFrom tools file_ext
#' @importFrom purrr map2_dfr
#' @import openxlsx
#' @importFrom data.table fread
#' @import dplyr
#'
#' @export
#'
sampler <- function() {
  set.seed(12345)
  wb <- createWorkbook()

  dataName <- file.choose()

  if(file_ext(dataName)=='xlsx') {
    data <- read.xlsx(dataName)
  } else if(file_ext(dataName)=='csv') {
    data <- fread(dataName)
  } else paste("Not a valid data file")

  sampleSize <- whSample::ssize(nrow(data))

  sampleType <- utils::menu(c("Simple Random Sample",
                              "Stratified Random Sample",
                              "Tabbed Stratified Sample"),
                            graphics=T,
                            title="Sample Type")

  if(sampleType == 1L) {
    addWorksheet(wb, "Simple Random Sample")
    writeData(wb, "Simple Random Sample", data[sample(nrow(data)),])
    saveWorkbook(wb, file="Samples.xlsx", overwrite=T)

  } else {

    stratifyOn <- names(data)[utils::menu(names(data), graphics=T,
                                          title="Stratify on")]

    dataSamples <- data %>% group_by_at(stratifyOn) %>% count() %>%
      data.table::setDT() %>% mutate(prop = prop.table(n)) %>%
      mutate(numSamples = ceiling(ifelse(
        prop * sampleSize < 1, 1, prop * sampleSize
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
}
