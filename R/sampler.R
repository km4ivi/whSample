#' Generate Sample Lists from Excel or CSV Files
#'
#' \code{sampler} generates Simple Random or Stratified samples
#'
#'@return Writes samples to an Excel workbook
#'@export

sampler <- function()

dataName <- file.choose()

if (file_ext(dataName)=='xlsx') {
  data <- read.xlsx(dataName)
  } else if (file_ext(dataName) == 'csv') {
    data <- fread(dataName)
  } else print("Not a valid data file")

sample.size <- ssize(nrow(data))

set.seed(12345)       # Allows for replication; otherwise not necessary

print(sample.size)

# wb <- createWorkbook()
#
# addWorksheet(wb,"Simple Random Sample")
# writeData(wb,"Simple Random Sample", data[sample(nrow(data)),])
#
# saveWorkbook(wb, file = "Samples.xlsx", overwrite = TRUE)
#
# data.samples <-
#   data %>% count(Function) %>%
#   mutate(prop=prop.table(n)) %>%
#   mutate(numSamples=ceiling(
#     ifelse(prop*sample.size<1, 1, prop*sample.size)))
#
# data.list <- split(data,data$Function)
#
# data3 <- map2_dfr(data.list, data.samples$numSamples,
#                              ~sample_n(.x, .y))
# data3 %>% count(Function)
#
# addWorksheet(wb,"Random by Function")
# writeData(wb,"Random by Function", data3)
# data.split <- split(data3,data3$Function)
#
# Map(function(data, name){
#   addWorksheet(wb, name)
#   writeData(wb, name, data)
# }, data.split, names(data.split))
#
# saveWorkbook(wb, file = "Samples.xlsx", overwrite = TRUE)
