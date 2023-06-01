library(readxl)
library(dplyr)
library(tidyverse)
library(writexl)
library(WriteXLS)
library(openxlsx)
#install.packages("xlsx")                          # Install xlsx package in R


setwd("/Users/parinya/Library/Mobile Documents/com~apple~CloudDocs/1.TBN/2.Call Center/Load TopservCrl/CRL Topserv/2023/เตรียม Data ลง R program")


## Select 16 sheet crl_topserv data
topHO <-list()
for ( i in 1:17) {
  topHO[[i]] <- read_excel("topserv_HO.xlsx",sheet=i)
}
topSO <-list()
for ( i in 1:17) {
  topSO[[i]] <- read_excel("topserv_SO.xlsx",sheet=i)
}
topRK <-list()
for ( i in 1:17) {
  topRK[[i]] <- read_excel("topserv_RK.xlsx",sheet=i)
}
topKS <-list()
for ( i in 1:17) {
  topKS[[i]] <- read_excel("topserv_KS.xlsx",sheet=i)
}
write.xlsx(topHO,file="topserv_HO.xlsx")
write.xlsx(topSO,file="topserv_SO.xlsx")
write.xlsx(topRK,file="topserv_RK.xlsx")
write.xlsx(topKS,file="topserv_KS.xlsx")




library(openxlsx)

### New month #####
Newsheet = "Sheet18(june)"

HO <- loadWorkbook("topserv_HO.xlsx")
addWorksheet(HO,Newsheet)

HO2 <- read_excel("HO.xlsx")
writeData(HO,Newsheet,HO2)
saveWorkbook(HO,"topserv_HO.xlsx",overwrite = TRUE)


SO <- loadWorkbook("topserv_SO.xlsx")
addWorksheet(SO,Newsheet)
SO2 <- read_excel("SO.xlsx")
writeData(SO,Newsheet,SO2)
saveWorkbook(SO,"topserv_SO.xlsx",overwrite = TRUE)

RK <- loadWorkbook("topserv_RK.xlsx")
addWorksheet(RK,Newsheet)
RK2 <- read_excel("RK.xlsx")
writeData(RK,Newsheet,RK2)
saveWorkbook(RK,"topserv_RK.xlsx",overwrite = TRUE)

KS <- loadWorkbook("topserv_KS.xlsx")
addWorksheet(KS,Newsheet)
KS2 <- read_excel("KS.xlsx")
writeData(KS,Newsheet,KS2)
saveWorkbook(KS,"topserv_KS.xlsx",overwrite = TRUE)

