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









# Write xlsx with multiple sheets
# HO2 <- read_excel("HO.xlsx")
# 
# #####----------------- วิธีที่ 2 ------------------------
# HO <-list()
# for ( i in 13:16) {
#   HO[[i]] <- read_excel("topserv_HO.xlsx",sheet=i)
# }
# 
# H1 <- HO[[13]]
# H2 <- HO[[14]]
# H3 <- HO[[15]]
# H4 <- HO[[16]]
# H5 <- updateTopserv
# HO <- bind_rows(H1,H2,H3,H4,H5)
# 
# 
# 
# 
# 
# updateTopserv <- read_excel("SO.xlsx")
# SO <-list()
# for ( i in 13:17) {
#   SO[[i]] <- read_excel("topserv_SO.xlsx",sheet=i)
# }
# 
# 
# S1 <- SO[[13]]
# S2 <- SO[[14]]
# S3 <- SO[[15]]
# S4 <- SO[[16]]
# S5 <- updateTopserv
# SO <- bind_rows(S1,S2,S3,S4,S5)
# 
# #write.csv(SO2, "SO2.csv",row.names = FALSE)
# 
# 
# updateTopserv <- read_excel("RK.xlsx")
# RK <-list()
# for ( i in 13:17)  {
#   RK[[i]] <- read_excel("topserv_RK.xlsx",sheet=i)
# }
# 
# 
# 
# R1 <- RK[[13]]
# R2 <- RK[[14]]
# R3 <- RK[[15]]
# R4 <- RK[[16]]
# R5 <- updateTopserv
# RK <- bind_rows(R1,R2,R3,R4,R5)
# 
# updateTopserv <- read_excel("KS.xlsx")
# KS <-list()
# for ( i in 13:17)  {
#   KS[[i]] <- read_excel("topserv_KS.xlsx",sheet=i)
# }
# 
# 
# K1 <- KS[[13]]
# K2 <- KS[[14]]
# K3 <- KS[[15]]
# K4 <- KS[[16]]
# K5 <- updateTopserv
# KS <- bind_rows(K1,K2,K3,K4,K5)
# 
# 
# 
# #-------------------------------------------------
# 
# ## update crl_topserv data
# updateTopserv <- read_excel("HO.xlsx")
# 
# HO <-list()
# for ( i in 13:16) {
#   HO[[i]] <- read_excel("topserv_HO.xlsx",sheet=i)
# }
# 
# H1 <- HO[[13]]
# H2 <- HO[[14]]
# H3 <- HO[[15]]
# H4 <- HO[[16]]
# H5 <- updateTopserv
# HO <- bind_rows(H1,H2,H3,H4,H5)
# 
# updateTopserv <- read_excel("SO.xlsx")
# SO <-list()
# for ( i in 13:17) {
#   SO[[i]] <- read_excel("topserv_SO.xlsx",sheet=i)
# }
# 
# 
# S1 <- SO[[13]]
# S2 <- SO[[14]]
# S3 <- SO[[15]]
# S4 <- SO[[16]]
# S5 <- updateTopserv
# SO <- bind_rows(S1,S2,S3,S4,S5)
# 
# #write.csv(SO2, "SO2.csv",row.names = FALSE)
# 
# 
# updateTopserv <- read_excel("RK.xlsx")
# RK <-list()
# for ( i in 13:17)  {
#   RK[[i]] <- read_excel("topserv_RK.xlsx",sheet=i)
# }
# 
# 
# 
# R1 <- RK[[13]]
# R2 <- RK[[14]]
# R3 <- RK[[15]]
# R4 <- RK[[16]]
# R5 <- updateTopserv
# RK <- bind_rows(R1,R2,R3,R4,R5)
# 
# updateTopserv <- read_excel("KS.xlsx")
# KS <-list()
# for ( i in 13:17)  {
#   KS[[i]] <- read_excel("topserv_KS.xlsx",sheet=i)
# }
# 
# 
# K1 <- KS[[13]]
# K2 <- KS[[14]]
# K3 <- KS[[15]]
# K4 <- KS[[16]]
# K5 <- updateTopserv
# KS <- bind_rows(K1,K2,K3,K4,K5)
