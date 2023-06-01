library(readxl)
library(dplyr)
library(tidyverse)
library(writexl)
library(WriteXLS)
# CONTACT 




# Select month + target4call column : 128
# i = 12
# target4call <- read_excel("topserv_HO.xlsx",sheet=i)



setwd("/Users/parinya/Library/Mobile Documents/com~apple~CloudDocs/1.TBN/2.Call Center/Load TopservCrl/CRL Topserv/2023/เตรียม Data ลง R program")
# get file names
print(list.files())

## 1. load contact

contact_crl05 <- read_excel("ผลการติดต่อ052023.xlsx",sheet=1)
#contact_crl04 <- read_excel("ผลการติดต่อ042023.xlsx",sheet=1)
#contact_crl <- bind_rows(contact_crl05,contact_crl04)
contact_crl <- contact_crl05 %>% 
  select(12,calltype=18,25,30)
contact_crl <- contact_crl %>% 
  filter(calltype=='กิจกรรม UIO Utilization'|calltype=='โทรออก (เช็คระยะ) ตามแผนการโทร'|calltype==  'โทรเข้า (ด้านบริการ)'|calltype=='โทรออก (เช็คระยะ) นอกแผนการโทร')
contact_crl <- contact_crl %>% 
  select(vin=1,3,4)

contact_crl %>%
  summarise(n())

contact_crl <- contact_crl %>% 
  distinct(vin ,  .keep_all = TRUE)

contact_crl %>%
  summarise(n())
#--------------------------------------------------------



# Topsrc_crl 

## 2. load Topsrc_crl
##  crl <- read_excel("topserv_SO2022.xls",sheet=1)

UIO <- read_excel("2023.03 UIO ช้อมูล.xlsx",sheet=1)
names(UIO)
UIO <- UIO %>% 
  select(vin=2,5,6,9)

## 3. bind 12 month in Topserv_branch : Repleat SO,HO,RK,KS
SO <-list()
for ( i in 13:17) {
  SO[[i]] <- read_excel("topserv_SO.xlsx",sheet=i)
}


S1 <- SO[[13]]
S2 <- SO[[14]]
S3 <- SO[[15]]
S4 <- SO[[16]]
S5 <- SO[[17]]
SO <- bind_rows(S1,S2,S3,S4,S5)

#write.csv(SO2, "SO2.csv",row.names = FALSE)



HO <-list()
for ( i in 13:17) {
  HO[[i]] <- read_excel("topserv_HO.xlsx",sheet=i)
}


H1 <- HO[[13]]
H2 <- HO[[14]]
H3 <- HO[[15]]
H4 <- HO[[16]]
H5 <- HO[[17]]
HO <- bind_rows(H1,H2,H3,H4,H5)



RK <-list()
for ( i in 13:17)  {
  RK[[i]] <- read_excel("topserv_RK.xlsx",sheet=i)
}



R1 <- RK[[13]]
R2 <- RK[[14]]
R3 <- RK[[15]]
R4 <- RK[[16]]
R5 <- RK[[17]]
RK <- bind_rows(R1,R2,R3,R4,R5)


KS <-list()
for ( i in 13:17)  {
  KS[[i]] <- read_excel("topserv_KS.xlsx",sheet=i)
}


K1 <- KS[[13]]
K2 <- KS[[14]]
K3 <- KS[[15]]
K4 <- KS[[16]]
K5 <- KS[[17]]
KS <- bind_rows(K1,K2,K3,K4,K5)

## 3. bind 4 branch im topserv_crl

## (2) bind rows

topserv_crl <- bind_rows(HO,SO,RK,KS)

topserv_crl <- topserv_crl %>% 
  select(vin=7,11,13)





# 3. target : topserv 12 month Load data
#june2022 = sheet11


i = 12
target4call <- read_excel("topserv_SO.xlsx",sheet=i)

names(target4call)



topserv_contact <- topserv_crl %>%
  full_join(contact_crl,by = c("vin" = "vin"))  %>% 
  distinct(vin ,  .keep_all = TRUE)



target_call <- target4call %>%
  left_join(topserv_contact,by = c("หมายเลขตัวถัง"  = "vin"))


target_call <- target_call %>%
  left_join(UIO,by = c("หมายเลขตัวถัง"  = "vin"))

names(target_call)


#names(UIO_jo_topserv)

names(target_call)
#View(target_call)
target_call <- target_call %>%
  select(2,Owener_address=3,4,User_address=5,6,vin = 7,target_month=13,lastopenjob=21,lastpayin=22,lastContact=23,resultcontact=24,25:27) 

###################
## regular expression
library(stringr)



target_call$Owener_phone <- str_extract_all(target_call$Owener_address, "(?<=โทรศัพท์).+")
target_call$user_phone <- str_extract_all(target_call$User_address, "(?<=โทรศัพท์).+")

names(target_call)
#------------------------------
# target_call <- target_call %>%
#   select(1,3,5:13) 

target_call %>%
  summarise(n())


# set1 : all_active & Info_month ทุกเดือนที่แนะนำ & all age  
target_calldistinct <- target_call %>%
  filter(is.na(lastpayin))

## 4. Summarise การหาค่าสถิติ
target_call %>%
  summarise(n())

names(target_call)


# Remove duplicated rows based on vin

target_calldistinct <- target_calldistinct %>% distinct(vin,  .keep_all = TRUE)

target_calldistinct %>%
  summarise(n())

# ex2
# Remove duplicates based on Sepal.Width columns
# target_call1[!duplicated(target_call1$vin), ]

setwd("/Users/parinya/Library/Mobile Documents/com~apple~CloudDocs/1.TBN/2.Call Center/Load TopservCrl/CRL Topserv/2023/เตรียม Data ลง R program/data for call topserv")

#write_xlsx(target_call,"2.target_call_12month.xlsx")

#write_xlsx(target_calldistinct,"2.Last_service_12month_duplicatevin.xlsx")
write_xlsx(target_calldistinct,"2.Last_service_6month_duplicatevin.xlsx")




