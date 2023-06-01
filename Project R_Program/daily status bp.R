library(readxl)
library(dplyr)
library(tidyverse)
# import writexl library
library(writexl)
library(WriteXLS)
library(ggplot2)
library(lubridate)

setwd("/Users/parinya/Library/Mobile Documents/com~apple~CloudDocs/1.TBN/BP/1.BP Daily/042023/BP 23042023")

# setwd(รายงานสถานะแต่ละวัน bp)

# CONTACT 

## 1. load contact
TBN_HO <- read_excel("รายงานสถานะรถ บางนา-TBN.xlsx",sheet=2,skip = 1)
TBN_SO <- read_excel("1.รายงานสถานะรถ ลาดกระบัง-TBN.xlsx",sheet="สถานะรถปี 2566 ",skip=2)
TBN_RK <- read_excel("รายงานสถานะรถ รามคำแหง - TBN.xlsx",sheet="2566",skip = 1)
TBN_KS <- read_excel("รายงานสถานะรถ TBNกาญจนา.xlsx",sheet=1,skip = 1,na ="-")

str(TBN_HO )

# rename column + replace "_" == ""  + num : Rows by column value
TBN_HO <- TBN_HO %>%
  select(1,contact_date=2,3,plateId=4,bpe_date=5,ins_date=6,wait_ins=7,approve_date=8,wait_approve=9,customer_date=10,wait_customer=11,TBN_app=12,tbn_finish = 13 ,overdue=14,15:16,laber=17,part=18,19:23)
# 

str(TBN_HO)


# change type chr to numeric
class(TBN_HO$part)
TBN_HO$part

TBN_HO$part <- as.character(sub("_", "",TBN_HO$part)) %>%
  as.numeric(TBN_HO$part)

class(TBN_HO$part )
#-------------------------------
str(TBN_SO )
# +++ วันที่ประกันอนุมัติ : chr ++++


TBN_SO <- TBN_SO %>%
  select(1,contact_date=2,3,plateId=4,bpe_date=5,ins_date=6,wait_ins=7,approve_date=8,wait_approve=9,customer_date=10,wait_customer=11,TBN_app=12,tbn_finish = 13 ,overdue=14,15:17,laber=18,part=19,20:24) 

names(TBN_SO)


### change type chr to as.Date
class(TBN_SO$approve_date)



# 
# TBN_SO$approve_date <- as.character(sub("ติดเรื่องกรมธรรม์", "0",TBN_SO$approve_date)) %>%
#   as.Date(TBN_SO$approve_date)
# TBN_SO$approve_date <- as.Date(TBN_SO$approve_date)
TBN_SO$approve_date
#-------------------------------
TBN_SO$approve_date


str(TBN_RK )
# ค่าอะไหล่      : chr


TBN_RK <- TBN_RK %>%
  select(1,contact_date=2,3,plateId=4,bpe_date=5,ins_date=6,wait_ins=7,approve_date=8,wait_approve=9,customer_date=10,wait_customer=11,TBN_app=12,tbn_finish = 13 ,overdue=14,15:16,laber=17,part=18,19:23)

TBN_RK$part
names(TBN_RK)

# change type chr to numeric
TBN_RK$part <- as.numeric(TBN_RK$part)

class(TBN_RK$part)
TBN_RK$part
#-------------------------------


TBN_KS <- TBN_KS %>%
  select(1,contact_date=2,3,plateId=4,bpe_date=5,ins_date=6,wait_ins=7,approve_date=8,wait_approve=9,customer_date=10,wait_customer=11,TBN_app=12,tbn_finish = 13 ,overdue=14,15:16,laber=17,part=18,19:20,30)

str(TBN_KS)
# wait_ins ,wait_approve,wait_customer, overdue  : chr

TBN_KS$wait_ins <- as.numeric(TBN_KS$wait_ins)
TBN_KS$wait_approve <- as.numeric(TBN_KS$wait_approve)
TBN_KS$wait_customer <- as.numeric(TBN_KS$wait_customer)
TBN_KS$overdue <- as.numeric(TBN_KS$overdue)

class(TBN_KS$overdue)
names(TBN_KS)

TBN_KS$overdue

## check type
str(TBN_HO)
str(TBN_SO)
str(TBN_RK)
str(TBN_KS)

# ----------------------

## Load sub 
# delete head ==>skip = 1

sub_HO <- read_excel("รายงานสถานะรถ บางนา-SUB.xlsx",sheet="บางนา (Oct-now)",skip = 1)
sub_SO <- read_excel("รายงานสถานะรถ ลาดกระบัง-SUB.xlsx",sheet="ลาดกระบัง (Oct-now)",skip = 1)
sub_RK <- read_excel("รายงานสถานะรถ รามคำแหง-SUB.xlsx",sheet="รามคำแหง (Jan-now)",skip = 1)
sub_KS <- read_excel("รายงานสถานะรถ กาญจนาภิเษก - SUB.xlsx",sheet=1,skip = 1)



# View(sub_SO)
# View(sub_RK)
# View(sub_HO)


## (2) join table


HO <- TBN_HO %>%
  left_join(sub_HO,by = c("plateId" = "ทะเบียน"))

SO <- TBN_SO %>%
  left_join(sub_SO,by = c("plateId" = "ทะเบียน"))

RK <- TBN_RK %>%
  left_join(sub_RK,by = c("plateId" = "ทะเบียน"))

KS <- TBN_KS %>%
  left_join(sub_KS,by = c("plateId" = "ทะเบียน"))


#------------------- Step 2 ------------------------------


## 4. Summarise การหาค่าสถิติ
#View(SO)
names(HO)

str(SO)

HO <- HO %>%
  select(plateId=4,2,5,6,7,8,9,10,11,12,13,14,17,18,20,remark=23,service_date=25,target_date=26,sub_finish = 27,overdue_sub=28,comment = 49) %>%
  filter(!is.na(HO[[4]])) 






#------------
names(SO)

SO <- SO %>%
  select(4,2,5,6,7,8,9,10,11,12,13,14,18,19,21,remark=24,service_date=26,target_date=27,sub_finish = 28,overdue_sub=29,comment = 50) %>%
  filter(!is.na(SO[[4]]))


#------------
names(RK)
RK <- RK %>%
  select(4,2,5:13,14,17,18,20,remark=23,service_date=25,target_date=26,sub_finish = 27,overdue_sub=28,comment = 49) %>%
  filter(!is.na(RK[[4]])) 

#------------
names(KS)
KS <- KS %>%
  select(4,2,5:14,17,18,20,remark=21,service_date=23,target_date=24,sub_finish = 25,overdue_sub=26,comment = 48) %>%
  filter(!is.na(KS[[4]])) 


#--------- add column ----------------

HO$contact_month <- as.numeric(format(HO$contact_date,'%m'))
HO$contact_year <- as.numeric(format(HO$contact_date,'%Y'))
SO$contact_month <- as.numeric(format(SO$contact_date,'%m'))
SO$contact_year <- as.numeric(format(SO$contact_date,'%Y'))
RK$contact_month <- as.numeric(format(RK$contact_date,'%m'))
RK$contact_year <- as.numeric(format(RK$contact_date,'%Y'))
KS$contact_month <- as.numeric(format(KS$contact_date,'%m'))
KS$contact_year <- as.numeric(format(KS$contact_date,'%Y'))


str(KS)


HO$ins_month <- as.numeric(format(HO$ins_date,'%m'))
HO$ins_year <- as.numeric(format(HO$ins_date,'%Y'))
SO$ins_month <- as.numeric(format(SO$ins_date,'%m'))
SO$ins_year <- as.numeric(format(SO$ins_date,'%Y'))
RK$ins_month <- as.numeric(format(RK$ins_date,'%m'))
RK$ins_year <- as.numeric(format(RK$ins_date,'%Y'))
KS$ins_month <- as.numeric(format(KS$ins_date,'%m'))
KS$ins_year <- as.numeric(format(KS$ins_date,'%Y'))


#format(dmy(SO$approve_date ), "%m-%d-%Y")

HO$approve_month <- as.numeric(format(HO$approve_date,'%m'))
HO$approve_year <- as.numeric(format(HO$approve_date,'%Y'))
SO$approve_month <- as.numeric(format(SO$approve_date,'%m'))
SO$approve_year <- as.numeric(format(SO$approve_date,'%Y'))
RK$approve_month <- as.numeric(format(RK$approve_date,'%m'))
RK$approve_year <- as.numeric(format(RK$approve_date,'%Y'))
KS$approve_month <- as.numeric(format(KS$approve_date,'%m'))
KS$approve_year <- as.numeric(format(KS$approve_date,'%Y'))


HO$customer_month <- as.numeric(format(HO$customer_date,'%m'))
HO$customer_year <- as.numeric(format(HO$customer_date,'%Y'))
SO$customer_month <- as.numeric(format(SO$customer_date,'%m'))
SO$customer_year <- as.numeric(format(SO$customer_date,'%Y'))
RK$customer_month <- as.numeric(format(RK$customer_date,'%m'))
RK$customer_year <- as.numeric(format(RK$customer_date,'%Y'))
KS$customer_month <- as.numeric(format(KS$customer_date,'%m'))
KS$customer_year <- as.numeric(format(KS$customer_date,'%Y'))


#HO$customer_day <- as.numeric(format(HO$customer_date,'%d'))
#SO$customer_day <- as.numeric(format(SO$customer_date,'%d'))
#RK$customer_day <- as.numeric(format(RK$customer_date,'%d'))
#KS$customer_day <- as.numeric(format(KS$customer_date,'%d'))


# change Date format in multi column

HO$contact_date <- ymd(HO$contact_date)
HO$bpe_date <-ymd(HO$bpe_date)
HO$ins_date <- ymd(HO$ins_date)
HO$approve_date <-ymd(HO$approve_date)
HO$customer_date <-ymd(HO$customer_date)
HO$TBN_app <- ymd(HO$TBN_app)
HO$tbn_finish <- ymd(HO$tbn_finish)
HO$service_date <- ymd(HO$service_date)
HO$target_date <- ymd(HO$target_date)
HO$sub_finish <- ymd(HO$sub_finish)


SO$contact_date <- ymd(SO$contact_date)
SO$bpe_date <-ymd(SO$bpe_date)
SO$ins_date <- ymd(SO$ins_date)
SO$approve_date <-ymd(SO$approve_date)
SO$customer_date <-ymd(SO$customer_date)
SO$TBN_app <- ymd(SO$TBN_app)
SO$tbn_finish <- ymd(SO$tbn_finish)
SO$service_date <- ymd(SO$service_date)
SO$target_date <- ymd(SO$target_date)
SO$sub_finish <- ymd(SO$sub_finish)



RK$contact_date <- ymd(RK$contact_date)
RK$bpe_date <-ymd(RK$bpe_date)
RK$ins_date <- ymd(RK$ins_date)
RK$approve_date <-ymd(RK$approve_date)
RK$customer_date <-ymd(RK$customer_date)
RK$TBN_app <- ymd(RK$TBN_app)
RK$tbn_finish <- ymd(RK$tbn_finish)
RK$service_date <- ymd(RK$service_date)
RK$target_date <- ymd(RK$target_date)
RK$sub_finish <- ymd(RK$sub_finish)

KS$contact_date <- ymd(KS$contact_date)
KS$bpe_date <-ymd(KS$bpe_date)
KS$ins_date <- ymd(KS$ins_date)
KS$approve_date <-ymd(KS$approve_date)
KS$customer_date <-ymd(KS$customer_date)
KS$TBN_app <- ymd(KS$TBN_app)
KS$tbn_finish <- ymd(KS$tbn_finish)
KS$service_date <- ymd(KS$service_date)
KS$sub_finish <- ymd(KS$sub_finish)
KS$target_date <- ymd(KS$target_date)


HO
SO
RK
KS
names(HO)
names(SO)
names(RK)
names(KS)


HO <- HO %>%
  select(1,2,22,23,3:11,12,13:29) %>%
  add_column(Branch = "HO") 

SO <- SO %>%
  select(1,2,22,23,3:11,12,13:29) %>%
  add_column(Branch = "SO")

RK <- RK %>%
  select(1,2,22,23,3:11,12,13:29) %>%
  add_column(Branch = "RK")

KS <- KS %>%
  select(1,2,22,23,3:11,12,13:29) %>%
  add_column(Branch = "KS")

#----------------------------------------
str(HO)
str(SO)
str(RK)
str(KS)



####  bind_row     ############################
bp_tbn <- bind_rows(HO,SO,RK,KS)
names(bp_tbn)






list_of_datasets <- list("HO" = HO,"SO" = SO,"RK" = RK,"KS" = KS)
write_xlsx(list_of_datasets,"TBN.xlsx")
#--------- add column Revernue -----------------
str(bp_tbn)

bp_tbn$revenue <- rowSums(bp_tbn[,c("laber", "part")], na.rm=TRUE)


#--------- add column DiffDate -----------------
# change na to today 


bp_tbn$customer_date[is.na(bp_tbn$customer_date)] <- Sys.Date()

bp_tbn$customer_date

bp_tbn$difference_in_process <-as.numeric(difftime(bp_tbn$customer_date,bp_tbn$contact_date,units = "days"))

bp_tbn$difference_in_process



str(bp_tbn)





##--------------- Summary table --------------------

library(ggplot2)
library(lubridate)

bp<- bp_tbn %>%
  select(1:6,24,25,7,8,26,27,9,10,28,29,11:23,30:32)
names(bp)
#write_xlsx(bp,"bp.xlsx")

#-------------Diff (customer_date - sub service_date )------------------------------



bp$diffservicetbn_sub <-as.numeric(difftime(bp$service_date,bp$customer_date,units = "days"))

bp$service_month <- as.numeric(format(bp$service_date,'%m'))
bp$service_year <- as.numeric(format(bp$service_date,'%Y'))

#-------------- sub_finish ---------------

bp$sub_finish_month <- as.numeric(format(bp$sub_finish,'%m'))
bp$sub_finish_year <- as.numeric(format(bp$sub_finish,'%Y'))

# add diff column -------
bp$difffinish <-as.numeric(difftime(bp$sub_finish,bp$target_date,units = "days"))
#View(bp)


bp$contact_day <- as.numeric(format(bp$contact_date,'%d'))
bp$ins_day <- as.numeric(format(bp$ins_date,'%d'))
bp$approve_day <- as.numeric(format(bp$approve_date,'%d'))
bp$customer_day <- as.numeric(format(bp$customer_date,'%d'))
bp$service_day <- as.numeric(format(bp$service_date,'%d'))
bp$sub_finishday <- as.numeric(format(bp$sub_finish,'%d'))

names(bp)


#-------------Diff (customer_date - sub service_date )------------------------------

names(bp)

bp$diffservicetbn_sub <-as.numeric(difftime(bp$service_date,bp$customer_date,units = "days"))

bp$service_month <- as.numeric(format(bp$service_date,'%m'))
bp$service_year <- as.numeric(format(bp$service_date,'%Y'))

#-------------- sub_finish ---------------

bp$sub_finish_month <- as.numeric(format(bp$sub_finish,'%m'))
bp$sub_finish_year <- as.numeric(format(bp$sub_finish,'%Y'))

# add diff column -------
bp$difffinish <-as.numeric(difftime(bp$sub_finish,bp$target_date,units = "days"))




#============================== 1.Contact Summary ====================

revenueContact042023 <- bp %>%
  group_by(Branch,contact_date,contact_day) %>%
  filter( contact_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))

# Branch
revenueContactHO042023 <- bp %>%
  group_by(Branch,contact_date,contact_day) %>%
  filter( contact_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "HO")  %>%   
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))

revenueContactSO042023 <- bp %>%
  group_by(Branch,contact_date,contact_day) %>%
  filter( contact_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "SO")  %>%   
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))

revenueContactRK042023 <- bp %>%
  group_by(Branch,contact_date,contact_day) %>%
  filter( contact_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "RK")  %>%   
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))

revenueContactKS042023 <- bp %>%
  group_by(Branch,contact_date,contact_day) %>%
  filter( contact_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "KS")  %>%   
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))








#==== 2.sent Insurance ====================

sentInsurance042023 <- bp %>%
  group_by(Branch,ins_date) %>%
  filter( ins_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue)  )

# branch
sentInsuranceHO042023 <- bp %>%
  group_by(Branch,ins_date,ins_day) %>%
  filter( ins_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "HO")  %>%    
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue)  )

sentInsuranceSO042023 <- bp %>%
  group_by(Branch,ins_date,ins_day) %>%
  filter( ins_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "SO")  %>%    
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue)  )

sentInsuranceRK042023 <- bp %>%
  group_by(Branch,ins_date,ins_day) %>%
  filter( ins_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "RK")  %>%    
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue)  )

sentInsuranceKS042023 <- bp %>%
  group_by(Branch,ins_date,ins_day) %>%
  filter( ins_year== "2023")  %>%
  filter( contact_month== "4")  %>%  
  filter( Branch== "KS")  %>%    
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue)  )



names(bp)


ins_approve042023 <- bp %>%
  group_by(Branch,approve_date,approve_day)  %>%
  filter( approve_year== "2023")  %>%
  filter( approve_month== "4")  %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue))

# branch
ins_approveHO042023 <- bp %>%
  group_by(Branch,approve_date,approve_day)  %>%
  filter( approve_year== "2023")  %>%
  filter( approve_month== "4")  %>%
  filter(Branch == "HO") %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue))

ins_approveSO042023 <- bp %>%
  group_by(Branch,approve_date,approve_day)  %>%
  filter( approve_year== "2023")  %>%
  filter( approve_month== "4")  %>%
  filter(Branch == "SO") %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue))

ins_approveRK042023 <- bp %>%
  group_by(Branch,approve_date,approve_day)  %>%
  filter( approve_year== "2023")  %>%
  filter( approve_month== "4")  %>%
  filter(Branch == "RK") %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue))

ins_approveKS042023 <- bp %>%
  group_by(Branch,approve_date,approve_day)  %>%
  filter( approve_year== "2023")  %>%
  filter( approve_month== "4")  %>%
  filter(Branch == "KS") %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue))




#View(ins_approve)


names(bp)

customer042023 <- bp %>%
  group_by(Branch,customer_date,customer_day) %>%
  filter( customer_year== "2023")  %>%
  filter(customer_month=="4")  %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue))

# branch

customerHO042023 <- bp %>%
  group_by(Branch,customer_date,customer_day) %>%
  filter( customer_year== "2023")  %>%
  filter(customer_month=="4")  %>%
  filter(Branch == "HO") %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue))

customerSO042023 <- bp %>%
  group_by(Branch,customer_date,customer_day) %>%
  filter( customer_year== "2023")  %>%
  filter(customer_month=="4")  %>%
  filter(Branch == "SO") %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue))

customerRK042023 <- bp %>%
  group_by(Branch,customer_date,customer_day) %>%
  filter( customer_year== "2023")  %>%
  filter(customer_month=="4")  %>%
  filter(Branch == "RK") %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue))

customerKS042023 <- bp %>%
  group_by(Branch,customer_date,customer_day) %>%
  filter( customer_year== "2023")  %>%
  filter(customer_month=="4")  %>%
  filter(Branch == "KS") %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue))






tbn_sub042023 <- bp %>%
  group_by(Branch,service_date,service_day) %>%
  filter(service_year=="2023") %>%
  filter(service_month=="4") %>%
  summarize(count_subservice= n(),revenue_subservice= sum(revenue))

# branch
tbn_subHO042023 <- bp %>%
  group_by(Branch,service_date,service_day) %>%
  filter(service_year=="2023") %>%
  filter(service_month=="4") %>%
  filter(Branch == "HO") %>%
  summarize(count_subservice= n(),revenue_subservice= sum(revenue))

tbn_subSO042023 <- bp %>%
  group_by(Branch,service_date,service_day) %>%
  filter(service_year=="2023") %>%
  filter(service_month=="4") %>%
  filter(Branch == "SO") %>%
  summarize(count_subservice= n(),revenue_subservice= sum(revenue))

tbn_subRK042023 <- bp %>%
  group_by(Branch,service_date,service_day) %>%
  filter(service_year=="2023") %>%
  filter(service_month=="4") %>%
  filter(Branch == "RK") %>%
  summarize(count_subservice= n(),revenue_subservice= sum(revenue))

tbn_subKS042023 <- bp %>%
  group_by(Branch,service_date,service_day) %>%
  filter(service_year=="2023") %>%
  filter(service_month=="4") %>%
  filter(Branch == "KS") %>%
  summarize(count_subservice= n(),revenue_subservice= sum(revenue))





names(bp)



sub_finishservice042023 <- bp %>%
  group_by(Branch,sub_finish,sub_finishday) %>%
  filter(sub_finish_year=="2023") %>%
  filter(sub_finish_month=="4") %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))

# branch 
sub_finishserviceHO042023 <- bp %>%
  group_by(Branch,sub_finish,sub_finishday) %>%
  filter(sub_finish_year=="2023") %>%
  filter(sub_finish_month=="4") %>%
  filter(Branch=="HO") %>% 
  summarize(count_sub_finish= n(),revenue_sub_finish= sum(revenue))

sub_finishserviceSO042023 <- bp %>%
  group_by(Branch,sub_finish,sub_finishday) %>%
  filter(sub_finish_year=="2023") %>%
  filter(sub_finish_month=="4") %>%
  filter(Branch=="SO") %>% 
  summarize(count_sub_finish= n(),revenue_sub_finish= sum(revenue))

sub_finishserviceRK042023 <- bp %>%
  group_by(Branch,sub_finish,sub_finishday) %>%
  filter(sub_finish_year=="2023") %>%
  filter(sub_finish_month=="4") %>%
  filter(Branch=="RK") %>% 
  summarize(count_sub_finish= n(),revenue_sub_finish= sum(revenue))

sub_finishserviceKS042023 <- bp %>%
  group_by(Branch,sub_finish,sub_finishday) %>%
  filter(sub_finish_year=="2023") %>%
  filter(sub_finish_month=="4") %>%
  filter(Branch=="KS") %>% 
  summarize(count_sub_finish= n(),revenue_sub_finish= sum(revenue))

names(bp)



### ------------- total summary   ------------------------
names(bp)

setwd("/Users/parinya/Library/Mobile Documents/com~apple~CloudDocs/1.TBN/BP/1.BP Daily/042023")


daymaster <- read_excel("day.xlsx",sheet = 1)

#View(revenueContactHO042023)



#================== Day =====================================
HO <- daymaster %>%
  left_join(revenueContactHO042023,by = c("day" = "contact_day")) %>%
  left_join(sentInsuranceHO042023,by = c("day" = "ins_day")) %>%
  left_join(ins_approveHO042023,by = c("day" = "approve_day")) %>%
  left_join(customerHO042023,by = c("day" = "customer_day")) %>%
  left_join(tbn_subHO042023,by = c("day" = "service_day")) %>%
  left_join(sub_finishserviceHO042023,by = c("day" = "sub_finishday")) 

names(HO)

countHO <- HO %>%
  select(1,Branch=2,4,8,12,16,count_insub=20,24) 

revenueHO <- HO %>%
  select(1,Branch=2,5,9,13,17,21,25) 

#View(countHO)

#=======================================================
SO <- daymaster %>%
  left_join(revenueContactSO042023,by = c("day" = "contact_day")) %>%
  left_join(sentInsuranceSO042023,by = c("day" = "ins_day")) %>%
  left_join(ins_approveSO042023,by = c("day" = "approve_day")) %>%
  left_join(customerSO042023,by = c("day" = "customer_day")) %>%
  left_join(tbn_subSO042023,by = c("day" = "service_day")) %>%
  left_join(sub_finishserviceSO042023,by = c("day" = "sub_finishday")) 

names(SO)

countSO <- SO %>%
  select(1,Branch=2,4,8,12,16,count_insub=20,24) 

revenueSO <- SO %>%
  select(1,Branch=2,5,9,13,17,21,25) 

#View(countSO)
#=======================================================
RK <- daymaster %>%
  left_join(revenueContactRK042023,by = c("day" = "contact_day")) %>%
  left_join(sentInsuranceRK042023,by = c("day" = "ins_day")) %>%
  left_join(ins_approveRK042023,by = c("day" = "approve_day")) %>%
  left_join(customerRK042023,by = c("day" = "customer_day")) %>%
  left_join(tbn_subRK042023,by = c("day" = "service_day")) %>%
  left_join(sub_finishserviceRK042023,by = c("day" = "sub_finishday")) 

names(RK)

countRK <- RK %>%
  select(1,Branch=2,4,8,12,16,count_insub=20,24) 

revenueRK <- RK %>%
  select(1,Branch=2,5,9,13,17,21,25) 

#View(countRK)
#=======================================================

KS <- daymaster %>%
  left_join(revenueContactKS042023,by = c("day" = "contact_day")) %>%
  left_join(sentInsuranceKS042023,by = c("day" = "ins_day")) %>%
  left_join(ins_approveKS042023,by = c("day" = "approve_day")) %>%
  left_join(customerKS042023,by = c("day" = "customer_day")) %>%
  left_join(tbn_subKS042023,by = c("day" = "service_day")) %>%
  left_join(sub_finishserviceKS042023,by = c("day" = "sub_finishday")) 

names(KS)

countKS <- KS %>%
  select(1,Branch=2,4,8,12,16,count_insub=20,24) 

revenueKS <- KS %>%
  select(1,Branch=2,5,9,13,17,21,25) 

#View(countKS)
#=======================================================

# computing column wise sum
data_without_na <- countKS [3:8] %>%                      
  replace(is.na(.), 0) 


#replace NA values with zero in rebs and pts columns

# countKS <- countKS %>% mutate(
#   count_by_contact = ifelse(is.na(count_by_contact), 0, count_by_contact),
#   count_by_insurance = ifelse(is.na(count_by_insurance), 0, count_by_insurance),
#   count_by_ins_approve = ifelse(is.na(count_by_ins_approve), 0, count_by_ins_approve),
#   count_by_customerIncome = ifelse(is.na(count_by_customerIncome), 0, count_by_customerIncome),
#   count_insub = ifelse(is.na(count_insub), 0, count_insub),
#   count_sub_finish = ifelse(is.na(count_sub_finish), 0, count_sub_finish),
#   )


names(revenueKS)

#colSums(countKS, na.rm=TRUE)
HOsum <-colSums(countHO[, c(3:8)], na.rm=TRUE)
SOsum <-colSums(countSO[, c(3:8)], na.rm=TRUE)
RKsum <-colSums(countRK[, c(3:8)], na.rm=TRUE)
KSsum <-colSums(countKS[, c(3:8)], na.rm=TRUE)

tbnsum <- bind_rows(HOsum,SOsum,RKsum,KSsum)
tbnsum <- tbnsum %>%
  mutate(br=c("HO","SO","RK","KS"))

View(tbnsum)

HOrevsum <-colSums(revenueHO[, c(3:8)], na.rm=TRUE)
SOrevsum <-colSums(revenueSO[, c(3:8)], na.rm=TRUE)
RKrevsum <-colSums(revenueRK[, c(3:8)], na.rm=TRUE)
KSrevsum <-colSums(revenueKS[, c(3:8)], na.rm=TRUE)

tbnrevsum <- bind_rows(HOrevsum,SOrevsum,RKrevsum,KSrevsum)

tbnrevsum <- tbnrevsum %>%
  mutate(br=c("HO","SO","RK","KS"))

str(tbnsum)




#=======================================================

setwd("/Users/parinya/Library/Mobile Documents/com~apple~CloudDocs/1.TBN/BP/1.BP Daily/042023/Daily status bp")




write_xlsx(countHO ,"countHO.xlsx")
write_xlsx(countSO ,"countSO.xlsx")
write_xlsx(countRK ,"countRK.xlsx")
write_xlsx(countKS ,"countKS.xlsx")

write_xlsx(revenueHO ,"revenueHO.xlsx")
write_xlsx(revenueSO ,"revenueSO.xlsx")
write_xlsx(revenueRK ,"revenueRK.xlsx")
write_xlsx(revenueKS ,"revenueKS.xlsx")

write_xlsx(tbnsum ,"tbnsum.xlsx")
write_xlsx(tbnrevsum ,"tbnrevsum.xlsx")


write_xlsx(revenueContact042023 ,"a1.Contact042023.xlsx")
write_xlsx(sentInsurance042023,"a2.sentInsurance042023.xlsx")
write_xlsx(ins_approve042023,"a3.ins_approve042023.xlsx")
write_xlsx(customer042023,"a4.customerIncome042023.xlsx")
write_xlsx(tbn_sub042023,"a6.subservice042023.xlsx")
write_xlsx(sub_finishservice042023,"7.sub_finishservice042023.xlsx")








# write_xlsx(revenueContactHO042023 ,"1.ContactHO.xlsx")
# write_xlsx(revenueContactSO042023 ,"1.ContactSO.xlsx")
# write_xlsx(revenueContactRK042023 ,"1.ContactRK.xlsx")
# write_xlsx(revenueContactKS042023 ,"1.ContactKS.xlsx")
# 
# write_xlsx(sentInsuranceHO042023 ,"2.sentInsuranceHO.xlsx")
# write_xlsx(sentInsuranceSO042023 ,"2.sentInsuranceSO.xlsx")
# write_xlsx(sentInsuranceRK042023 ,"2.sentInsuranceRK.xlsx")
# write_xlsx(sentInsuranceKS042023 ,"2.sentInsuranceKS.xlsx")
# 
# 
# write_xlsx(ins_approveHO042023 ,"3.ins_approveHO.xlsx")
# write_xlsx(ins_approveSO042023 ,"3.ins_approveSO.xlsx")
# write_xlsx(ins_approveRK042023 ,"3.ins_approveRK.xlsx")
# write_xlsx(ins_approveKS042023 ,"3.ins_approveKS.xlsx")
# 
# 
# write_xlsx(customerHO042023 ,"4.customerHO.xlsx")
# write_xlsx(customerSO042023 ,"4.customerSO.xlsx")
# write_xlsx(customerRK042023 ,"4.customerRK.xlsx")
# write_xlsx(customerKS042023 ,"4.customerKS.xlsx")
# 
# write_xlsx(tbn_subHO042023 ,"6.subserviceHO.xlsx")
# write_xlsx(tbn_subSO042023 ,"6.subserviceSO.xlsx")
# write_xlsx(tbn_subRK042023 ,"6.subserviceRK.xlsx")
# write_xlsx(tbn_subKS042023 ,"6.subserviceKS.xlsx")
# 
# 
# write_xlsx(sub_finishserviceHO042023 ,"7.sub_finishserviceHO.xlsx")
# write_xlsx(sub_finishserviceSO042023 ,"7.sub_finishserviceSO.xlsx")
# write_xlsx(sub_finishserviceRK042023 ,"7.sub_finishserviceRK.xlsx")
# write_xlsx(sub_finishserviceKS042023 ,"7.sub_finishserviceKS.xlsx")
