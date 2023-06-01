library(readxl)
library(dplyr)
library(tidyverse)
# import writexl library
library(writexl)
library(WriteXLS)
library(ggplot2)
library(lubridate)


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
# df <- TBN_HO
# 
# 
# df$Date_difference_in_days <-as.numeric(difftime(df$contact_date,df$customer_date,units = "days"))
# df$Date_difference_in_days 





str(TBN_HO)









names(TBN_HO)

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

#TBN_SO$approve_date           # delete "ติดเรื่องกรมธรรม์" 



# TBN_SO$approve_date  <- sub("ติดเรื่องกรมธรรม์", "",TBN_SO$approve_date)

# TBN_SO$approve_date  %>%
#   as.Date(approve, origin = "1900-01-01")





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

#------------------------------------------------------------------








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

## na === bp_tbn$customer_date[is.na(bp_tbn$customer_date)]
## today ===Sys.Date()

bp_tbn$customer_date[is.na(bp_tbn$customer_date)] <- Sys.Date()

bp_tbn$customer_date

bp_tbn$difference_in_process <-as.numeric(difftime(bp_tbn$customer_date,bp_tbn$contact_date,units = "days"))

bp_tbn$difference_in_process



str(bp_tbn)




#View(bp_tbn)
#write_xlsx(bp_tbn,"bp_tbn.xlsx")



# bp_tbn$customer_date <- as.POSIXct(strptime(as.character(bp_tbn$customer_date),"%Y-%m-%d"))
# bp_tbn$contact_date <- as.POSIXct(strptime(as.character(bp_tbn$contact_date),"%Y-%m-%d"))
# 
# class(bp_tbn$customer_date)
# diff_dates <- as.numeric(difftime(bp_tbn$customer_date ,bp_tbn$contact_date,tz, units = "days"))
# 
# 
# days_diff <- bp_tbn %>%
#   as.numeric(difftime(customer_date, contact_date))
# 



##--------------- Summary table --------------------

library(ggplot2)
library(lubridate)
names(bp_tbn)
bp<- bp_tbn %>%
  select(1:6,24,25,7,8,26,27,9,10,28,29,11:23,30:32)
names(bp)
write_xlsx(bp,"bp.xlsx")
#View(bp)
# แยกเดือน

# 1.table : จำนวนรถ n()-> แต่ละสถานะ
# 2.table : รถเข้าซ่อม n() แต่ละเดือน - แต่ละสาขา
# 3.table : ระยะเวลารอ แต่ละสถานะ

# 1.table n() : 
# 1.1 group_by one variable

total_contact <- bp %>%
  summarise(total_contact=n()) 

t_contact <- bp %>%
  group_by(Branch,contact_year) %>%
  summarize(count_by_contact_year = n())

print(total_contact)
print(t_contact)

# ** 1.2 group_by multiple variables **
names(bp)

customerContact <- bp %>%
  group_by(Branch,contact_year,contact_month) %>%
  summarize(count_by_contact=  n())

#==== 1.Contact ====================
revenueContact2022 <- bp %>%
  group_by(Branch,contact_year,contact_month) %>%
  filter( contact_year== "2022")  %>%
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))

revenueContact2023 <- bp %>%
  group_by(Branch,contact_year,contact_month) %>%
  filter( contact_year== "2023")  %>%
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))

revenueContact <- bp %>%
  group_by(contact_year,Branch,contact_month) %>%
  arrange(desc(contact_year) )  %>%
  summarize(count_by_contact= n(),
            revenue_Contact = sum(revenue))

write_xlsx(revenueContact2023 ,"1.Contact2023 .xlsx")
write_xlsx(revenueContact2022 ,"1.1Contact2022.xlsx")
write_xlsx(revenueContact ,"w1.Contact.xlsx")
#View(revenueContact2023)

sentInsurance2022 <- bp %>%
  group_by(Branch,ins_year,ins_month) %>%
  filter( ins_year== "2022")  %>%
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue),
            avg_waitinsurance_days = mean(wait_ins, na.rm=TRUE),
            min_waitinsurance_days = min(wait_ins, na.rm=TRUE),
            max_waitinsurance_days = max(wait_ins, na.rm=TRUE)
  )
sentInsurance2023 <- bp %>%
  group_by(Branch,ins_year,ins_month) %>%
  filter( ins_year== "2023")  %>%
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue),
            avg_waitinsurance_days = mean(wait_ins, na.rm=TRUE),
            min_waitinsurance_days = min(wait_ins, na.rm=TRUE),
            max_waitinsurance_days = max(wait_ins, na.rm=TRUE)
  )

sentInsuranceNohandle <- bp %>%
  group_by(Branch,ins_year,ins_month) %>%
  filter( is.na(ins_year)) %>%
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue),
            avg_waitinsurance_days = mean(wait_ins, na.rm=TRUE),
            min_waitinsurance_days = min(wait_ins, na.rm=TRUE),
            max_waitinsurance_days = max(wait_ins, na.rm=TRUE)
  )

sentInsurance<- bp %>%
  group_by(ins_year,Branch,ins_month) %>%
  arrange(ins_year) %>%
  summarize(count_by_insurance=  n(),
            revenue_insurance= sum(revenue),
            avg_waitinsurance_days = mean(wait_ins, na.rm=TRUE),
            min_waitinsurance_days = min(wait_ins, na.rm=TRUE),
            max_waitinsurance_days = max(wait_ins, na.rm=TRUE)
  )

write_xlsx(sentInsurance2023,"2.1sentInsurance2023.xlsx")
write_xlsx(sentInsurance2022,"2.2sentInsurance2022.xlsx")
write_xlsx(sentInsuranceNohandle,"2.3sentInsuranceNohandle.xlsx")
write_xlsx(sentInsurance,"w2.sentInsurance.xlsx")
#View(sentInsurance)
names(bp)

ins_approve2023 <- bp %>%
  group_by(Branch,approve_year,approve_month)  %>%
  filter( approve_year== "2023")  %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue),
            avg_waitapprove_days = mean(wait_approve, na.rm=TRUE),
            min_waitapprove_days = min(wait_approve, na.rm=TRUE),
            max_waitapprove_days = max(wait_approve, na.rm=TRUE)
  )
ins_approve2022 <- bp %>%
  group_by(Branch,approve_year,approve_month)  %>%
  filter( approve_year== "2022")  %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue),
            avg_waitapprove_days = mean(wait_approve, na.rm=TRUE),
            min_waitapprove_days = min(wait_approve, na.rm=TRUE),
            max_waitapprove_days = max(wait_approve, na.rm=TRUE)
  )

ins_approve2022 <- bp %>%
  group_by(Branch,approve_year,approve_month)  %>%
  filter( approve_year== "2022")  %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue),
            avg_waitapprove_days = mean(wait_approve, na.rm=TRUE),
            min_waitapprove_days = min(wait_approve, na.rm=TRUE),
            max_waitapprove_days = max(wait_approve, na.rm=TRUE)
  )

ins_approveNohandle <- bp %>%
  group_by(Branch,approve_year,approve_month)  %>%
  filter( is.na(approve_year)) %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue),
            avg_waitapprove_days = mean(wait_approve, na.rm=TRUE),
            min_waitapprove_days = min(wait_approve, na.rm=TRUE),
            max_waitapprove_days = max(wait_approve, na.rm=TRUE))

ins_approve <- bp %>%
  group_by(approve_year,Branch,approve_month)  %>%
  arrange( desc(approve_year)) %>%
  summarize(count_by_ins_approve=  n(),
            revenue_ins_approve= sum(revenue),
            avg_waitapprove_days = mean(wait_approve, na.rm=TRUE),
            min_waitapprove_days = min(wait_approve, na.rm=TRUE),
            max_waitapprove_days = max(wait_approve, na.rm=TRUE))

write_xlsx(ins_approve2023,"3.1ins_approve2023.xlsx")
write_xlsx(ins_approve2022,"3.2ins_approve2022.xlsx")
write_xlsx(ins_approveNohandle,"3.3ins_approveNohandle.xlsx")
write_xlsx(ins_approve,"w3.ins_approve.xlsx")
#View(ins_approve)


names(bp)
customerIncome <- bp %>%
  group_by(Branch,customer_year,customer_month) %>%
  summarize(count_by_customerIncome=  n())

#View(customerIncome)


customer2023 <- bp %>%
  group_by(Branch,customer_year,customer_month) %>%
  filter( approve_year== "2023")  %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue),
            avg_WaitCustomer_days = mean(wait_customer, na.rm=TRUE),
            min_WaitCustomer_days = min(wait_customer, na.rm=TRUE),
            max_WaitCustomer_days = max(wait_customer, na.rm=TRUE)
  )

customer2022 <- bp %>%
  group_by(Branch,customer_year,customer_month) %>%
  filter( approve_year== "2022")  %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue),
            avg_WaitCustomer_days = mean(wait_customer, na.rm=TRUE),
            min_WaitCustomer_days = min(wait_customer, na.rm=TRUE),
            max_WaitCustomer_days = max(wait_customer, na.rm=TRUE)
  )

customerNohandle <- bp %>%
  group_by(Branch,customer_year,customer_month) %>%
  filter( is.na(customer_year)) %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue),
            avg_WaitCustomer_days = mean(wait_customer, na.rm=TRUE),
            min_WaitCustomer_days = min(wait_customer, na.rm=TRUE),
            max_WaitCustomer_days = max(wait_customer, na.rm=TRUE)
  )

customer <- bp %>%
  group_by(customer_year,Branch,customer_month) %>%
  arrange(desc(customer_year)) %>%
  summarize(count_by_customerIncome=  n(),
            revenue_customerIncome= sum(revenue),
            avg_WaitCustomer_days = mean(wait_customer, na.rm=TRUE),
            min_WaitCustomer_days = min(wait_customer, na.rm=TRUE),
            max_WaitCustomer_days = max(wait_customer, na.rm=TRUE))

write_xlsx(customer2023,"4.customerIncome2023.xlsx")
write_xlsx(customer2022,"4.customerIncome2022.xlsx")
write_xlsx(customerNohandle,"4.customerIncomeNohandle.xlsx")
write_xlsx(customer,"w4.customerIncome.xlsx")

#View(customer)

# totalprocess2023 <- bp %>%
#   group_by(Branch,customer_year,customer_month) %>%
#   filter(customer_year=="2023") %>%
#   summarize(count_by_totalprocess=  n(),
#             avg_process_days = mean(difference_in_process, na.rm=TRUE),
#             min_process_days = min(difference_in_process, na.rm=TRUE),
#             max_process_days = max(difference_in_process, na.rm=TRUE))

revenue_inprocess2023 <- bp %>%
  group_by(Branch,customer_year,customer_month) %>%
  filter(customer_year=="2023") %>%
  summarize(count_by_totalprocess= n(),revenue_inprocess= sum(revenue),
            avg_process_days = mean(difference_in_process, na.rm=TRUE),
            min_process_days = min(difference_in_process, na.rm=TRUE),
            max_process_days = max(difference_in_process, na.rm=TRUE))


revenue_inprocessNohandle <- bp %>%
  group_by(Branch,customer_year,customer_month) %>%
  filter(is.na(customer_year)) %>%
  summarize(count_by_totalprocess= n(),revenue_inprocess= sum(revenue),
            avg_process_days = mean(difference_in_process, na.rm=TRUE),
            min_process_days = min(difference_in_process, na.rm=TRUE),
            max_process_days = max(difference_in_process, na.rm=TRUE))

revenue_inprocess <- bp %>%
  group_by(customer_year,Branch,customer_month) %>%
  arrange(desc(customer_year)) %>%
  summarize(count_by_totalprocess= n(),revenue_inprocess= sum(revenue),
            avg_process_days = mean(difference_in_process, na.rm=TRUE),
            min_process_days = min(difference_in_process, na.rm=TRUE),
            max_process_days = max(difference_in_process, na.rm=TRUE))

bp$difference_in_process
#View(totalprocess)
names(bp)


write_xlsx(revenue_inprocess2023,"5.total_tbnprocess2023.xlsx")
write_xlsx(revenue_inprocessNohandle,"5.total_tbnprocessNohandle.xlsx")
write_xlsx(revenue_inprocess,"w5.total_tbnprocess.xlsx")
#------------- Choose branch ----------------
ch_branch <- customer %>% 
  filter(Branch == "HO")

#View(ch_branch)

#-------------Diff (customer_date - sub service_date )------------------------------

names(bp)

bp$diffservicetbn_sub <-as.numeric(difftime(bp$service_date,bp$customer_date,units = "days"))

bp$service_month <- as.numeric(format(bp$service_date,'%m'))
bp$service_year <- as.numeric(format(bp$service_date,'%Y'))

tbn_sub2023 <- bp %>%
  group_by(Branch,service_year,service_month) %>%
  filter(service_year=="2023") %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))

tbn_sub2022 <- bp %>%
  group_by(Branch,service_year,service_month) %>%
  filter(service_year=="2022") %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))

tbn_subNohandle <- bp %>%
  group_by(Branch,service_year,service_month) %>%
  filter(is.na(service_year))%>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))


tbn_sub <- bp %>%
  group_by(service_year,Branch,service_month) %>%
  arrange(desc(service_year)) %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))


write_xlsx(tbn_sub2023,"6.subservice2023.xlsx")
write_xlsx(tbn_sub2022,"6.subservice2022.xlsx")
write_xlsx(tbn_subNohandle,"6.subNohandle.xlsx")
write_xlsx(tbn_sub,"w6.subservice.xlsx")

# tbn_sub2 <- bp %>%
#   group_by(Branch,customer_year,service_year,service_month) %>%
#   summarize(count_insub= n(),revenue_insub= sum(revenue))
# ,
#             avg_process_days = mean(difference_in_process, na.rm=TRUE),
#             min_process_days = min(difference_in_process, na.rm=TRUE),
#             max_process_days = max(difference_in_process, na.rm=TRUE))

View(tbn_sub)



names(bp)

#-------------- sub_finish ---------------

bp$sub_finish_month <- as.numeric(format(bp$sub_finish,'%m'))
bp$sub_finish_year <- as.numeric(format(bp$sub_finish,'%Y'))

# add diff column -------
bp$difffinish <-as.numeric(difftime(bp$sub_finish,bp$target_date,units = "days"))




sub_finishservice2023 <- bp %>%
  group_by(Branch,sub_finish_year,sub_finish_month) %>%
  filter(sub_finish_year=="2023") %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))

sub_finishservice2022 <- bp %>%
  group_by(Branch,sub_finish_year,sub_finish_month) %>%
  filter(sub_finish_year=="2022") %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))

sub_finishserviceNohandle <- bp %>%
  group_by(Branch,sub_finish_year,sub_finish_month) %>%
  filter(is.na(sub_finish_year)) %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue))

sub_finishservice <- bp %>%
  group_by(sub_finish_year,Branch,sub_finish_month) %>%
  arrange(desc(sub_finish_year)) %>%
  summarize(count_insub= n(),revenue_insub= sum(revenue),
            avg_due_days = mean(difffinish, na.rm=TRUE),
            min_due_days = min(difffinish, na.rm=TRUE),
            max_due_days = max(difffinish, na.rm=TRUE))

names(bp)
View(sub_finishservice)
write_xlsx(sub_finishservice2023,"7.sub_finishservice2023.xlsx")
write_xlsx(sub_finishservice2022,"7.sub_finishservice2022.xlsx")
write_xlsx(sub_finishserviceNohandle,"7.sub_finishserviceNohandle.xlsx")
write_xlsx(sub_finishservice,"w7.sub_finishservice.xlsx")




