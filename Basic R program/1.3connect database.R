

getwd()
list.files()


# package 
# tidyverse RSQLite RPostgreSQL lubridate


library(tidyverse )
library(RSQLite)
library(RPostgreSQL)
library(lubridate)

# connect database

dbConnect(SQLite() , "chinook.db")
con <-dbConnect(SQLite() , "chinook.db")
con

## list table name เรียกดู Table in db
dbListTables(con)

## list field in a table  # field = column มีคอลัมน์อะไรบ้างใน Table
dbListFields(con , "customers")

## 
df <-dbGetQuery(con , "select * from customers limit 50")

df


clean_df <- clean_names(df)

df2 = dbGetQuery(con , "select * from albums, artists
                 where albums.artistid = artists.artistid")
df2


##################

# write a table

dbWriteTable(con , "cars" , mtcars)
dbListTables(con)
dbGetQuery(con , "select * from cars limit 5;")


## close connection
dbDisconnect(con)





