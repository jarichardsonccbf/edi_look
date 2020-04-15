library(tidyverse)
library(data.table)

source("method.R")

mmdata <- fread("data/mm_data.csv")

valid_column_names <- make.names(names=names(mmdata), unique=TRUE, allow_ = TRUE)
names(mmdata) <- valid_column_names

mmdata <- mmdata %>%
  filter(Status == "ACTIVE") %>%
  filter(BusinessType == "DSD") %>%
  select(Name, PreferredOrderMethod, SubChannel, Channel, KeyAccount, Volume, Dead.Net.Revenue, Dead.Net.Gross.Profit) %>%
  filter(Volume > 0)

# KA.summary <- mmdata %>%
#  select(PreferredOrderMethod, KeyAccount, Channel) %>%
#  unique()

rm(valid_column_names)

# Deal with club coke CR

cc.cr <- mmdata %>%
  filter(KeyAccount == "CLUB COKE - HOME MAR" & Channel == "Convenience Store/Pe") %>%
  mutate(Ordering.Method = "EDID")

mmdata2 <- mmdata %>%
  filter(KeyAccount != "CLUB COKE - HOME MAR" | Channel != "Convenience Store/Pe")

# Insert ordering method based on key account

ka.joins <- mmdata2 %>%
  left_join(KA, by = "KeyAccount") %>%
  filter(!is.na(Ordering.Method))

# parse out outlets that don't have key account info

non.ka <- mmdata2 %>%
  left_join(KA, by = "KeyAccount") %>%
  filter(is.na(Ordering.Method)) %>%
  select(-c("Ordering.Method"))

# add in channel ordering method if present

channel.joins <- non.ka %>%
  left_join(channel, by = "Channel") %>%
  filter(!is.na(Ordering.Method))

therest.joins <- non.ka %>%
  left_join(channel, by = "Channel") %>%
  filter(is.na(Ordering.Method)) %>%
  select(-c("Ordering.Method")) %>%
  mutate(Ordering.Method = PreferredOrderMethod)

df <- rbind(ka.joins, channel.joins, therest.joins, cc.cr)

rm(channel, channel.joins, KA, ka.joins, mmdata, non.ka, therest.joins, mmdata2, cc.cr)

outlet.summary <- df %>%
  group_by(Ordering.Method) %>%
  summarise(n = n()) %>%
  mutate(Percent = n / sum(n) * 100) %>%
  rename(`Number of Outlets` = n)

volume.summary <- df %>%
  group_by(Ordering.Method) %>%
  summarise(Volume = sum(as.numeric(Volume), na.rm = TRUE)) %>%
  mutate(Percent = Volume / sum(Volume) * 100)

library(xlsx)

options(java.parameters = "-Xmx1000m")

write.xlsx(df, file = "deliverables/order_methods.xlsx", sheetName = "Master Data", row.names = FALSE)

write.xlsx(as.data.frame(outlet.summary), file = "deliverables/order_methods.xlsx", sheetName= "Method Summary", append = TRUE, row.names = FALSE)

write.xlsx(as.data.frame(volume.summary), file = "deliverables/order_methods.xlsx", sheetName= "Volume Summary", append = TRUE, row.names = FALSE)
