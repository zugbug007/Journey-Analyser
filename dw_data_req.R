# Load the library
library(RSiteCatalyst)
library(lubridate)
library(stringr)
library(readr)
library(dplyr)

rsid <- Sys.getenv("ADOBE_RSID")
options(scipen=999)

# Set the ME report date range (YYYY-MM-DD format)
me_project_name       <- "Member_Engagement_"
# me_date_from          <- "2025-10-09"
# me_date_to            <- "2025-10-16"

# August ME
me_date_from          <- "2025-08-14"
me_date_to            <- "2025-08-21"

me_segment_id         <- "s1957_68e8ed503e029e4199c43afc"   # ME: 09102025
me_excel_filename     <- paste0("data/",me_project_name,me_date_from,"-",me_date_to, ".csv")

# Set the Prospects report date range (YYYY-MM-DD format)
prospects_project_name    <- "Prospect_Email_"
# prospects_date_from       <- "2025-10-16"
# prospects_date_to         <- "2025-10-23"

# August
prospects_date_from       <- "2025-08-21"
prospects_date_to         <- "2025-08-31"

prospects_segment_id      <- "s1957_68d518bee6151f1d5f70d9c9"   # Prospect Email: 16102025
prospects_excel_filename  <- paste0("data/",prospects_project_name,prospects_date_from,"-",prospects_date_to, ".csv")

# Set Granularity, Metrics & Dimensions
metrics <- c("pageviews", "visits", "uniquevisitors", "revenue", "event114", "event5", "event125", "event79", "event10")
elements <- c("evar29", "trackingcode", "evar26", "evar8", "evar18", "evar7", "evar22", "visitnumber", "linkexit")
granularity <- "day"  #options: "day", "week", "month", "year", "hour"

# Authenticate
SCAuth(Sys.getenv("ADOBE_API_USERNAME"), Sys.getenv("ADOBE_API_SECRET"))

# START MEMBER ENGAMEMENT DATA REQUEST =======================================================

# Submit the Member Engagement Data Warehouse Request
member_engagement_report_data <- QueueDataWarehouse(
    reportsuite.id = rsid,
    date.from = me_date_from,
    date.to = me_date_to,
    segment.id = me_segment_id,
    metrics = metrics,
    elements = elements,
    enqueueOnly = FALSE,
    date.granularity = granularity
 )

member_engagement_report_data2 <- member_engagement_report_data %>% 
  rename(`MCVID (v29) (evar29)` = MCVID..v29...evar29.,
         `Tracking Code` = Tracking.Code,
         `Page Name (v26) (evar26)` = Page.Name..v26...evar26.,
         `URL (v8) (evar8)` = URL..v8...evar8.,
         `24H Clock by Minute` = Time.Parting..v18...evar18.,
         `Content Type (v7) (evar7)` = Content.Type..v7...evar7.,
         `Content Title (v22) (evar22)` = Content.Title..v22...evar22.,
         `Visit Number` = Visit.Number,
         `Exit Links` = Exit.Links,
         `Page Views` = Page.Views,
         `Unique Visitors` = Unique.Visitors,
         `Shop - Revenue` = Shop...Revenue,
         `Donate Revenue (Serialized) (ev114) (event114)` = Donate.Revenue..Serialized...ev114...event114.,
         `Membership Revenue (ev5) (event5)` = Membership.Revenue..ev5...event5.,
         `Holidays Booking Total Revenue (Serialised) (ev125) (event125)` = Holidays.Booking.Total.Revenue..Serialised...ev125...event125.,
         `Renew Revenue - Serialized (ev79) (event79)` = Renew.Revenue...Serialized..ev79...event79.,
         `Video Start (ev10) (event10)` = Video.Start..ev10...event10.) %>% 
  mutate(
    Time_12H = str_extract(`24H Clock by Minute`, "^[^|]+"),
    `24H Clock by Minute` = format(as.POSIXct(Time_12H,format='%I:%M %p'),format="%H:%M")
  ) %>% 
  select(-Time_12H)

write.csv(member_engagement_report_data2,me_excel_filename, row.names = FALSE)

# END OF MEMBER ENGAGEMENT DATA REQUEST ======================================================================================================

# START PROSPECT DATA REQUEST           ======================================================================================================

# Submit the Prospect Data Warehouse Request
prospect_report_data <- QueueDataWarehouse(
  reportsuite.id = rsid,
  date.from = prospects_date_from,
  date.to = prospects_date_to,
  segment.id = prospects_segment_id,
  metrics = metrics,
  elements = elements,
  enqueueOnly = FALSE,
  date.granularity = granularity
)

prospect_report_data2 <- prospect_report_data %>% 
  rename(`MCVID (v29) (evar29)` = MCVID..v29...evar29.,
         `Tracking Code` = Tracking.Code,
         `Page Name (v26) (evar26)` = Page.Name..v26...evar26.,
         `URL (v8) (evar8)` = URL..v8...evar8.,
         `24H Clock by Minute` = Time.Parting..v18...evar18.,
         `Content Type (v7) (evar7)` = Content.Type..v7...evar7.,
         `Content Title (v22) (evar22)` = Content.Title..v22...evar22.,
         `Visit Number` = Visit.Number,
         `Exit Links` = Exit.Links,
         `Page Views` = Page.Views,
         `Unique Visitors` = Unique.Visitors,
         `Shop - Revenue` = Shop...Revenue,
         `Donate Revenue (Serialized) (ev114) (event114)` = Donate.Revenue..Serialized...ev114...event114.,
         `Membership Revenue (ev5) (event5)` = Membership.Revenue..ev5...event5.,
         `Holidays Booking Total Revenue (Serialised) (ev125) (event125)` = Holidays.Booking.Total.Revenue..Serialised...ev125...event125.,
         `Renew Revenue - Serialized (ev79) (event79)` = Renew.Revenue...Serialized..ev79...event79.,
         `Video Start (ev10) (event10)` = Video.Start..ev10...event10.) %>% 
  mutate(
    # Step A: Isolate the 12H Time and AM/PM part
    Time_12H = str_extract(`24H Clock by Minute`, "^[^|]+"),
    `24H Clock by Minute` = format(as.POSIXct(Time_12H,format='%I:%M %p'),format="%H:%M")
  ) %>% 
  select(-Time_12H)

write.csv(prospect_report_data2,prospects_excel_filename, row.names = FALSE)
