library(readr)
library(tidyverse)
library(lubridate)
library(openxlsx)
library(visNetwork)
library(tidyr) # For the pivot_wider function

project_name <-           "Prospect Analytics"
sub_title <-              "September 2025"
communication_type_set <- "Prospect Newsletter"
send_date_set <-          "2025-08-21"
#customer_status_set <-    "Net New"
customer_status_set <-   "Known"
#mosaic_cohort_set <-      "Cohort 1"
#mosaic_cohort_set <-     "Cohort 2.1" #A
#mosaic_cohort_set <-     "Cohort 2.2" #G & H
#mosaic_cohort_set <-     "Cohort 3" #B, C & E
mosaic_cohort_set <-     "Everyone Else"
mosaic_letter_set <-      "N" # Set as required.

project_data <- data.frame(
  Project_Name = project_name,
  Sub_Title = sub_title,
  Communication_Type = communication_type_set,
  Send_Date = send_date_set
  # Customer_Status = customer_status_set,
  # Mosaic_Cohort = mosaic_cohort_set,
  # Mosaic_letter = mosaic_letter_set
)

title <- paste("Journey:",communication_type_set, "-",mosaic_cohort_set, "-","Mosaic:",mosaic_letter_set, "-",customer_status_set, "-",send_date_set)

# Import File from Data Warehouse
Prospect_Email_All_Groups <- read_csv("C:/Users/alan.millington/OneDrive - National Trust/Analytics/R Projects/MOSAIC-Analysis/Prospect_Email_-_All_Groups/Prospect_Email_-_All_Groups.csv", 
                                      col_types = cols(`MCVID (v29) (evar29)` = col_character(), 
                                                       `Tracking Code` = col_character(), 
                                                       `Page Name (v26) (evar26)` = col_character(), 
                                                       `URL (v8) (evar8)` = col_character(), 
                                                       `24H Clock by Minute` = col_time(format = "%H:%M"), 
                                                       `Content Type (v7) (evar7)` = col_character(), 
                                                       `Content Title (v22) (evar22)` = col_character(), 
                                                       `Visit Number` = col_integer(), `Exit Links` = col_character()))

# Rename All columns
Prospect_Email_All_Groups <- Prospect_Email_All_Groups %>% 
  rename(date = Date,
         cust_id = `MCVID (v29) (evar29)`,
         tracking_code = `Tracking Code`,
         page_name = `Page Name (v26) (evar26)`,
         page_url = `URL (v8) (evar8)`,
         session_time = `24H Clock by Minute`,
         content_type = `Content Type (v7) (evar7)`,
         content_title = `Content Title (v22) (evar22)`,
         visit_num = `Visit Number`,
         exit_link = `Exit Links`,
         page_views = `Page Views`,
         unique_visitors = `Unique Visitors`,
         visits = Visits,
         shop_revenue = `Shop - Revenue`,
         donate_revenue = `Donate Revenue (Serialized) (ev114) (event114)`,
         membership_revenue = `Membership Revenue (ev5) (event5)`,
         holiday_revenue = `Holidays Booking Total Revenue (Serialised) (ev125) (event125)`,
         renew_revenue = `Renew Revenue - Serialized (ev79) (event79)`,
         video_start = `Video Start (ev10) (event10)`
  )
         
# Process & Clean Input Data      
# Copy data for back up: no changes other than columns renamed.

# Prospect_Email_All_Groups # Original Data frame
prospect_data <- Prospect_Email_All_Groups

# change the input date format to a standard R object date format.
prospect_data$date <- as.Date(prospect_data$date, format = "%B %d, %Y")

# Create timestamp column
prospect_data$timestamp <- with(prospect_data, as.POSIXct(paste(prospect_data$date, prospect_data$session_time), format="%Y-%m-%d %H:%M:%S"))

# Remove NA's from customer ID column
prospect_data <- prospect_data %>% drop_na(cust_id)

# Remove % in tracking code column - (test data)
prospect_data <- prospect_data %>% filter(!str_detect(tracking_code, "%")) 

# Copy the tracking Code column for future preservation
prospect_data <- prospect_data %>% mutate(tracking_code_orig = tracking_code)

# Separate the tracking code column into main pieces of the tracking code.
prospect_data <- prospect_data %>% separate(tracking_code, into = c("channel", "region", "campaign", "send_date", "cohort"), sep = "_")

# verify the cohort field contains 'NT'
prospect_data <- prospect_data %>% filter(str_detect(cohort, "NT"))

# split the cohort column to remove the region clicked from the email into a separate column
prospect_data <- prospect_data %>% separate(cohort, into = c("cohort", "em_click_region"), sep = "-")

# Change the send date column into a date object.
prospect_data <- prospect_data %>% mutate(send_date = as.Date(send_date, format = "%d%m%Y"))

prospect_data <- prospect_data %>%
  mutate(
    # Extract fixed-width sub strings from the 'cohort' string
    source_code = substr(cohort, 3, 8),
    date_str = substr(cohort, 9, 14),
    comm_type_code = substr(cohort, 15, 16),
    cust_status_code = substr(cohort, 17, 17),
    mosaic_letter = substr(cohort, 18, 18),
    
    # Convert the date string to a standard R date object
    date = dmy(date_str),
    
    # Decode the communication type
    communication_type = case_when(
      comm_type_code == "PW" ~ "Prospect Welcome",
      comm_type_code == "PN" ~ "Prospect Newsletter",
      comm_type_code == "DO" ~ "Double Opt-In",
      TRUE ~ NA_character_ # Populate with NA if code is not PW or PN
    ),
    
    # Decode the customer status
    customer_status = case_when(
      cust_status_code == "1" ~ "Known",
      cust_status_code == "2" ~ "Net New",
      TRUE ~ NA_character_ # Populate with NA if code is not 1 or 2
    ),
    
    # Classify the mosaic letter into cohorts
    mosaic_cohort = case_when(
      mosaic_letter %in% c("K", "O") ~ "Cohort 1",
      mosaic_letter %in% c("A") ~ "Cohort 2.1",
      mosaic_letter %in% c("G", "H") ~ "Cohort 2.2",
      mosaic_letter %in% c("B", "C", "E") ~ "Cohort 3",
      mosaic_letter %in% c("D", "F", "I", "J", "L", "M", "N", "U") ~ "Everyone Else",
      TRUE ~ NA_character_ # Populate with NA for any other letter
    )
  )

prospect_data <- prospect_data %>%
  filter(substr(source_code, 1, 3) %in% c('DCP', 'MCC', 'DCN'))

page_type_lookup <- prospect_data %>% 
  select(content_title, content_type) %>% 
  distinct(content_title, content_type)

# Check for duplicates
page_type_lookup %>% 
  group_by(content_title) %>% 
  filter(n()>1)
  
#View(prospect_data)
# Cleaning Complete

# Subset the Specific cohorts to usable group level.

prospect_data_filtered <- prospect_data %>% filter(
                                                communication_type == communication_type_set & 
                                                send_date == send_date_set &
                                                customer_status == customer_status_set &
                                                mosaic_cohort == mosaic_cohort_set &
                                                mosaic_letter == mosaic_letter_set
                                                  )

# prospect_data_filtered <- prospect_data %>% filter(
#     communication_type == "Prospect Newsletter" & 
#     send_date ==          "2025-08-21" &
#     customer_status ==    "Net New" &
#     mosaic_cohort ==      "Cohort 1" &
#     mosaic_letter ==      "O"
# )

search_term <- "Heritage Open Days"

# 3. Process the data to create the list of journeys
journeys <- prospect_data_filtered %>%
  # filter(if_any(
  #   .cols = where(is.character),
  #   .fns = ~ str_detect(., pattern = regex(search_term, ignore_case = TRUE))
  # )) %>% 
  # Group by each unique journey (customer and visit number)
  #group_by(cust_id, visit_num) %>% # Use this to see the in-session activity
  group_by(cust_id) %>%             # Use this to see entire user activity across multiple sessions.
  # Order the pages within each journey chronologically
  arrange(timestamp) %>%
  # Combine the content_title for each group into a list-column
  summarise(
    journey_path = list(content_title[!duplicated(content_title)]),
    .groups = 'drop'
  ) %>% pull(journey_path)
  # Pull the list-column out into a final list variable

# Skip - Content Type
# Content Type v7
# journeys <- prospect_data_filtered %>%
#   # Group by each unique journey (customer and visit number)
#   group_by(cust_id, visit_num) %>%
#   # Order the pages within each journey chronologically
#   arrange(timestamp) %>%
#   # Combine the content_title for each group into a list-column
#   summarise(
#     journey_path = list(content_type[!duplicated(content_type)]),
#     .groups = 'drop'
#   ) %>% pull(journey_path)

# 3. Create a data frame of all unique transitions
transitions_df <- map_dfr(journeys, ~tibble(
  from = .x[-length(.x)],
  to   = .x[-1]
)) %>% distinct()
  # Keep only the unique sequence of pages


transitions_df <- map_dfr(journeys, ~tibble(
  from = .x[-length(.x)],
  to   = .x[-1]
)) %>% distinct()


# Create the nodes data frame
# This must have a unique 'id' column
nodes <- tibble(
  id = unique(c(transitions_df$from, transitions_df$to)),
  label = id, # Use the page name as the label
  shape = "dot",
  group = "",
  size = 15,
  title = id # Tooltip on hover
)

# Check for duplicates
duplicates <- nodes %>% 
  group_by(id) %>% 
  filter(n()>1)

# Create the edges data frame
# This must have 'from' and 'to' columns
edges <- transitions_df %>%
  mutate(
    arrows = "to",
    title = paste("From:", from, "<br>To:", to) # Tooltip on hover
  )

# 5. Create the interactive graph
# This will render in the RStudio 'Viewer' pane or a web browser
visNetwork(nodes, edges, main = title, width = "100%", height = "100vh") %>%
  # Use a layout that spreads nodes out
  visIgraphLayout(layout = "layout_nicely") %>%
  visNodes(
    shape = "dot",
    color = list(
      background = "#0085AF",
      border = "#013848",
      highlight = "#FF8000"
    ),
    shadow = list(enabled = TRUE, size = 10)
  ) %>%
  visEdges(
    shadow = FALSE,
    width = 2,
    color = list(color = "#0085AF", highlight = "#C62F4B")
  ) %>%
  # Add options for interactivity
  visOptions(
    highlightNearest = list(enabled = TRUE, degree = 2, hover = TRUE),
    selectedBy = "group",
    nodesIdSelection = TRUE # Adds a dropdown to select a page
  ) %>%
  # Improve physics for a more stable layout
  visPhysics(solver = "forceAtlas2Based")

#-----------------------------------------------------------------------
# Build Reporting Tables
#-----------------------------------------------------------------------
# 
# ## "--- Engagement Metrics Comparison by Cohort ---"
# engagement_comparison <- prospect_data %>%
#   group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>%
#   summarise(journey_length = n(), .groups = 'drop') %>%
#   group_by(communication_type,mosaic_cohort) %>%
#   summarise(
#     total_journeys = n(),
#     avg_journey_length = round(mean(journey_length), digits = 1)
#   ) %>% 
#   arrange(desc(communication_type))
# 
# #View(engagement_comparison)
# 
# ## "--- Engagement Metrics Comparison using Mosaic Letter---"
# engagement_comparison_mosaic <- prospect_data %>%
#   # First, find the length of each unique journey
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   summarise(journey_length = n(), .groups = 'drop') %>%
#   group_by(communication_type,mosaic_letter) %>%
#   summarise(
#     total_journeys = n(),
#     avg_journey_length = round(mean(journey_length),digits = 1)
#   ) %>% 
#   arrange(desc(communication_type))
# 
# # print("--- Engagement Metrics Comparison using Mosaic Letter---")
# # View(engagement_comparison_mosaic)
# 
# ## Table 2: Top 3 Most Visited Pages per Cohort
# top_pages_comparison <- prospect_data %>%
#   group_by(communication_type, mosaic_cohort, content_title) %>%
#   summarise(page_views = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_cohort) %>%
#   # slice_max gets the top N rows based on a variable
#   slice_max(order_by = page_views, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# # print("--- Top 3 Pages Comparison by Cohort Group---")
# # View(top_pages_comparison)
# 
# ## Table 2.1: Top 3 Most Visited Pages by Mosaic Letter
# top_pages_comparison_mosiac <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, content_title) %>%
#   summarise(page_views = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   # slice_max gets the top N rows based on a variable
#   slice_max(order_by = page_views, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# # print("--- Top 3 Pages Comparison by Mosaic Letter ---")
#  #View(top_pages_comparison_mosiac)
# 
# ## Table 3: Top 3 Most Common Transitions (A -> B) per Cohort
# top_paths_comparison <- prospect_data %>%
#   group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   ungroup() %>%
#   # Remove rows where there is no next page (end of a journey)
#   filter(!is.na(next_page)) %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_cohort, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_cohort) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type)) 
# 
# # print("--- Top 3 Paths Comparison ---")
# # View(top_paths_comparison)
# 
#  #4. Calculate the transition probabilities
#   # Calculation is not correct #fix, need to calculated based on the cohort group?
#  transition_probs <- prospect_data %>%
#    group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>%
#    arrange(timestamp, .by_group = TRUE) %>%
#    # Create 'from' and 'to' columns for each transition
#    mutate(next_page = lead(content_title)) %>%
#    ungroup() %>%
#    # Remove rows where there is no next page (end of a journey)
#    filter(!is.na(next_page)) %>%
#    # Count the transitions
#    group_by(communication_type, mosaic_cohort, from = content_title, to = next_page) %>%
#    summarise(transition_count = n(), .groups = 'drop') %>%
#    group_by(communication_type, mosaic_cohort) %>%
#    slice_max(order_by = transition_count, n = 10) %>% 
#    arrange(desc(communication_type)) %>% 
#    mutate(probability = transition_count / sum(transition_count)) %>%    # Calculate probability
#    ungroup() %>% 
#    arrange(desc(transition_count))
#  
# ## Table 3.1: Top 3 Most Common Transitions (A -> B) per Cohort
# top_paths_comparison_mosaic <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   ungroup() %>%
#   # Remove rows where there is no next page (end of a journey)
#   filter(!is.na(next_page)) %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# #print("--- Top 3 Paths Comparison ---")
# #View(top_paths_comparison_mosaic)
# 
# ## Membership
# ## Table 4: Top Paths to Membership Revenue Funnel
# ## Users Enter the membership funnel but may not complete it.
# top_paths_mem_funnel <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#  # filter(any(membership_revenue > 0| donate_revenue > 0 | holiday_revenue > 0 | shop_revenue > 0)) %>% # + Renew
#   filter(any(page_name == 'M|Membership|Step 1.0 - Your Details')) %>%  
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# ## Membership
# ## Table 5: Top Paths to Membership Revenue Funnel - Conversion
# ## Users Enter the membership funnel AND revenue exists.
# top_paths_mem_funnel_revenue <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(membership_revenue > 0)) %>% # + Renew
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# # ----------------------------------------------------------------------------
# ## Membership Renewals
# ## Table 6: Top Paths to Membership Revenue Funnel
# ## Users Enter the membership renewal funnel but may not complete it.
# top_paths_renew_funnel <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   # filter(any(membership_revenue > 0| donate_revenue > 0 | holiday_revenue > 0 | shop_revenue > 0)) %>% # + Renew
#   filter(any(content_type == 'Renew')) %>% 
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# ## Membership Renewals
# ## Table 7: Top Paths to Membership Renewals Funnel - Conversion
# ## Users Enter the membership renewal funnel AND revenue exists.
# top_paths_renew_funnel_revenue <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(renew_revenue > 0)) %>% # + Renew
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# # ----------------------------------------------------------------------------
# ## Donate
# ## Table 8: Top Paths to Donation Revenue Funnel
# ## Users Enter the Donation Funnel but may not complete it.
# top_paths_donate_funnel <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(content_type == 'Donate')) %>% 
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# ## Donate
# ## Table 9: Top Paths to Donations Funnel - Conversion
# ## Users Enter the donation funnel AND revenue exists.
# top_paths_donate_funnel_revenue <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(donate_revenue > 0)) %>% # + Renew
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# # ----------------------------------------------------------------------------
# ## Shop
# ## Table 10: Top Paths to the NT Shop Revenue Funnel
# ## Users Enter the and view Shop pages but may not make a purchase.
# top_paths_shop_funnel <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(content_type == 'ShopHome')) %>% 
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# ## Shop
# ## Table 11: Top Paths to Shop Funnel - Conversion
# ## Users Enter the Shop funnel AND revenue exists.
# top_paths_shop_funnel_revenue <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(shop_revenue > 0)) %>% # + Renew
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# # ----------------------------------------------------------------------------
# ## Holidays
# ## Table 12: Top Paths to the NT Holidays Revenue Funnel
# ## Users Enter the and view Holidays pages but may not make a purchase.
# top_paths_holidays_funnel <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(str_detect(content_type, "Holidays"))) %>% 
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# ## Holidays
# ## Table 13: Top Paths to Holidays Funnel - Conversion
# ## Users Enter the Holidays funnel AND revenue exists.
# top_paths_holidays_funnel_revenue <- prospect_data %>%
#   group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
#   arrange(timestamp, .by_group = TRUE) %>%
#   # Create 'from' and 'to' columns for each transition
#   mutate(next_page = lead(content_title)) %>%
#   filter(any(holiday_revenue > 0)) %>% # + Renew
#   ungroup() %>%
#   # Count the transitions
#   group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
#   summarise(transition_count = n(), .groups = 'drop') %>%
#   group_by(communication_type, mosaic_letter) %>%
#   slice_max(order_by = transition_count, n = 10) %>% 
#   arrange(desc(communication_type))
# 
# #-----------------------------------------------------------------------
# # Export data to Excel
# #-----------------------------------------------------------------------
# 
# if (!require(openxlsx)) {
#   install.packages("openxlsx")
#   library(openxlsx)
# }
# 
# excel_filename <- paste0(project_name,"_Prospect_Analytics_",format(Sys.time(), "%d%m%Y_%H%M%S"), ".xlsx")
# current_project_sheet_name <- project_name
# 
# list_of_datasets <- list(
#   "Project Info" = if (exists("project_data")) project_data else NULL,
#   "Engagement Comparision By Cohort" = if (exists("engagement_comparison")) engagement_comparison else NULL,
#   "Engagement Comparision By Mosaic Letter" = if (exists("engagement_comparison_mosaic")) engagement_comparison_mosaic else NULL,
#   "Top Pages Comparision By Cohort" = if (exists("top_pages_comparison")) top_pages_comparison else NULL,
#   "Top Pages Comparision By Mosaic Letter" = if (exists("top_pages_comparison_mosiac")) top_pages_comparison_mosiac else NULL,
#   "Top Paths Comparison By Cohort" = if (exists("top_paths_comparison")) top_paths_comparison else NULL,
#   "Top Paths Comparison By Mosiac Letter" = if (exists("top_paths_comparison_mosaic")) top_paths_comparison_mosaic else NULL,
#   "Transition Probabilites" = if (exists("transition_probs")) transition_probs else NULL,
#   "Journey to Membership Funnel" = if (exists("top_paths_mem_funnel")) top_paths_mem_funnel else NULL,
#   "Journey to Membership Conversion" = if (exists("top_paths_mem_funnel_revenue")) top_paths_mem_funnel_revenue else NULL,
#   "Journey to Membership Renewal Funnel" = if (exists("top_paths_renew_funnel")) top_paths_renew_funnel else NULL,
#   "Journey to Membership Renewal Conversion" = if (exists("top_paths_renew_funnel_revenue")) top_paths_renew_funnel_revenue else NULL,
#   "Journey to Donations Funnel" = if (exists("top_paths_donate_funnel")) top_paths_donate_funnel else NULL,
#   "Journey to Donations Conversion" = if (exists("top_paths_donate_funnel_revenue")) top_paths_donate_funnel_revenue else NULL,
#   "Journey to Shop Pages" = if (exists("top_paths_shop_funnel")) top_paths_shop_funnel else NULL,
#   "Journey to Shop Conversion" = if (exists("top_paths_shop_funnel_revenue")) top_paths_shop_funnel_revenue else NULL,
#   "Journey to Holidays Pages" = if (exists("top_paths_holidays_funnel")) top_paths_holidays_funnel else NULL,
#   "Journey to Holiday Conversion" = if (exists("top_paths_holidays_funnel_revenue")) top_paths_holidays_funnel_revenue else NULL,
#   "All Prospect Data" = if (exists("prospect_data")) prospect_data else NULL
# )
# 
# # Filter out any NULL elements (if a df didn't exist)
# list_of_datasets <- list_of_datasets[!sapply(list_of_datasets, is.null)]
# 
# # Create or load workbook
# if (file.exists(excel_filename)) {
#   wb <- loadWorkbook(excel_filename)
# } else {
#   wb <- createWorkbook()
# }
# 
# # Remove existing sheet for the current project_name if it exists, then add a new one
# if (current_project_sheet_name %in% names(wb)) {
#   removeWorksheet(wb, current_project_sheet_name)
# }
# addWorksheet(wb, current_project_sheet_name)
# 
# # Write data to the sheet
# current_row <- 1
# # Create a style for subsection headers
# header_style <- createStyle(textDecoration = "bold", fontSize = 14, fgFill = "#DDEBF7") # Light blue fill
# 
# if (length(list_of_datasets) > 0) {
#   for (i in 1:length(list_of_datasets)) {
#     subsection_title <- names(list_of_datasets)[i]
#     df_to_write <- list_of_datasets[[i]]
#     
#     # Write subsection title
#     writeData(wb, sheet = current_project_sheet_name, x = subsection_title, startCol = 1, startRow = current_row)
#     addStyle(wb, sheet = current_project_sheet_name, style = header_style, rows = current_row, cols = 1:5) # Apply style to a few cells
#     mergeCells(wb, sheet = current_project_sheet_name, rows = current_row, cols = 1:9) # Merge cells for title
#     current_row <- current_row + 1
#     
#     # Write data frame
#     if (!is.null(df_to_write) && nrow(df_to_write) > 0) {
#       writeData(wb, sheet = current_project_sheet_name, 
#                 x = df_to_write, startCol = 1, 
#                 startRow = current_row,
#                 headerStyle = createStyle(textDecoration = "bold", 
#                                           border = "TopBottomLeftRight"))
#       current_row <- current_row + nrow(df_to_write) + 2 # Add 2 for header and a blank row
#     } else {
#       writeData(wb, sheet = current_project_sheet_name, x = "Data not available or empty.", startCol = 1, startRow = current_row)
#       current_row <- current_row + 3 # Add 3 for message and a blank row
#     }
#   }
# } else {
#   writeData(wb, sheet = current_project_sheet_name, x = "No dataframes were specified or found for export.", startCol = 1, startRow = current_row)
#   addStyle(wb, sheet = current_project_sheet_name, style = header_style, rows = current_row, cols = 1:5)
# }
# # Save the workbook
# tryCatch({
#   saveWorkbook(wb, excel_filename, overwrite = TRUE)
#   cat(paste0("Report data for '", current_project_sheet_name, "' successfully written to '", excel_filename, "'\n"))
# }, error = function(e) {
#   cat(paste0("Error saving Excel file: ", e$message, "\n"))
#   cat("Please ensure the file is not open elsewhere and you have write permissions.\n")
# })

# library(inspectdf)
# inspect_cat(prospect_data) %>% show_plot() # Frequency Report (Useful)

# library(smartEDA)
# ExpReport(
#   prospect_data,
#   op_dir = "report/",
#   op_file = "MOSAIC-Report.html"
# )