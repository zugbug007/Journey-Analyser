# 1. LOAD LIBRARIES ---------------------------------------------------------
# Consolidated list of all required libraries from both apps
library(shiny)
library(tidyverse)
library(lubridate)
library(openxlsx)
library(visNetwork)
library(DT)
library(stringr)
library(purrr)
library(shinycssloaders) # For loading spinners
library(bslib)           # For modern themes

# Increase the maximum upload size to 30MB
options(shiny.maxRequestSize = 30 * 1024^2)


# 2. HELPER FUNCTIONS -------------------------------------------------------

# --- Helper function for Member Engagement ---
parse_tracking_string <- function(x) {
  output_structure <- tibble(
    cohort = NA_character_,
    membership_type = NA_character_,
    em_click_region = NA_character_,
    em_click_detail = NA_character_,
    test_variant = NA_character_
  )
  if (is.na(x) || !is.character(x) || nchar(x) == 0) return(output_structure)
  
  membership_lookup <- c("SE" = "Senior", "SP" = "Senior Plus", "TF" = "Teen Family", "YI" = "Young Independent", "YF" = "Young Family", "MI" = "Mature Independent", "SD" = "Unknown", "NG" = "Unknown")
  cohort_pattern    <- "^ME(SE|SP|TF|YI|YF|MI|SD|NG)$"
  variant_pattern   <- "^[A-D]$"
  region_pattern    <- "^(Mod[0-9]*|Nav|Footer)$"
  detail_pattern    <- "^(COL[0-9]*|CTA[0-9]*|IMAGE[0-9]*)$"
  
  parts <- str_split(x, "-")[[1]]
  
  if (str_detect(parts[1], cohort_pattern)) {
    output_structure$cohort <- parts[1]
    mem_code <- str_sub(parts[1], 3, 4)
    output_structure$membership_type <- membership_lookup[mem_code]
    parts <- parts[-1] 
  } else {
    return(output_structure)
  }
  
  if (length(parts) > 0 && str_detect(last(parts), variant_pattern)) {
    output_structure$test_variant <- last(parts)
    parts <- head(parts, -1)
  }
  
  if (length(parts) > 0) {
    region_matches <- str_subset(parts, region_pattern)
    detail_matches <- str_subset(parts, detail_pattern)
    if (length(region_matches) > 0) output_structure$em_click_region <- paste(region_matches, collapse = ", ")
    if (length(detail_matches) > 0) output_structure$em_click_detail <- paste(detail_matches, collapse = ", ")
  }
  
  return(output_structure)
}

# --- Helper function for Prospects ---
parse_prospect_cohort <- function(x) {
  # Define the structure to be returned
  output_structure <- tibble(
    cohort_clean = NA_character_,
    em_click_region = NA_character_,
    em_click_detail = NA_character_
  )
  if (is.na(x) || !is.character(x) || nchar(x) == 0) return(output_structure)
  
  # Split the cohort string by hyphen
  parts <- str_split(x, "-")[[1]]
  
  # The first part is always the clean cohort ID
  output_structure$cohort_clean <- parts[1]
  
  # Process the remaining parts if they exist
  if (length(parts) > 1) {
    remaining_parts <- parts[-1]
    
    # Define the patterns to look for
    region_pattern <- "^(Mod[0-9]*|Nav|Footer)$"
    detail_pattern <- "^(COL[0-9]*|CTA[0-9]*|IMAGE[0-9]*)$"
    
    # Find matches for region and detail
    region_matches <- str_subset(remaining_parts, region_pattern)
    detail_matches <- str_subset(remaining_parts, detail_pattern)
    
    # Populate the columns, collapsing multiple matches if necessary
    if (length(region_matches) > 0) {
      output_structure$em_click_region <- paste(region_matches, collapse = ", ")
    }
    if (length(detail_matches) > 0) {
      output_structure$em_click_detail <- paste(detail_matches, collapse = ", ")
    }
  }
  
  return(output_structure)
}


# 3. USER INTERFACE (UI) ----------------------------------------------------
ui <- fluidPage(
  
  # Set the browser window title here
  title = "Journey Analyser",
  theme = bs_theme(version = 5, bootswatch = "cosmo"),
  
  # Use a standard header tag for the dynamic on-page title
  h2(textOutput("app_title")),
  
  sidebarLayout(
    
    sidebarPanel(
      
      h4("Step 1: Select Analysis Type"),
      selectInput("analysis_type", "Analysis Type", 
                  choices = c("Member Engagement", "Prospects"), 
                  selected = "Member Engagement"),
      
      tags$hr(),
      
      h4("Step 2: Upload Data"),
      fileInput("file_upload", "Upload Journey CSV File",
                accept = c("text/csv", ".csv")),
      
      tags$hr(),
      
      # This entire section of filters will be rendered dynamically from the server
      uiOutput("dynamic_sidebar_filters")
      
    ),
    
    mainPanel(
      # The main panel content (tabs) will also be rendered dynamically
      uiOutput("dynamic_main_panel")
    )
  )
)


# 4. SERVER LOGIC -----------------------------------------------------------
server <- function(input, output, session) {
  
  # A. DYNAMIC UI ELEMENTS --------------------------------------------------
  
  # This now renders the text for the h2 tag in the UI
  output$app_title <- renderText({
    paste(input$analysis_type, "Journey Analyser")
  })
  
  # This reactive value will hold the state of the file upload
  output$file_uploaded <- reactive({
    !is.null(input$file_upload)
  })
  outputOptions(output, 'file_uploaded', suspendWhenHidden = FALSE)
  
  
  # B. DATA READING AND VALIDATION ------------------------------------------
  
  # This reactive reads the raw CSV and validates its structure against the selected analysis type
  validated_data <- reactive({
    req(input$file_upload)
    
    # Read the raw data first
    df <- read_csv(input$file_upload$datapath, col_types = cols(.default = "c"))
    
    # --- Validation Logic ---
    sample_codes <- df$`Tracking Code` %>% na.omit() %>% head(100)
    
    fifth_elements <- map_chr(str_split(sample_codes, "_"), function(.x) {
      if (length(.x) >= 5) .x[[5]] else NA_character_
    }) %>% na.omit()
    
    is_me_data <- any(str_starts(fifth_elements, "ME"))
    is_prospect_data <- any(str_detect(fifth_elements, "NT"))
    
    # Check for mismatches
    if (input$analysis_type == "Member Engagement" && !is_me_data && is_prospect_data) {
      validate(need(FALSE, "File mismatch. You selected 'Member Engagement', but the uploaded file appears to be for 'Prospects' analysis."))
    }
    
    if (input$analysis_type == "Prospects" && !is_prospect_data && is_me_data) {
      validate(need(FALSE, "File mismatch. You selected 'Prospects', but the uploaded file appears to be for 'Member Engagement' analysis."))
    }
    
    # If validation passes, return the raw dataframe for processing
    return(df)
  })
  
  # This reactive takes the validated data and applies the correct cleaning pipeline
  base_data <- reactive({
    df <- validated_data() # This depends on the validated data
    
    # --- CONDITIONAL DATA PROCESSING ---
    if (input$analysis_type == "Member Engagement") {
      # === MEMBER ENGAGEMENT PIPELINE ===
      member_engagement_data <- df %>%
        rename(
          date = Date, cust_id = `MCVID (v29) (evar29)`, tracking_code = `Tracking Code`,
          page_name = `Page Name (v26) (evar26)`, page_url = `URL (v8) (evar8)`, session_time = `24H Clock by Minute`,
          content_type = `Content Type (v7) (evar7)`, content_title = `Content Title (v22) (evar22)`,
          visit_num = `Visit Number`, exit_link = `Exit Links`, page_views = `Page Views`,
          unique_visitors = `Unique Visitors`, visits = Visits, shop_revenue = `Shop - Revenue`,
          donate_revenue = `Donate Revenue (Serialized) (ev114) (event114)`, membership_revenue = `Membership Revenue (ev5) (event5)`,
          holiday_revenue = `Holidays Booking Total Revenue (Serialised) (ev125) (event125)`,
          renew_revenue = `Renew Revenue - Serialized (ev79) (event79)`, video_start = `Video Start (ev10) (event10)`
        ) %>%
        mutate(
          across(c(visits, shop_revenue, donate_revenue, membership_revenue, holiday_revenue, renew_revenue, video_start), as.numeric),
          date = mdy(date),
          session_time = hms(session_time),
          timestamp = as.POSIXct(paste(date, session_time), format = "%Y-%m-%d %H:%M:%S")
        ) %>%
        drop_na(cust_id) %>%
        filter(!str_detect(tracking_code, "%")) %>%
        mutate(tracking_code_orig = tracking_code) %>%
        separate(tracking_code, into = c("channel", "region", "campaign", "send_date_str", "cohort_group"), sep = "_", remove = FALSE) %>%
        mutate(send_date = dmy(send_date_str))
      
      parsed_data <- map_dfr(member_engagement_data$cohort_group, parse_tracking_string)
      
      final_data <- bind_cols(member_engagement_data, parsed_data) %>%
        drop_na(campaign, send_date, membership_type, test_variant)
      
      return(final_data)
      
    } else {
      # === PROSPECTS PIPELINE (UPDATED) ===
      
      # First, do the standard renaming and type conversion
      prospect_data_intermediate <- df %>%
        rename(
          date = Date, cust_id = `MCVID (v29) (evar29)`, tracking_code = `Tracking Code`,
          page_name = `Page Name (v26) (evar26)`, page_url = `URL (v8) (evar8)`, session_time = `24H Clock by Minute`,
          content_type = `Content Type (v7) (evar7)`, content_title = `Content Title (v22) (evar22)`,
          visit_num = `Visit Number`, exit_link = `Exit Links`, page_views = `Page Views`,
          unique_visitors = `Unique Visitors`, visits = Visits, shop_revenue = `Shop - Revenue`,
          donate_revenue = `Donate Revenue (Serialized) (ev114) (event114)`, membership_revenue = `Membership Revenue (ev5) (event5)`,
          holiday_revenue = `Holidays Booking Total Revenue (Serialised) (ev125) (event125)`,
          renew_revenue = `Renew Revenue - Serialized (ev79) (event79)`, video_start = `Video Start (ev10) (event10)`
        ) %>%
        mutate(
          across(c(visits, shop_revenue, donate_revenue, membership_revenue, holiday_revenue, renew_revenue, video_start), as.numeric),
          date = mdy(date),
          session_time = hms(session_time),
          timestamp = as.POSIXct(paste(date, session_time), format = "%Y-%m-%d %H:%M:%S")
        ) %>%
        drop_na(cust_id) %>%
        filter(!str_detect(tracking_code, "%")) %>%
        mutate(tracking_code_orig = tracking_code) %>%
        separate(tracking_code, into = c("channel", "region", "campaign", "send_date_str", "cohort_full"), sep = "_", remove = FALSE) %>%
        filter(str_detect(cohort_full, "NT"))
      
      # Apply the new parsing function to the complex cohort field
      parsed_cohort_data <- map_dfr(prospect_data_intermediate$cohort_full, parse_prospect_cohort)
      
      # Bind the parsed data and finalize the cleaning process
      prospect_data <- bind_cols(prospect_data_intermediate, parsed_cohort_data) %>%
        select(-cohort_full) %>% # remove the original combined field
        rename(cohort = cohort_clean) %>% # overwrite cohort with the clean version
        mutate(send_date = dmy(send_date_str)) %>%
        mutate(
          source_code = substr(cohort, 3, 8), date_str = substr(cohort, 9, 14),
          comm_type_code = substr(cohort, 15, 16), cust_status_code = substr(cohort, 17, 17),
          mosaic_letter = substr(cohort, 18, 18),
          communication_type = case_when(comm_type_code == "PW" ~ "Prospect Welcome", comm_type_code == "PN" ~ "Prospect Newsletter", comm_type_code == "DO" ~ "Double Opt-In", TRUE ~ NA_character_),
          customer_status = case_when(cust_status_code == "1" ~ "Known", cust_status_code == "2" ~ "Net New", TRUE ~ NA_character_),
          mosaic_cohort = case_when(mosaic_letter %in% c("K", "O") ~ "Cohort 1", mosaic_letter %in% c("A") ~ "Cohort 2.1", mosaic_letter %in% c("G", "H") ~ "Cohort 2.2", mosaic_letter %in% c("B", "C", "E") ~ "Cohort 3", mosaic_letter %in% c("D", "F", "I", "J", "L", "M", "N", "U") ~ "Everyone Else", TRUE ~ NA_character_)
        ) %>%
        filter(substr(source_code, 1, 3) %in% c('DCP', 'MCC', 'DCN')) %>%
        drop_na(communication_type, customer_status, mosaic_cohort, mosaic_letter)
      
      return(prospect_data)
    }
  })
  
  
  # C. DYNAMIC UI RENDERING -------------------------------------------------
  
  # Renders the correct set of sidebar filters based on analysis type
  output$dynamic_sidebar_filters <- renderUI({
    req(base_data()) # Require data to be loaded
    
    if (input$analysis_type == "Member Engagement") {
      # --- ME Filters ---
      tagList(
        h4("Step 3: Filter Journey"),
        p("Select a cohort to visualize their journey graph."),
        selectInput("campaign_set", "Campaign", choices = sort(unique(base_data()$campaign))),
        selectInput("send_date_set", "Send Date", choices = sort(unique(base_data()$send_date), decreasing = TRUE)),
        selectInput("membership_type_set", "Membership Type", choices = sort(unique(base_data()$membership_type))),
        selectInput("test_variant_set", "Test Variant", choices = c("All", unique(na.omit(base_data()$test_variant)))),
        selectInput("em_click_region_set", "Email Click Region", choices = c("All", unique(na.omit(base_data()$em_click_region)))),
        selectInput("em_click_detail_set", "Email Click Detail", choices = c("All", unique(na.omit(base_data()$em_click_detail)))),
        checkboxInput("use_content_type", "Group Nodes by Content Type", value = FALSE),
        tags$hr(),
        h4("Step 4: Download Full Report"),
        p("Generates an Excel file with all summary tables for the uploaded data."),
        downloadButton("download_report", "Download Excel Report", class = "btn-success")
      )
    } else {
      # --- Prospects Filters ---
      tagList(
        h4("Step 3: Filter Journey"),
        p("Select a cohort to visualize their journey graph."),
        selectInput("comm_type_set", "Communication Type", choices = unique(base_data()$communication_type)),
        selectInput("send_date_set", "Send Date", choices = sort(unique(base_data()$send_date), decreasing = TRUE)),
        selectInput("cust_status_set", "Customer Status", choices = unique(base_data()$customer_status)),
        selectInput("mosaic_cohort_set", "Mosaic Cohort", choices = c("All", unique(base_data()$mosaic_cohort))),
        # This mosaic letter dropdown is now dependent on the cohort selection
        uiOutput("mosaic_letter_ui_placeholder"),
        selectInput("em_click_region_set", "Email Click Region", choices = c("All", unique(na.omit(base_data()$em_click_region)))),
        selectInput("em_click_detail_set", "Email Click Detail", choices = c("All", unique(na.omit(base_data()$em_click_detail)))),
        checkboxInput("use_content_type", "Group Nodes by Content Type", value = FALSE),
        tags$hr(),
        h4("Step 4: Download Full Report"),
        p("Generates an Excel file with all summary tables for the uploaded data."),
        downloadButton("download_report", "Download Excel Report", class = "btn-success")
      )
    }
  })
  
  # Special dynamic UI for the mosaic letter (Prospects only)
  output$mosaic_letter_ui_placeholder <- renderUI({
    # Only show this dropdown if a specific cohort is selected
    req(input$analysis_type == "Prospects", input$mosaic_cohort_set != "All")
    
    choices <- base_data() %>%
      filter(mosaic_cohort == input$mosaic_cohort_set) %>%
      pull(mosaic_letter) %>%
      unique() %>% sort()
    # *** CHANGE: Add "All" to the Mosaic Letter choices ***
    selectInput("mosaic_letter_set", "Mosaic Letter", choices = c("All", choices))
  })
  
  
  # Renders the correct set of main panel tabs based on analysis type
  output$dynamic_main_panel <- renderUI({
    tagList(
      # This panel shows the main content AFTER a file is uploaded
      conditionalPanel(
        condition = "output.file_uploaded",
        {
          # --- Build a list of tab panels conditionally ---
          
          # Start with the common first tab
          tabs_list <- list(
            tabPanel(
              "Journey Graph", icon = icon("project-diagram"),
              h3(textOutput("graph_title")),
              p("This interactive graph shows user paths. Nodes are web pages, edges are transitions."),
              visNetworkOutput("network_plot", height = "85vh")
            )
          )
          
          # Add the second tab based on analysis type
          if (input$analysis_type == "Member Engagement") {
            tabs_list[[2]] <- tabPanel(
              "Engagement & Path Reports", icon = icon("table"),
              h3("Engagement Metrics"), DT::dataTableOutput("engagement_comparison_table"),
              tags$hr(),
              h3("Top 10 Most Visited Pages"), DT::dataTableOutput("top_pages_comparison_table"),
              tags$hr(),
              h3("Top 10 Most Common User Paths"), DT::dataTableOutput("top_paths_comparison_table")
            )
          } else {
            tabs_list[[2]] <- tabPanel(
              "Engagement Reports", icon = icon("chart-bar"),
              h3("Engagement Metrics"), DT::dataTableOutput("engagement_comparison_table"),
              tags$hr(), DT::dataTableOutput("engagement_comparison_mosaic_table"),
              tags$hr(),
              h3("Top 10 Most Visited Pages"), DT::dataTableOutput("top_pages_comparison_table"),
              tags$hr(), DT::dataTableOutput("top_pages_comparison_mosaic_table")
            )
          }
          
          # Add the third tab based on analysis type
          if (input$analysis_type == "Member Engagement") {
            tabs_list[[3]] <- tabPanel(
              "Conversion Funnel Analysis", icon = icon("filter"),
              h3("Conversion Funnel Paths"),
              p("Top paths within journeys that either entered a conversion funnel or resulted in a conversion."),
              h4("Membership"), DT::dataTableOutput("top_paths_mem_funnel_table"), DT::dataTableOutput("top_paths_mem_funnel_revenue_table"),
              tags$hr(), h4("Renewals"), DT::dataTableOutput("top_paths_renew_funnel_table"), DT::dataTableOutput("top_paths_renew_funnel_revenue_table"),
              tags$hr(), h4("Donations"), DT::dataTableOutput("top_paths_donate_funnel_table"), DT::dataTableOutput("top_paths_donate_funnel_revenue_table"),
              tags$hr(), h4("Shop"), DT::dataTableOutput("top_paths_shop_funnel_table"), DT::dataTableOutput("top_paths_shop_funnel_revenue_table"),
              tags$hr(), h4("Holidays"), DT::dataTableOutput("top_paths_holidays_funnel_table"), DT::dataTableOutput("top_paths_holidays_funnel_revenue_table")
            )
          } else {
            tabs_list[[3]] <- tabPanel(
              "Path & Funnel Analysis", icon = icon("filter"),
              h3("Top 10 Most Common User Paths"), DT::dataTableOutput("top_paths_comparison_table"),
              tags$hr(), DT::dataTableOutput("top_paths_comparison_mosaic_table"),
              tags$hr(),
              h3("Conversion Funnel Paths"),
              p("Top 10 paths within journeys that either entered a conversion funnel or resulted in a conversion."),
              h4("Membership"), DT::dataTableOutput("top_paths_mem_funnel_table"), DT::dataTableOutput("top_paths_mem_funnel_revenue_table"),
              tags$hr(), h4("Renewals"), DT::dataTableOutput("top_paths_renew_funnel_table"), DT::dataTableOutput("top_paths_renew_funnel_revenue_table"),
              tags$hr(), h4("Donations"), DT::dataTableOutput("top_paths_donate_funnel_table"), DT::dataTableOutput("top_paths_donate_funnel_revenue_table"),
              tags$hr(), h4("Shop"), DT::dataTableOutput("top_paths_shop_funnel_table"), DT::dataTableOutput("top_paths_shop_funnel_revenue_table"),
              tags$hr(), h4("Holidays"), DT::dataTableOutput("top_paths_holidays_funnel_table"), DT::dataTableOutput("top_paths_holidays_funnel_revenue_table")
            )
          }
          
          # Add arguments and call tabsetPanel
          tabs_list$type <- "tabs"
          tabs_list$id <- "main_tabs"
          
          # Wrap the result in the spinner
          withSpinner(do.call(tabsetPanel, tabs_list), type = 6, color = "#0085AF")
        }
      ),
      # This panel shows the welcome message BEFORE a file is uploaded
      conditionalPanel(
        condition = "!output.file_uploaded",
        h2("Please select an analysis type and upload a CSV file.", style = "text-align: center; color: #888; margin-top: 50px;")
      )
    )
  })
  
  
  # D. REACTIVE DATA FILTERING ----------------------------------------------
  
  filtered_data <- reactive({
    req(base_data()) # Base data must exist
    
    data <- base_data()
    
    if (input$analysis_type == "Member Engagement") {
      req(input$campaign_set, input$send_date_set, input$membership_type_set, input$test_variant_set, input$em_click_region_set, input$em_click_detail_set)
      data <- data %>% filter(campaign == input$campaign_set, send_date == input$send_date_set, membership_type == input$membership_type_set)
      if (input$test_variant_set != "All") data <- data %>% filter(test_variant == input$test_variant_set)
      if (input$em_click_region_set != "All") data <- data %>% filter(em_click_region == input$em_click_region_set)
      if (input$em_click_detail_set != "All") data <- data %>% filter(em_click_detail == input$em_click_detail_set)
      
    } else {
      # For Prospects, mosaic_letter_set might not exist if "All" cohorts are selected
      req(input$comm_type_set, input$send_date_set, input$cust_status_set, input$mosaic_cohort_set, input$em_click_region_set, input$em_click_detail_set)
      
      data <- data %>% filter(communication_type == input$comm_type_set, send_date == input$send_date_set, customer_status == input$cust_status_set)
      
      if (input$mosaic_cohort_set != "All") {
        req(input$mosaic_letter_set) # Require letter only if a specific cohort is chosen
        data <- data %>% filter(mosaic_cohort == input$mosaic_cohort_set)
        
        # *** CHANGE: Conditionally filter by mosaic_letter ***
        if (input$mosaic_letter_set != "All") {
          data <- data %>% filter(mosaic_letter == input$mosaic_letter_set)
        }
      }
      
      if (input$em_click_region_set != "All") data <- data %>% filter(em_click_region == input$em_click_region_set)
      if (input$em_click_detail_set != "All") data <- data %>% filter(em_click_detail == input$em_click_detail_set)
    }
    
    return(data)
  })
  
  
  # E. OUTPUTS (PLOTS AND TABLES) -------------------------------------------
  
  # -- E1. Dynamic Title for the Graph --
  output$graph_title <- renderText({
    req(filtered_data())
    
    title <- if (input$analysis_type == "Member Engagement") {
      paste("Journey:", input$campaign_set, "- Mem Type:", input$membership_type_set, "- Variant:", input$test_variant_set)
    } else {
      # *** CHANGE: Update title to handle "All" for mosaic letter ***
      base_title <- paste("Journey:", input$comm_type_set, "-", input$mosaic_cohort_set)
      if (input$mosaic_cohort_set != "All") {
        base_title <- paste(base_title, "- Mosaic:", input$mosaic_letter_set)
      }
      base_title
    }
    
    if (isTRUE(input$use_content_type)) paste(title, "(Grouped by Content Type)") else title
  })
  
  # -- E2. Journey Graph (Unified Logic) --
  output$network_plot <- renderVisNetwork({
    req(nrow(filtered_data()) > 0)
    
    grouping_col <- if (isTRUE(input$use_content_type)) "content_type" else "content_title"
    
    # 1. Calculate transition metrics
    metrics <- filtered_data() %>%
      group_by(cust_id) %>% arrange(timestamp) %>%
      mutate(from = .data[[grouping_col]], to = lead(.data[[grouping_col]])) %>%
      ungroup() %>% filter(!is.na(to)) %>%
      group_by(from, to) %>% summarise(original_width = sum(visits, na.rm = TRUE), .groups = 'drop')
    
    validate(need(nrow(metrics) > 0, "No journeys with more than one page view found for this selection."))
    
    # 2. Robustly scale the edge widths
    new_min <- 1; new_max <- 15
    if (max(metrics$original_width, na.rm = TRUE) == min(metrics$original_width, na.rm = TRUE)) {
      metrics$scaled_width <- new_min
    } else {
      metrics <- metrics %>%
        mutate(scaled_width = new_min + ((original_width - min(original_width)) / (max(original_width) - min(original_width))) * (new_max - new_min))
    }
    
    # 3. Create nodes with robust group lookup
    nodes <- tibble(id = unique(c(metrics$from, metrics$to)), label = id, title = id)
    
    if (isTRUE(input$use_content_type)) {
      nodes <- nodes %>% mutate(group = id)
    } else {
      lookup <- filtered_data() %>%
        group_by(content_title) %>%
        summarise(content_type = first(na.omit(content_type)), .groups = 'drop')
      
      nodes <- nodes %>% left_join(lookup, by = c("id" = "content_title")) %>% rename(group = content_type)
    }
    
    # 4. Create edges with correct column names and tooltip
    edges <- metrics %>%
      mutate(
        width = scaled_width, # Use 'width' for the visual property
        arrows = "to",
        title = paste("From:", from, "<br>To:", to, "<br>Total Visits:", original_width) # Use original count for tooltip
      ) %>%
      select(from, to, width, arrows, title)
    
    # 5. Render the network
    visNetwork(nodes, edges, width = "100%") %>%
      visIgraphLayout(layout = "layout_nicely") %>%
      visNodes(shape = "dot", size = 16) %>%
      visEdges(shadow = FALSE, color = list(color = "#0085AF", highlight = "#C62F4B")) %>%
      visOptions(highlightNearest = list(enabled = TRUE, degree = 1, hover = TRUE), nodesIdSelection = TRUE, selectedBy = "group") %>%
      visPhysics(solver = "forceAtlas2Based")
  })
  
  # -- E3. Data Tables --
  render_dt <- function(data, caption) {
    DT::renderDataTable(DT::datatable(data, caption = caption, options = list(pageLength = 5, scrollX = TRUE), rownames = FALSE))
  }
  
  # Helper function for creating path tables
  create_paths_table <- function(data, group_vars, journey_filter = NULL) {
    
    journeys <- data %>%
      group_by(across(all_of(c(group_vars, "cust_id"))))
    
    if (!is.null(substitute(journey_filter))) {
      journeys <- journeys %>% filter({{ journey_filter }})
    }
    
    journeys %>%
      arrange(timestamp, .by_group = TRUE) %>%
      mutate(next_page = lead(content_title)) %>%
      ungroup() %>%
      filter(!is.na(next_page)) %>%
      group_by(across(all_of(c(group_vars, "from" = "content_title", "to" = "next_page")))) %>%
      summarise(transition_count = n(), .groups = 'drop') %>%
      group_by(across(all_of(group_vars))) %>%
      slice_max(order_by = transition_count, n = 10) %>%
      arrange(desc(across(all_of(group_vars[1]))))
  }
  
  # Reports Tab Tables
  output$engagement_comparison_table <- render_dt({
    if (input$analysis_type == "Member Engagement") {
      base_data() %>% group_by(campaign, membership_type, cust_id, visit_num) %>% summarise(len = n(), .groups = 'drop') %>% group_by(campaign, membership_type) %>% summarise(journeys = n(), avg_len = round(mean(len), 1))
    } else {
      base_data() %>% group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>% summarise(len = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_cohort) %>% summarise(journeys = n(), avg_len = round(mean(len), 1))
    }
  }, "Engagement Metrics Comparison")
  
  output$top_pages_comparison_table <- render_dt({
    if (input$analysis_type == "Member Engagement") {
      base_data() %>% group_by(campaign, membership_type, content_title) %>% summarise(views = n(), .groups = 'drop') %>% group_by(campaign, membership_type) %>% slice_max(order_by = views, n = 10)
    } else {
      base_data() %>% group_by(communication_type, mosaic_cohort, content_title) %>% summarise(views = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_cohort) %>% slice_max(order_by = views, n = 10)
    }
  }, "Top 10 Most Visited Pages")
  
  output$top_paths_comparison_table <- render_dt({
    groups <- if (input$analysis_type == "Member Engagement") c("campaign", "membership_type") else c("communication_type", "mosaic_cohort")
    create_paths_table(base_data(), groups)
  }, "Top 10 Most Common Paths")
  
  # PROSPECTS-ONLY Tables
  output$engagement_comparison_mosaic_table <- render_dt({
    req(input$analysis_type == "Prospects")
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% summarise(len = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% summarise(journeys = n(), avg_len = round(mean(len), 1))
  }, "Engagement Metrics by Mosaic Letter")
  
  output$top_pages_comparison_mosaic_table <- render_dt({
    req(input$analysis_type == "Prospects")
    base_data() %>% group_by(communication_type, mosaic_letter, content_title) %>% summarise(views = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = views, n = 10)
  }, "Top 10 Pages by Mosaic Letter")
  
  output$top_paths_comparison_mosaic_table <- render_dt({
    req(input$analysis_type == "Prospects")
    create_paths_table(base_data(), c("communication_type", "mosaic_letter"))
  }, "Top 10 Paths by Mosaic Letter")
  
  # Funnel Tables
  get_funnel_group_vars <- reactive({ if (input$analysis_type == "Member Engagement") c("campaign", "membership_type") else c("communication_type", "mosaic_letter") })
  
  output$top_paths_mem_funnel_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(str_detect(page_name, 'M\\|Membership\\|Step 1.0')))}, "Top Paths to Membership Funnel")
  output$top_paths_mem_funnel_revenue_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(membership_revenue > 0))}, "Top Paths to Membership Conversion")
  output$top_paths_renew_funnel_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(content_type == 'Renew'))}, "Top Paths to Renewal Funnel")
  output$top_paths_renew_funnel_revenue_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(renew_revenue > 0))}, "Top Paths to Renewal Conversion")
  output$top_paths_donate_funnel_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(content_type == 'Donate'))}, "Top Paths to Donations Funnel")
  output$top_paths_donate_funnel_revenue_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(donate_revenue > 0))}, "Top Paths to Donations Conversion")
  output$top_paths_shop_funnel_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(content_type == 'ShopHome'))}, "Top Paths to Shop Pages")
  output$top_paths_shop_funnel_revenue_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(shop_revenue > 0))}, "Top Paths to Shop Conversion")
  output$top_paths_holidays_funnel_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(str_detect(content_type, "Holidays")))}, "Top Paths to Holidays Pages")
  output$top_paths_holidays_funnel_revenue_table <- render_dt({ create_paths_table(base_data(), get_funnel_group_vars(), journey_filter = any(holiday_revenue > 0))}, "Top Paths to Holiday Conversion")
  
  
  # F. DOWNLOAD HANDLER -----------------------------------------------------
  output$download_report <- downloadHandler(
    filename = function() {
      paste0(gsub(" ", "_", input$analysis_type), "_Report_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx")
    },
    content = function(file) {
      
      data_to_export <- base_data()
      
      if (input$analysis_type == "Member Engagement") {
        # --- ME Report Generation ---
        group_vars <- c("campaign", "membership_type")
        list_of_datasets <- list(
          "Engagement Comparison" = data_to_export %>% group_by(campaign, membership_type, cust_id) %>% summarise(len = n(), .groups = 'drop') %>% group_by(campaign, membership_type) %>% summarise(journeys = n(), avg_len = round(mean(len), 1)),
          "Top Pages Comparison" = data_to_export %>% group_by(campaign, membership_type, content_title) %>% summarise(views = n(), .groups = 'drop') %>% group_by(campaign, membership_type) %>% slice_max(order_by = views, n = 10),
          "Top Paths Comparison" = create_paths_table(data_to_export, group_vars),
          "Membership Funnel" = create_paths_table(data_to_export, group_vars, journey_filter = any(str_detect(page_name, 'M\\|Membership\\|Step 1.0'))),
          "Membership Conversion" = create_paths_table(data_to_export, group_vars, journey_filter = any(membership_revenue > 0)),
          "Renewal Funnel" = create_paths_table(data_to_export, group_vars, journey_filter = any(content_type == 'Renew')),
          "Renewal Conversion" = create_paths_table(data_to_export, group_vars, journey_filter = any(renew_revenue > 0))
          # ... add other ME funnel tables here if needed ...
        )
      } else {
        # --- Prospects Report Generation ---
        group_vars_cohort <- c("communication_type", "mosaic_cohort")
        group_vars_letter <- c("communication_type", "mosaic_letter")
        list_of_datasets <- list(
          "Engagement By Cohort" = data_to_export %>% group_by(communication_type, mosaic_cohort, cust_id) %>% summarise(len = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_cohort) %>% summarise(journeys = n(), avg_len = round(mean(len), 1)),
          "Engagement By Letter" = data_to_export %>% group_by(communication_type, mosaic_letter, cust_id) %>% summarise(len = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% summarise(journeys = n(), avg_len = round(mean(len), 1)),
          "Top Pages By Cohort" = data_to_export %>% group_by(communication_type, mosaic_cohort, content_title) %>% summarise(views = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_cohort) %>% slice_max(order_by = views, n = 10),
          "Top Pages By Letter" = data_to_export %>% group_by(communication_type, mosaic_letter, content_title) %>% summarise(views = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = views, n = 10),
          "Top Paths By Cohort" = create_paths_table(data_to_export, group_vars_cohort),
          "Top Paths By Letter" = create_paths_table(data_to_export, group_vars_letter),
          "Membership Conversion" = create_paths_table(data_to_export, group_vars_letter, journey_filter = any(membership_revenue > 0))
          # ... add other PRO funnel tables here if needed ...
        )
      }
      
      wb <- createWorkbook()
      for (sheet_name in names(list_of_datasets)) {
        addWorksheet(wb, sheet_name)
        writeData(wb, sheet = sheet_name, x = list_of_datasets[[sheet_name]], headerStyle = createStyle(textDecoration = "bold"))
        setColWidths(wb, sheet = sheet_name, cols = 1:ncol(list_of_datasets[[sheet_name]]), widths = "auto")
      }
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# 5. RUN THE APPLICATION ----------------------------------------------------
shinyApp(ui = ui, server = server)
