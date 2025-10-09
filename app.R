# 1. LOAD LIBRARIES ---------------------------------------------------------
library(shiny)
library(tidyverse)
library(lubridate)
library(openxlsx)
library(visNetwork)
library(DT)
library(scales)

# Increase the maximum upload size to 30MB
options(shiny.maxRequestSize = 30*1024^2)

# 2. USER INTERFACE (UI) ----------------------------------------------------
ui <- fluidPage(
  
  # Set a theme for a cleaner look
  theme = bslib::bs_theme(version = 5, bootswatch = "cosmo"),
  
  # App Title
  titlePanel("Journey Analyser"),
  
  # Sidebar Layout
  sidebarLayout(
    
    # -- Sidebar for Controls --
    sidebarPanel(
      
      h4("Step 1: Upload Data"),
      
      fileInput("file_upload", "Upload a CSV File",
                accept = c("text/csv", "text/comma-separated-values,text/plain", ".csv")),
      
      tags$hr(),
      
      # Conditional panel that appears only after a file is uploaded
      conditionalPanel(
        condition = "output.file_uploaded",
        
        h4("Step 2: Filter Journey"),
        p("Select a cohort to visualize their journey graph."),
        
        # Checkbox to switch between content_title and content_type
        checkboxInput("use_content_type", "Group Nodes by Content Type", value = FALSE),
        
        # Dynamic UI for dropdowns will be rendered here from the server
        uiOutput("communication_type_ui"),
        uiOutput("send_date_ui"),
        uiOutput("customer_status_ui"),
        uiOutput("mosaic_cohort_ui"),
        uiOutput("mosaic_letter_ui"),
        
        tags$hr(),
        
        h4("Step 3: Download Full Report"),
        p("Generates an Excel file with all summary tables for the uploaded data."),
        downloadButton("download_report", "Download Excel Report", class = "btn-success")
      )
    ),
    
    # -- Main Panel for Outputs --
    mainPanel(
      # Conditional panel to show content only after file upload
      conditionalPanel(
        condition = "output.file_uploaded",
        
        tabsetPanel(
          type = "tabs",
          
          # Tab 1: Journey Graph
          tabPanel(
            "Journey Graph",
            icon = icon("project-diagram"),
            h3(textOutput("graph_title")),
            p("This interactive graph shows the paths users took after arriving from the selected email campaign. Nodes represent web pages, and edges represent the transitions between them."),
            visNetworkOutput("network_plot", height = "70vh")
          ),
          
          # Tab 2: Engagement Reports
          tabPanel(
            "Engagement Reports",
            icon = icon("chart-bar"),
            h3("Engagement Metrics"),
            p("Summary of user engagement across different segments."),
            DT::dataTableOutput("engagement_comparison_table"),
            tags$hr(),
            DT::dataTableOutput("engagement_comparison_mosaic_table"),
            tags$hr(),
            h3("Top 10 Most Visited Pages"),
            p("The most frequently viewed pages, grouped by segment."),
            DT::dataTableOutput("top_pages_comparison_table"),
            tags$hr(),
            DT::dataTableOutput("top_pages_comparison_mosaic_table")
          ),
          
          # Tab 3: Path & Funnel Analysis
          tabPanel(
            "Path & Funnel Analysis",
            icon = icon("filter"),
            h3("Top 10 Most Common User Paths"),
            p("The most common page-to-page transitions (A -> B) observed for each segment."),
            DT::dataTableOutput("top_paths_comparison_table"),
            tags$hr(),
            DT::dataTableOutput("top_paths_comparison_mosaic_table"),
            tags$hr(),
            h3("Conversion Funnel Paths"),
            p("Top 10 paths within journeys that either entered a conversion funnel or resulted in a conversion."),
            h4("Membership"),
            DT::dataTableOutput("top_paths_mem_funnel_table"),
            DT::dataTableOutput("top_paths_mem_funnel_revenue_table"),
            tags$hr(),
            h4("Renewals"),
            DT::dataTableOutput("top_paths_renew_funnel_table"),
            DT::dataTableOutput("top_paths_renew_funnel_revenue_table"),
            tags$hr(),
            h4("Donations"),
            DT::dataTableOutput("top_paths_donate_funnel_table"),
            DT::dataTableOutput("top_paths_donate_funnel_revenue_table"),
            tags$hr(),
            h4("Shop"),
            DT::dataTableOutput("top_paths_shop_funnel_table"),
            DT::dataTableOutput("top_paths_shop_funnel_revenue_table"),
            tags$hr(),
            h4("Holidays"),
            DT::dataTableOutput("top_paths_holidays_funnel_table"),
            DT::dataTableOutput("top_paths_holidays_funnel_revenue_table")
          )
        )
      ),
      # Message to show before file is uploaded
      conditionalPanel(
        condition = "!output.file_uploaded",
        h2("Please upload a CSV file to begin analysis.", style = "text-align: center; color: #888; margin-top: 50px;")
      )
    )
  )
)

# 3. SERVER LOGIC -----------------------------------------------------------
server <- function(input, output, session) {
  
  # Reactive expression to check if a file has been uploaded
  output$file_uploaded <- reactive({
    return(!is.null(input$file_upload))
  })
  outputOptions(output, 'file_uploaded', suspendWhenHidden = FALSE)
  
  # A. REACTIVE DATA: Read and clean the uploaded file ----------------------
  base_data <- reactive({
    # Require a file to be uploaded before proceeding
    req(input$file_upload)
    
    # Read the uploaded CSV file
    Prospect_Email_All_Groups <- read_csv(
      input$file_upload$datapath,
      col_types = cols(
        `MCVID (v29) (evar29)` = col_character(),
        `Tracking Code` = col_character(),
        `Page Name (v26) (evar26)` = col_character(),
        `URL (v8) (evar8)` = col_character(),
        `24H Clock by Minute` = col_time(format = "%H:%M"),
        `Content Type (v7) (evar7)` = col_character(),
        `Content Title (v22) (evar22)` = col_character(),
        `Visit Number` = col_integer(),
        `Exit Links` = col_character()
      )
    )
    
    # All data processing and cleaning code from the original script
    prospect_data <- Prospect_Email_All_Groups %>%
      rename(
        date = Date,
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
      ) %>%
      mutate(date = as.Date(date, format = "%B %d, %Y")) %>%
      mutate(timestamp = as.POSIXct(paste(date, session_time), format = "%Y-%m-%d %H:%M:%S")) %>%
      drop_na(cust_id) %>%
      filter(!str_detect(tracking_code, "%")) %>%
      mutate(tracking_code_orig = tracking_code) %>%
      separate(tracking_code, into = c("channel", "region", "campaign", "send_date_str", "cohort"), sep = "_") %>%
      filter(str_detect(cohort, "NT")) %>%
      separate(cohort, into = c("cohort", "em_click_region"), sep = "-") %>%
      mutate(send_date = as.Date(send_date_str, format = "%d%m%Y")) %>%
      mutate(
        source_code = substr(cohort, 3, 8),
        date_str = substr(cohort, 9, 14),
        comm_type_code = substr(cohort, 15, 16),
        cust_status_code = substr(cohort, 17, 17),
        mosaic_letter = substr(cohort, 18, 18),
        date = dmy(date_str),
        communication_type = case_when(
          comm_type_code == "PW" ~ "Prospect Welcome",
          comm_type_code == "PN" ~ "Prospect Newsletter",
          comm_type_code == "DO" ~ "Double Opt-In",
          TRUE ~ NA_character_
        ),
        customer_status = case_when(
          cust_status_code == "1" ~ "Known",
          cust_status_code == "2" ~ "Net New",
          TRUE ~ NA_character_
        ),
        mosaic_cohort = case_when(
          mosaic_letter %in% c("K", "O") ~ "Cohort 1",
          mosaic_letter %in% c("A") ~ "Cohort 2.1",
          mosaic_letter %in% c("G", "H") ~ "Cohort 2.2",
          mosaic_letter %in% c("B", "C", "E") ~ "Cohort 3",
          mosaic_letter %in% c("D", "F", "I", "J", "L", "M", "N", "U") ~ "Everyone Else",
          TRUE ~ NA_character_
        )
      ) %>%
      filter(substr(source_code, 1, 3) %in% c('DCP', 'MCC', 'DCN')) %>%
      drop_na(communication_type, customer_status, mosaic_cohort, mosaic_letter)
    
    return(prospect_data)
  })
  
  # B. DYNAMIC UI: Populate dropdowns based on uploaded data -----------------
  output$communication_type_ui <- renderUI({
    req(base_data())
    choices <- unique(base_data()$communication_type)
    selectInput("comm_type_set", "Communication Type", choices = choices)
  })
  
  output$send_date_ui <- renderUI({
    req(base_data())
    choices <- unique(base_data()$send_date)
    selectInput("send_date_set", "Send Date", choices = sort(choices, decreasing = TRUE))
  })
  
  output$customer_status_ui <- renderUI({
    req(base_data())
    choices <- unique(base_data()$customer_status)
    selectInput("cust_status_set", "Customer Status", choices = choices)
  })
  
  output$mosaic_cohort_ui <- renderUI({
    req(base_data())
    choices <- unique(base_data()$mosaic_cohort)
    selectInput("mosaic_cohort_set", "Mosaic Cohort", choices = choices)
  })
  
  output$mosaic_letter_ui <- renderUI({
    req(base_data(), input$mosaic_cohort_set)
    # Filter letters based on selected cohort
    choices <- base_data() %>%
      filter(mosaic_cohort == input$mosaic_cohort_set) %>%
      pull(mosaic_letter) %>%
      unique()
    selectInput("mosaic_letter_set", "Mosaic Letter", choices = choices)
  })
  
  # C. REACTIVE DATA: Filter data based on user inputs ----------------------
  filtered_data <- reactive({
    # Require all inputs to be available before filtering
    req(
      base_data(),
      input$comm_type_set,
      input$send_date_set,
      input$cust_status_set,
      input$mosaic_cohort_set,
      input$mosaic_letter_set
    )
    
    base_data() %>%
      filter(
        communication_type == input$comm_type_set,
        send_date == input$send_date_set,
        customer_status == input$cust_status_set,
        mosaic_cohort == input$mosaic_cohort_set,
        mosaic_letter == input$mosaic_letter_set
      )
  })
  
  # D. RENDER OUTPUTS: Create plots and tables ------------------------------
  
  # -- D1. Dynamic Title for the Graph --
  output$graph_title <- renderText({
    req(filtered_data())
    base_title <- paste(
      "Journey:", input$comm_type_set, "-",
      input$mosaic_cohort_set, "-",
      "Mosaic:", input$mosaic_letter_set, "-",
      input$cust_status_set, "-",
      input$send_date_set
    )
    # Append a note if the checkbox is ticked
    if (isTRUE(input$use_content_type)) {
      paste(base_title, "(Grouped by Content Type)")
    } else {
      base_title
    }
  })
  
  # -- D2. Journey Graph --
  output$network_plot <- renderVisNetwork({
    # Require filtered data to have rows
    req(nrow(filtered_data()) > 0)
    
    # Conditionally define the column to use for grouping based on the checkbox
    grouping_column <- if (isTRUE(input$use_content_type)) {
      "content_type"
    } else {
      "content_title"
    }
    
    # Calculate transition widths by summing the 'visits' metric
    transition_metrics <- filtered_data() %>%
      group_by(cust_id) %>%
      arrange(timestamp) %>%
      mutate(
        from = .data[[grouping_column]],
        to = lead(.data[[grouping_column]])
      ) %>%
      ungroup() %>%
      filter(!is.na(to)) %>% # Remove the last step of each journey
      group_by(from, to) %>%
      summarise(
        width = sum(visits, na.rm = TRUE),
        .groups = 'drop'
      )
    
    # Stop if there are no transitions to plot
    validate(need(nrow(transition_metrics) > 0, "No user journeys with more than one page view found for this selection."))
    
    # Scale the widths to a manageable visual range
    new_min <- 1
    new_max <- 15
    transition_metrics <- transition_metrics %>%
      mutate(
        scaled_width = new_min + ((width - min(width)) / (max(width) - min(width))) * (new_max - new_min)
      )
    if (any(is.nan(transition_metrics$scaled_width))) {
      transition_metrics$scaled_width <- new_min
    }
    
    # Create nodes from ALL unique pages in the filtered data
    all_pages <- unique(filtered_data()[[grouping_column]])
    nodes <- tibble(
      id = all_pages,
      label = id,
      shape = "dot",
      size = 15,
      title = id
    )
    
    # Add a 'group' column for coloring
    if (isTRUE(input$use_content_type)) {
      nodes <- nodes %>% mutate(group = id)
    } else {
      page_type_lookup <- filtered_data() %>%
        group_by(content_title) %>%
        summarise(content_type = first(na.omit(content_type)), .groups = 'drop')
      nodes <- nodes %>%
        left_join(page_type_lookup, by = c("id" = "content_title")) %>%
        rename(group = content_type)
    }
    
    # Create edges using the scaled widths for the visual and original for the tooltip
    edges <- transition_metrics %>%
      mutate(
        width = scaled_width,
        arrows = "to",
        title = paste("From:", from, "<br>To:", to, "<br>Total Visits:", width)
      )
    
    # Generate the graph
    visNetwork(nodes, edges, width = "100%") %>%
      visIgraphLayout(layout = "layout_nicely") %>%
      visNodes(
        shape = "dot",
        shadow = list(enabled = TRUE, size = 10)
      ) %>%
      visEdges(shadow = FALSE, color = list(color = "#0085AF", highlight = "#C62F4B")) %>%
      visOptions(
        # MODIFIED: highlight degree changed to 1
        highlightNearest = list(enabled = TRUE, degree = 1, hover = TRUE), 
        nodesIdSelection = TRUE,
        selectedBy = "group"
      ) %>%
      visLegend(useGroups = TRUE, main = "Content Type", width = 0.2, position = "right", ncol = 2) %>%
      visPhysics(solver = "forceAtlas2Based")
  })
  
  # -- D3. Summary and Funnel Tables --
  # Note: All summary tables use the full `base_data()` to show comparisons
  
  render_dt <- function(data, caption) {
    DT::renderDataTable(
      DT::datatable(
        data,
        caption = htmltools::tags$caption(style = 'caption-side: top; text-align: left; font-weight: bold;', caption),
        options = list(pageLength = 5, scrollX = TRUE),
        rownames = FALSE
      )
    )
  }
  
  # Engagement Tables
  output$engagement_comparison_table <- render_dt({
    base_data() %>%
      group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>%
      summarise(journey_length = n(), .groups = 'drop') %>%
      group_by(communication_type, mosaic_cohort) %>%
      summarise(total_journeys = n(), avg_journey_length = round(mean(journey_length), digits = 1)) %>%
      arrange(desc(communication_type))
  }, "Engagement Metrics by Cohort")
  
  output$engagement_comparison_mosaic_table <- render_dt({
    base_data() %>%
      group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
      summarise(journey_length = n(), .groups = 'drop') %>%
      group_by(communication_type, mosaic_letter) %>%
      summarise(total_journeys = n(), avg_journey_length = round(mean(journey_length), digits = 1)) %>%
      arrange(desc(communication_type))
  }, "Engagement Metrics by Mosaic Letter")
  
  # Top Pages Tables
  output$top_pages_comparison_table <- render_dt({
    base_data() %>%
      group_by(communication_type, mosaic_cohort, content_title) %>%
      summarise(page_views = n(), .groups = 'drop') %>%
      group_by(communication_type, mosaic_cohort) %>%
      slice_max(order_by = page_views, n = 10) %>%
      arrange(desc(communication_type))
  }, "Top 10 Pages by Cohort")
  
  output$top_pages_comparison_mosaic_table <- render_dt({
    base_data() %>%
      group_by(communication_type, mosaic_letter, content_title) %>%
      summarise(page_views = n(), .groups = 'drop') %>%
      group_by(communication_type, mosaic_letter) %>%
      slice_max(order_by = page_views, n = 10) %>%
      arrange(desc(communication_type))
  }, "Top 10 Pages by Mosaic Letter")
  
  # Top Paths Tables
  output$top_paths_comparison_table <- render_dt({
    base_data() %>%
      group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>%
      arrange(timestamp, .by_group = TRUE) %>%
      mutate(next_page = lead(content_title)) %>%
      ungroup() %>%
      filter(!is.na(next_page)) %>%
      group_by(communication_type, mosaic_cohort, from = content_title, to = next_page) %>%
      summarise(transition_count = n(), .groups = 'drop') %>%
      group_by(communication_type, mosaic_cohort) %>%
      slice_max(order_by = transition_count, n = 10) %>%
      arrange(desc(communication_type))
  }, "Top 10 Paths by Cohort")
  
  output$top_paths_comparison_mosaic_table <- render_dt({
    base_data() %>%
      group_by(communication_type, mosaic_letter, cust_id, visit_num) %>%
      arrange(timestamp, .by_group = TRUE) %>%
      mutate(next_page = lead(content_title)) %>%
      ungroup() %>%
      filter(!is.na(next_page)) %>%
      group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>%
      summarise(transition_count = n(), .groups = 'drop') %>%
      group_by(communication_type, mosaic_letter) %>%
      slice_max(order_by = transition_count, n = 10) %>%
      arrange(desc(communication_type))
  }, "Top 10 Paths by Mosaic Letter")
  
  # Funnel Path Tables
  output$top_paths_mem_funnel_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(page_name == 'M|Membership|Step 1.0 - Your Details')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Membership Funnel")
  
  output$top_paths_mem_funnel_revenue_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(membership_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Membership Conversion")
  
  output$top_paths_renew_funnel_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(content_type == 'Renew')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Renewal Funnel")
  
  output$top_paths_renew_funnel_revenue_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(renew_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Renewal Conversion")
  
  output$top_paths_donate_funnel_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(content_type == 'Donate')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Donations Funnel")
  
  output$top_paths_donate_funnel_revenue_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(donate_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Donations Conversion")
  
  output$top_paths_shop_funnel_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(content_type == 'ShopHome')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Shop Pages")
  
  output$top_paths_shop_funnel_revenue_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(shop_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Shop Conversion")
  
  output$top_paths_holidays_funnel_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(str_detect(content_type, "Holidays"))) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Holidays Pages")
  
  output$top_paths_holidays_funnel_revenue_table <- render_dt({
    base_data() %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(holiday_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
  }, "Top Paths to Holiday Conversion")
  
  
  # E. DOWNLOAD HANDLER: Create and serve the Excel file --------------------
  output$download_report <- downloadHandler(
    filename = function() {
      paste0("Journey_Analytics_Report_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx")
    },
    content = function(file) {
      
      # This code re-creates all data frames needed for the export
      # using the uploaded data.
      
      data_to_export <- base_data()
      
      engagement_comparison <- data_to_export %>% group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>% summarise(journey_length = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_cohort) %>% summarise(total_journeys = n(), avg_journey_length = round(mean(journey_length), digits = 1)) %>% arrange(desc(communication_type))
      engagement_comparison_mosaic <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% summarise(journey_length = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% summarise(total_journeys = n(), avg_journey_length = round(mean(journey_length), digits = 1)) %>% arrange(desc(communication_type))
      top_pages_comparison <- data_to_export %>% group_by(communication_type, mosaic_cohort, content_title) %>% summarise(page_views = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_cohort) %>% slice_max(order_by = page_views, n = 10) %>% arrange(desc(communication_type))
      top_pages_comparison_mosiac <- data_to_export %>% group_by(communication_type, mosaic_letter, content_title) %>% summarise(page_views = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = page_views, n = 10) %>% arrange(desc(communication_type))
      top_paths_comparison <- data_to_export %>% group_by(communication_type, mosaic_cohort, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% ungroup() %>% filter(!is.na(next_page)) %>% group_by(communication_type, mosaic_cohort, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_cohort) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_comparison_mosaic <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% ungroup() %>% filter(!is.na(next_page)) %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_mem_funnel <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(page_name == 'M|Membership|Step 1.0 - Your Details')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_mem_funnel_revenue <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(membership_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_renew_funnel <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(content_type == 'Renew')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_renew_funnel_revenue <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(renew_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_donate_funnel <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(content_type == 'Donate')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_donate_funnel_revenue <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(donate_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_shop_funnel <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(content_type == 'ShopHome')) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_shop_funnel_revenue <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(shop_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_holidays_funnel <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(str_detect(content_type, "Holidays"))) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      top_paths_holidays_funnel_revenue <- data_to_export %>% group_by(communication_type, mosaic_letter, cust_id, visit_num) %>% arrange(timestamp, .by_group = TRUE) %>% mutate(next_page = lead(content_title)) %>% filter(any(holiday_revenue > 0)) %>% ungroup() %>% group_by(communication_type, mosaic_letter, from = content_title, to = next_page) %>% summarise(transition_count = n(), .groups = 'drop') %>% group_by(communication_type, mosaic_letter) %>% slice_max(order_by = transition_count, n = 10) %>% arrange(desc(communication_type))
      
      list_of_datasets <- list(
        "Engagement By Cohort" = engagement_comparison,
        "Engagement By Mosaic Letter" = engagement_comparison_mosaic,
        "Top Pages By Cohort" = top_pages_comparison,
        "Top Pages By Mosaic Letter" = top_pages_comparison_mosiac,
        "Top Paths By Cohort" = top_paths_comparison,
        "Top Paths By Mosiac Letter" = top_paths_comparison_mosaic,
        "Journey to Membership Funnel" = top_paths_mem_funnel,
        "Journey to Membership Conversion" = top_paths_mem_funnel_revenue,
        "Journey to Renewal Funnel" = top_paths_renew_funnel,
        "Journey to Renewal Conversion" = top_paths_renew_funnel_revenue,
        "Journey to Donations Funnel" = top_paths_donate_funnel,
        "Journey to Donations Conversion" = top_paths_donate_funnel_revenue,
        "Journey to Shop Pages" = top_paths_shop_funnel,
        "Journey to Shop Conversion" = top_paths_shop_funnel_revenue,
        "Journey to Holidays Pages" = top_paths_holidays_funnel,
        "Journey to Holiday Conversion" = top_paths_holidays_funnel_revenue
      )
      
      # Use the openxlsx logic from the original script
      wb <- createWorkbook()
      
      header_style <- createStyle(textDecoration = "bold", fontSize = 14, fgFill = "#DDEBF7")
      
      for (sheet_name in names(list_of_datasets)) {
        addWorksheet(wb, sheet_name)
        writeData(wb, sheet = sheet_name, x = list_of_datasets[[sheet_name]],
                  headerStyle = createStyle(textDecoration = "bold", border = "TopBottomLeftRight"))
        setColWidths(wb, sheet = sheet_name, cols = 1:ncol(list_of_datasets[[sheet_name]]), widths = "auto")
      }
      
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# 4. RUN THE APPLICATION ----------------------------------------------------
shinyApp(ui = ui, server = server)