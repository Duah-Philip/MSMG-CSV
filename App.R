

# app.R - SimaPro CSV Converter
# A Shiny application to convert Excel-based LCI data to SimaPro CSV format



#options(repos = c(CRAN = "https://cloud.r-project.org"))
#options(expressions = 5e5)


library(shiny)
library(shinydashboard)
library(readxl)
library(dplyr)
library(DT)
library(tools)


# Function to convert dataframe to SimaPro CSV format
to_spcsv <- function(dataframe, file_path) {
  # Function for debug logging - capture in a list for later review
  debug_info <- list()
  debug_log <- function(msg) {
    debug_info[[length(debug_info) + 1]] <<- msg
    print(paste("DEBUG:", msg))  # Print to console for real-time debugging
  }
  
  # Log dataframe dimensions
  debug_log(paste("Dataframe dimensions:", nrow(dataframe), "rows,", ncol(dataframe), "columns"))
  
  # Initialize connection as NULL for proper error handling
  con <- NULL
  
  # Ensure connection is closed on function exit, even if errors occur
  on.exit({
    if (!is.null(con) && isOpen(con)) {
      close(con)
      debug_log("Connection closed properly")
    }
    # Return debug info in case of error
    if (exists("error_occurred") && error_occurred) {
      writeLines(paste(unlist(debug_info), collapse = "\n"), paste0(file_path, ".debug.log"))
    }
  })
  
  # Capture potential errors
  error_occurred <- FALSE
  
  # More lenient error checking - still check but with reduced requirements
  if (nrow(dataframe) < 3) {
    debug_log("ERROR: Input data has too few rows")
    stop("Input data must have at least 3 rows")
  }
  if (ncol(dataframe) < 3) {
    debug_log("ERROR: Input data has too few columns")
    stop("Input data must have at least 3 columns")
  }
  
  # Log first few rows for debugging
  debug_log("First row content:")
  if (nrow(dataframe) > 0) {
    debug_log(paste(as.character(dataframe[1,]), collapse = " | "))
  }
  
  # Open destination file and print the standard heading
  tryCatch({
    con <- file(file_path, "w")
    writeLines("{CSV separator: Semicolon}", con)
    writeLines("{CSV Format version: 7.0.0}", con)
    writeLines("{Decimal separator: .}", con)
    writeLines("{Date separator: /}", con)
    writeLines("{Short date format: dd/MM/yyyy}", con)
    writeLines("", con)
    debug_log(paste("File opened successfully:", file_path))
  }, error = function(e) {
    debug_log(paste("Error opening file:", e$message))
    error_occurred <- TRUE
    stop(paste("Failed to create output file:", e$message))
  })
  
  # List of fields required
  fields <- c(
    "Process", "Category type", "Time Period", "Geography",
    "Technology", "Representativeness", "Multiple output allocation",
    "Substitution allocation", "Cut off rules", "Capital goods",
    "Boundary with nature", "Record", "Generator", "Literature references",
    "Collection method", "Data treatment", "Verification",
    "Products", "Materials/fuels", "Resources", "Emissions to air",
    "Emissions to water", "Emissions to soil", "Final waste flows",
    "Non material emission", "Social issues", "Economic issues",
    "Waste to treatment", "End"
  )
  
  # Standard value of these fields - using a list to properly handle both scalar and vector values
  fields_value <- list(
    "", "", "Unspecified", "Unspecified", "Unspecified",
    "Unspecified", "Unspecified", "Unspecified", "Unspecified",
    "Unspecified", "Unspecified", "", "", "", "", "",
    "Comment", "", list(), list(), list(), list(), list(), list(), 
    "", list(), list(), list(), ""
  )
  
  # Convert NAs to empty strings
  dataframe[is.na(dataframe)] <- ""
  
  # Create a default process if necessary - this is a fallback measure
  default_process <- FALSE
  if (ncol(dataframe) < 5) {
    debug_log("WARNING: Excel has fewer than 5 columns, using default process")
    # Create a mock process column
    default_process <- TRUE
    dataframe <- cbind(dataframe, Default_Process = rep("", nrow(dataframe)))
    if (nrow(dataframe) > 0) dataframe[1, ncol(dataframe)] <- "Default_Process"
    debug_log(paste("New dataframe dimensions after adding default process:", 
                    nrow(dataframe), "rows,", ncol(dataframe), "columns"))
  }
  
  # Identify the processes - with extra safety measures
  tryCatch({
    # Ensure we can access the first row
    if (nrow(dataframe) >= 1) {
      if (default_process) {
        processes <- dataframe[1, ncol(dataframe), drop=FALSE]
        debug_log("Using default process name")
      } else if (ncol(dataframe) >= 5) {
        processes <- dataframe[1, 5:ncol(dataframe), drop=FALSE]
        debug_log(paste("Found", ncol(processes), "potential processes"))
      } else {
        processes <- data.frame(Default="Default_Process")
        debug_log("Excel format incorrect: Using default process name")
      }
      
      # Filter out empty process names
      if (!default_process) {
        valid_indices <- which(!is.na(processes) & processes != "", arr.ind = TRUE)
        debug_log(paste("Valid process indices:", paste(valid_indices, collapse=", ")))
        
        if (length(valid_indices) == 0) {
          debug_log("WARNING: No valid processes found, using default")
          processes <- data.frame(Default="Default_Process")
          default_process <- TRUE
        }
      }
    } else {
      debug_log("ERROR: Empty dataframe")
      stop("Dataframe has no rows")
    }
  }, error = function(e) {
    debug_log(paste("Error identifying processes:", e$message))
    error_occurred <- TRUE
    stop(paste("Failed to identify processes:", e$message))
  })
  
  # Screen through the processes - with much more robust error handling
  debug_log(paste("Starting process processing, with", length(processes), "potential processes"))
  
  # Handle case where processes might be a data frame instead of a vector
  process_count <- if(is.data.frame(processes)) ncol(processes) else length(processes)
  debug_log(paste("Process count:", process_count))
  
  for (i in 1:process_count) {
    tryCatch({
      # Process column index in the dataframe
      col_idx <- if(default_process) ncol(dataframe) else (i + 4)
      debug_log(paste("Processing column index:", col_idx))
      
      # Safety check - make sure column exists
      if (col_idx > ncol(dataframe)) {
        debug_log(paste("Column index", col_idx, "is out of bounds, skipping"))
        next  # Skip this iteration if column doesn't exist
      }
      
      # Extract process info - with checks
      if (ncol(dataframe) >= col_idx) {
        process <- dataframe[, col_idx, drop=FALSE]
        debug_log(paste("Extracted process column with", nrow(process), "rows"))
      } else {
        debug_log(paste("Cannot extract column", col_idx, "- out of bounds"))
        next
      }
      
      # Safe way to access the category type value
      fields_value[[2]] <- ""  # Default empty value
      
      if (nrow(dataframe) >= 6 && ncol(dataframe) >= col_idx) {
        category_type <- as.character(dataframe[6, col_idx])
        debug_log(paste("Category type found:", category_type))
        if (!is.na(category_type)) fields_value[[2]] <- category_type
      } else {
        debug_log("Not enough rows for category type or column doesn't exist")
      }
      
      # Products string formatting - with checks and defaults
      # Set a default product name if none exists
      product_name <- "Default_Product"  
      product_unit <- "kg"
      product_amount <- "1"
      product_comment <- ""
      
      if (nrow(dataframe) >= 1 && ncol(dataframe) >= col_idx) {
        temp <- as.character(dataframe[1, col_idx])
        if (!is.na(temp) && temp != "") product_name <- temp
      }
      
      if (nrow(dataframe) >= 2 && ncol(dataframe) >= col_idx) {
        temp <- as.character(dataframe[2, col_idx])
        if (!is.na(temp) && temp != "") product_unit <- temp
      }
      
      if (nrow(dataframe) >= 3 && ncol(dataframe) >= col_idx) {
        temp <- as.character(dataframe[3, col_idx])
        if (!is.na(temp) && temp != "") product_amount <- temp
      }
      
      if (nrow(dataframe) >= 5 && ncol(dataframe) >= col_idx) {
        temp <- as.character(dataframe[5, col_idx])
        if (!is.na(temp)) product_comment <- temp
      }
      
      debug_log(paste("Product info:", product_name, product_unit, product_amount, product_comment))
      
      products <- sprintf('"%s";"%s";"%s";"100%%";"not defined";"%s"',
                          product_name, product_unit, product_amount, product_comment)
      
      fields_value[[18]] <- products
      
      # Initialize lists for different exchange types
      matfuel_list <- c()
      raw_list <- c()
      air_list <- c()
      water_list <- c()
      soil_list <- c()
      finalwaste_list <- c()
      social_list <- c()
      economic_list <- c()
      wastetotreatment_list <- c()
      
      debug_log("Initialized exchange type lists")
      
      # Determine safe start row - default to 7 but be flexible
      start_row <- min(7, nrow(dataframe))
      if (start_row < 7) {
        debug_log(paste("WARNING: Using start row", start_row, "instead of 7"))
      }
      
      # Screen through the inputs and outputs of each process
      if (start_row <= nrow(dataframe)) {
        debug_log(paste("Processing exchanges from row", start_row, "to", nrow(dataframe)))
        
        for (j in start_row:nrow(dataframe)) {
          tryCatch({
            # Safe way to access exchange type - with checks
            exchange_type <- ""
            if (ncol(dataframe) >= 1) {
              temp <- dataframe[j, 1]
              if (!is.na(temp)) exchange_type <- as.character(temp)
            }
            
            # Safe way to access exchange value
            exchange_value <- ""
            if (ncol(dataframe) >= col_idx) {
              temp <- dataframe[j, col_idx]
              if (!is.na(temp)) exchange_value <- as.character(temp)
            }
            
            debug_log(paste("Row", j, "- Exchange type:", exchange_type, "Value:", exchange_value))
            
            # Only process non-empty values
            if (exchange_value != "") {
              # Safe way to access column 2 and 4
              exchange_name <- "Unknown"  # Default value
              if (ncol(dataframe) >= 2) {
                temp <- dataframe[j, 2]
                if (!is.na(temp)) exchange_name <- as.character(temp)
              }
              
              exchange_unit <- "kg"  # Default value
              if (ncol(dataframe) >= 4) {
                temp <- dataframe[j, 4]
                if (!is.na(temp)) exchange_unit <- as.character(temp)
              }
              
              debug_log(paste("Exchange details - Name:", exchange_name, "Unit:", exchange_unit))
              
              # Process according to exchange type
              if (exchange_type == "") {
                # Materials/fuels
                matfuel <- sprintf('"%s";"%s";"%s";"Undefined";0;0;0',
                                   exchange_name, exchange_unit, exchange_value)
                matfuel_list <- c(matfuel_list, matfuel)
                debug_log("Added to materials/fuels")
                
              } else if (tolower(exchange_type) == "raw" || exchange_type == "Raw") {
                # Resources - case insensitive match
                raw <- sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0',
                               exchange_name, "", exchange_unit, exchange_value)
                raw_list <- c(raw_list, raw)
                debug_log("Added to raw resources")
                
              } else if (tolower(exchange_type) == "air" || exchange_type == "Air") {
                # Emissions to air - case insensitive match
                air <- sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0',
                               exchange_name, "", exchange_unit, exchange_value)
                air_list <- c(air_list, air)
                debug_log("Added to air emissions")
                
              } else if (tolower(exchange_type) == "water" || exchange_type == "Water") {
                # Emissions to water - case insensitive match
                water <- sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0',
                                 exchange_name, "", exchange_unit, exchange_value)
                water_list <- c(water_list, water)
                debug_log("Added to water emissions")
                
              } else if (tolower(exchange_type) == "soil" || exchange_type == "Soil") {
                # Emissions to soil - case insensitive match
                soil <- sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0',
                                exchange_name, "", exchange_unit, exchange_value)
                soil_list <- c(soil_list, soil)
                debug_log("Added to soil emissions")
                
              } else if (tolower(exchange_type) == "waste" || exchange_type == "Waste") {
                # Final waste flows - case insensitive match
                finalwaste <- sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0',
                                      exchange_name, "", exchange_unit, exchange_value)
                finalwaste_list <- c(finalwaste_list, finalwaste)
                debug_log("Added to waste flows")
                
              } else if (tolower(exchange_type) == "social" || exchange_type == "Social") {
                # Social issues - case insensitive match
                social <- sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0',
                                  exchange_name, "", exchange_unit, exchange_value)
                social_list <- c(social_list, social)
                debug_log("Added to social issues")
                
              } else if (tolower(exchange_type) == "economic" || exchange_type == "Economic") {
                # Economic issues - case insensitive match
                economic <- sprintf('"%s";"%s";"%s";"%s";"Undefined";0;0;0',
                                    exchange_name, "", exchange_unit, exchange_value)
                economic_list <- c(economic_list, economic)
                debug_log("Added to economic issues")
                
              } else if (tolower(exchange_type) == "wastetotreatment" || exchange_type == "Wastetotreatment") {
                # Waste to treatment - case insensitive match
                wastetotreatment <- sprintf('"%s";"%s";"%s";"Undefined";0;0;0',
                                            exchange_name, exchange_unit, exchange_value)
                wastetotreatment_list <- c(wastetotreatment_list, wastetotreatment)
                debug_log("Added to waste to treatment")
              } else {
                debug_log(paste("Unknown exchange type:", exchange_type, "- skipping"))
              }
            }
          }, error = function(e) {
            debug_log(paste("Error processing row", j, ":", e$message))
            # Continue with next row rather than stopping completely
          })
        }
      } else {
        debug_log("No rows to process for exchanges")
      }
      
      # Assign the inputs and outputs to the fields_value list
      fields_value[[19]] <- matfuel_list
      fields_value[[20]] <- raw_list
      fields_value[[21]] <- air_list
      fields_value[[22]] <- water_list
      fields_value[[23]] <- soil_list
      fields_value[[24]] <- finalwaste_list
      fields_value[[26]] <- social_list
      fields_value[[27]] <- economic_list
      fields_value[[28]] <- wastetotreatment_list
      
      debug_log(paste("Exchange counts - Materials/fuels:", length(matfuel_list),
                      "Resources:", length(raw_list),
                      "Air:", length(air_list),
                      "Water:", length(water_list),
                      "Soil:", length(soil_list),
                      "Waste:", length(finalwaste_list),
                      "Social:", length(social_list),
                      "Economic:", length(economic_list),
                      "Waste treatment:", length(wastetotreatment_list)))
      
      # Write all fields to the file
      tryCatch({
        debug_log("Writing fields to output file")
        for (el in 1:length(fields)) {
          writeLines(fields[el], con)
          
          if (is.atomic(fields_value[[el]]) && length(fields_value[[el]]) <= 1) {
            # Handle single value fields
            writeLines(as.character(fields_value[[el]]), con)
            debug_log(paste("Wrote single value field:", fields[el]))
          } else {
            # Write all elements of list fields
            if (length(fields_value[[el]]) > 0) {
              for (item in fields_value[[el]]) {
                writeLines(as.character(item), con)
              }
              debug_log(paste("Wrote list field:", fields[el], "with", length(fields_value[[el]]), "items"))
            } else {
              # Write empty line for empty lists
              writeLines("", con)
              debug_log(paste("Wrote empty list field:", fields[el]))
            }
          }
          writeLines("", con)
        }
        debug_log("Successfully completed writing process")
      }, error = function(e) {
        debug_log(paste("Error writing to output file:", e$message))
        error_occurred <- TRUE
      })
      
    }, error = function(e) {
      debug_log(paste("Process loop error:", e$message))
      error_occurred <- TRUE
    })
  }
  
  # Write debug log if errors occurred
  if (exists("error_occurred") && error_occurred) {
    debug_log_file <- paste0(file_path, ".debug.log")
    writeLines(paste(unlist(debug_info), collapse = "\n"), debug_log_file)
    warning(paste("Errors occurred during conversion. Debug log written to", debug_log_file))
  }
  
  # Return success/failure
  return(!exists("error_occurred") || !error_occurred)
  # Connection will be closed by on.exit handler
}



# UI

ui <- dashboardPage(
  dashboardHeader(title = "SimaPro CSV Converter"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Home", tabName = "home", icon = icon("home")),
      menuItem("About", tabName = "about", icon = icon("info-circle")),
      menuItem("Instructions", tabName = "instructions", icon = icon("book"))
    )
  ),
  dashboardBody(
    tabItems(
      # Home tab
      tabItem(tabName = "home",
              fluidRow(
                box(
                  title = "Upload Excel File", width = 12, status = "primary",
                  fileInput("file", "Choose Excel file (.xlsx or .xls)",
                            accept = c(".xlsx", ".xls")),
                  actionButton("convert", "Convert to SimaPro CSV", 
                               icon = icon("file-export"), 
                               class = "btn-primary"),
                  downloadButton("downloadZip", "Download All CSVs as ZIP", class = "btn-success")
                )
              ),
              fluidRow(
                box(
                  title = "Processing Status", width = 12, status = "info",
                  verbatimTextOutput("processingStatus")
                )
              ),
              fluidRow(
                box(
                  title = "Excel Sheet Preview", width = 12, status = "warning",
                  selectInput("sheetSelect", "Select sheet to preview:", choices = NULL),
                  DT::dataTableOutput("excelPreview")
                )
              )
      ),
      
      # About tab
      tabItem(tabName = "about",
              box(
                title = "About SimaPro CSV Converter", width = 12, status = "info",
                h3("SimaPro CSV Converter"),
                p("This application converts Life Cycle Inventory (LCI) Excel files to SimaPro CSV format."),
                p("Based on Python scripts originally created by Massimo."),
                
                h4("What this application does:"),
                tags$ul(
                  tags$li("Converts multiple sheets at once (creates a .csv file for each sheet in the Excel file)"),
                  tags$li("Both database processes and all types of exchanges can be specified")
                ),
                
                h4("What it doesn't do:"),
                tags$ul(
                  tags$li("Specify sub-compartment of emission (e.g. 'high-population')"),
                  tags$li("Specify uncertainty"),
                  tags$li("Add comments")
                ),
                p("(All these parameters can be specified directly in the .csv file though)")
              )
      ),
      
      # Instructions tab
      tabItem(tabName = "instructions",
              box(
                title = "Instructions", width = 12, status = "success",
                h4("Procedure:"),
                tags$ol(
                  tags$li("Prepare the life cycle inventory in Excel (see format instructions below)"),
                  tags$li("Upload the Excel file using the file input on the Home tab"),
                  tags$li("Click 'Convert to SimaPro CSV' to process all sheets"),
                  tags$li("Download the ZIP file containing all CSV results"),
                  tags$li("From SimaPro, use Import > File and the following settings:",
                          tags$ul(
                            tags$li("File format: 'SimaPro CSV'"),
                            tags$li("Object link method: 'Try to link imported objects to existing objects first'"),
                            tags$li("CSV format separator: 'Tab'"),
                            tags$li("Other options: 'Replace existing processes...'")
                          ))
                ),
                
                h4("Excel File Format:"),
                tags$ul(
                  tags$li("Cells A1:D6 are fixed, do not insert rows or columns there"),
                  tags$li("Each column is a process of the foreground system, and is matched by an identically named row and in the same order"),
                  tags$li("Use exact LCI database process names under the foreground system"),
                  tags$li("Use 'Raw', 'Air', 'Water', 'Soil', 'Waste', 'Social', 'Economic' to indicate exchanges"),
                  tags$li("Use 'Wastetotreatment' to indicate database processes of the waste treatment category")
                ),
                
                h4("Example Excel Structure:"),
                tags$pre(
                  "Row 1: [blank] [blank] [blank] [blank] [Process Name 1] [Process Name 2] ...\n",
                  "Row 2: [blank] [blank] [blank] [blank] [Unit 1] [Unit 2] ...\n",
                  "Row 3: [blank] [blank] [blank] [blank] [Amount 1] [Amount 2] ...\n",
                  "Row 4: [blank] [blank] [blank] [blank] [blank] [blank] ...\n",
                  "Row 5: [blank] [blank] [blank] [blank] [Comment 1] [Comment 2] ...\n",
                  "Row 6: [blank] [blank] [blank] [blank] [Category 1] [Category 2] ...\n",
                  "Row 7+: [Type] [Name] [blank] [Unit] [Value 1] [Value 2] ..."
                )
              )
      )
    )
  )
)


# Server function

server <- function(input, output, session) {
  # Reactive values to store processed data
  values <- reactiveValues(
    sheets = NULL,
    csv_files = list(),
    processing_complete = FALSE,
    processing_log = "No file processed yet."
  )
  
  # Update sheet selection when a file is uploaded
  observeEvent(input$file, {
    req(input$file)
    
    # Get sheet names from the Excel file
    tryCatch({
      sheet_names <- excel_sheets(input$file$datapath)
      values$sheets <- sheet_names
      updateSelectInput(session, "sheetSelect", choices = sheet_names)
      
      # Auto-select first sheet if available
      if(length(sheet_names) > 0) {
        updateSelectInput(session, "sheetSelect", selected = sheet_names[1])
      }
      
      values$processing_log <- paste("Detected", length(sheet_names), "sheets in the uploaded file.")
    }, error = function(e) {
      values$processing_log <- paste("Error reading Excel file:", e$message)
    })
  })
  
  # Display preview of selected sheet
  output$excelPreview <- DT::renderDataTable({
    req(input$file, input$sheetSelect)
    
    # Read the selected sheet - suppress warnings about column names
    tryCatch({
      # Use a more robust method to read Excel sheets
      df <- suppressWarnings(read_excel(input$file$datapath, 
                                        sheet = input$sheetSelect, 
                                        col_names = FALSE,
                                        .name_repair = "minimal"))
      
      # Display the data
      DT::datatable(df, options = list(
        scrollX = TRUE,
        pageLength = 15,
        lengthMenu = c(5, 10, 15, 25, 50)
      ))
    }, error = function(e) {
      values$processing_log <- paste(values$processing_log, "\nError previewing sheet:", e$message)
      return(data.frame(Error = paste("Could not read sheet:", e$message)))
    })
  })
  
  # Process Excel sheets to SimaPro CSVs
  observeEvent(input$convert, {
    req(input$file)
    
    # Reset processing status and files
    values$processing_log <- "Starting conversion process...\n"
    values$csv_files <- list()
    values$processing_complete <- FALSE
    
    # Create a temporary directory for our CSV files
    temp_dir <- file.path(tempdir(), paste0("simapro_", format(Sys.time(), "%Y%m%d_%H%M%S")))
    dir.create(temp_dir, showWarnings = FALSE, recursive = TRUE)
    
    values$processing_log <- paste0(values$processing_log, "Created temporary directory: ", temp_dir, "\n")
    
    # Process each sheet
    for (sheet_name in values$sheets) {
      values$processing_log <- paste0(values$processing_log, "Processing sheet: ", sheet_name, "\n")
      
      tryCatch({
        # Read the Excel sheet with more robust settings
        df <- suppressWarnings(read_excel(input$file$datapath, 
                                          sheet = sheet_name, 
                                          col_names = FALSE,
                                          .name_repair = "minimal",
                                          guess_max = 10000))  # Increase sample size for type guessing
        
        values$processing_log <- paste0(values$processing_log, "  - Read sheet with ", nrow(df), " rows and ", ncol(df), " columns\n")
        
        # Create safe filename
        safe_sheet_name <- gsub("[^a-zA-Z0-9_.-]", "_", sheet_name)
        output_file <- file.path(temp_dir, paste0(safe_sheet_name, ".csv"))
        values$processing_log <- paste0(values$processing_log, "  - Output file: ", output_file, "\n")
        
        # Convert to SimaPro CSV
        result <- to_spcsv(df, output_file)
        
        if (result && file.exists(output_file)) {
          # Check file size to ensure it's not empty
          file_size <- file.info(output_file)$size
          values$processing_log <- paste0(values$processing_log, "  - Created CSV file (", file_size, " bytes)\n")
          
          if (file_size > 100) { # More than just headers
            values$csv_files[[sheet_name]] <- output_file
            values$processing_log <- paste0(values$processing_log, "  - Success!\n")
          } else {
            values$processing_log <- paste0(values$processing_log, "  - WARNING: File seems too small, might be incomplete\n")
          }
        } else {
          values$processing_log <- paste0(values$processing_log, "  - ERROR: File creation failed\n")
        }
      }, error = function(e) {
        values$processing_log <- paste0(values$processing_log, "  - ERROR: ", e$message, "\n")
      })
    }
    
    if (length(values$csv_files) > 0) {
      values$processing_complete <- TRUE
      values$processing_log <- paste0(values$processing_log,
                                      "\nConversion complete! ",
                                      length(values$csv_files),
                                      " CSV files created.\nClick 'Download All CSVs as ZIP' to download.\n")
    } else {
      values$processing_log <- paste0(values$processing_log,
                                      "\nConversion failed. No CSV files were created. Please check that your Excel sheets match the required format.\n")
    }
  })
  
  # Output processing status
  output$processingStatus <- renderText({
    return(values$processing_log)
  })
  

        
        # Download handler for ZIP file containing all CSVs
        output$downloadZip <- downloadHandler(
          filename = function() {
            paste0("SimaPro_CSV_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".zip")
          },
          content = function(file) {
            # If no files were generated, show a modal and return an empty file
            if (!values$processing_complete || length(values$csv_files) == 0) {
              showModal(modalDialog(
                title = "No Files Available",
                "Please convert Excel files to CSV first.",
                easyClose = TRUE,
                footer = modalButton("OK")
              ))
              file.create(file)
              return()
            }
            
            # List of CSVs to include
            files_to_zip <- unlist(values$csv_files)
            zip_dir <- unique(dirname(files_to_zip))[1]
            current_dir <- getwd()
            
            # Switch to the directory containing the CSVs before zipping
            setwd(zip_dir)
            on.exit(setwd(current_dir), add = TRUE)
            
            # Try zipping; on error, log and return an empty file
            tryCatch({
              #zip::zip(zipfile = file, files = basename(files_to_zip))
              utils::zip(zipfile = file, files = basename(files_to_zip))
            }, error = function(e) {
              values$processing_log <- paste0(
                values$processing_log,
                "\nError creating ZIP file: ", e$message
              )
              file.create(file)
            })
          },
          contentType = "application/zip"
        )
}


    


 shiny::shinyApp(ui, server)

