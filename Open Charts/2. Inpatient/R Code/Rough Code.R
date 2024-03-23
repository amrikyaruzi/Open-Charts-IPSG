rm(list = ls())

## Loading required packages - can use library calls individually as well
purrr::walk(c("tidyverse", "readxl", "here", "rlang", "janitor", "ggtext", "officer", "glue"),
            library, character.only = TRUE, warn.conflicts = FALSE)


## Month & Year to keep - REMEMBER TO CHANGE THESE TWO ACCORDINGLY!
what_months <- month.name[1:12]
what_year <- 2023



## Loading data, etc
data <- list.files(here("./Open Charts/2. Inpatient/Downloaded file"),
           pattern = ".xlsx$",
           full.names = TRUE) %>% set_names() %>%
  map_df(.x = ., .f = ~read_excel(.x), .id = "File")

data1 <- data %>% pivot_longer(cols = c(matches("[[:digit:]]")),
                      names_to = "Documents", values_to = "Response") %>%
  filter(!is.na(Response))

data1 <- data1 %>% mutate(Documents = case_when(
  
  str_detect(Documents, "2.2") ~ "2.2 - Nursing Re-Assessment (MEWS, PEWS & MEOWS) - N",
  str_detect(Documents, "4.4") ~ "4.4 - PRN Analgesic Orders : Appropriateness of Administration - N",
  str_detect(Documents, "3.3") ~ "3.3 - Patient care orders : Countersignature by Nurses (Nurse's part) - N",
  
  TRUE ~ Documents
  
))

data1 <- data1 %>% mutate(CONSULTANT = coalesce(CONSULTANT, DOCTOR)) %>%
  select(-c(ID, `Start time`, `Completion time`, Email, Name, `PATIENT NAMES`, DOCTOR))


data1 <- data1 %>% mutate(`PATIENT'S LOCATION (WARD)` = case_when(
  
  (is.na(`PATIENT'S LOCATION (WARD)`) & str_detect(File, "ICU")) ~ "ICU",
  (is.na(`PATIENT'S LOCATION (WARD)`) & str_detect(File, "CCU")) ~ "CCU",
  (is.na(`PATIENT'S LOCATION (WARD)`) & str_detect(File, "A&E")) ~ "A&E",
  
  TRUE ~ `PATIENT'S LOCATION (WARD)`)
  
  ) %>% select(-File)

data1 <- data1 %>% mutate(Documents = str_trim(Documents))

data1 <- data1 %>% mutate(SNN = str_extract(Documents, "\\d\\d*\\.*\\d*"),
                          TYPE = str_extract(Documents, pattern = "[PNM]$"),
                          Documents = str_replace_all(Documents,
                                             pattern = ".*-\\s(.+)\\s-.+",
                                             replacement = "\\1"),
                          Documents = gsub(pattern = ":", replacement = "-", x = Documents),
                          Documents = str_trim(Documents))

data1 <- data1 %>% separate_longer_delim(cols = Response, delim = ";")
data1 <- data1 %>% filter(Response != "")
data1 <- data1 %>% separate_wider_delim(cols = Response,
                                        delim = " - ",
                                        names = c("Values", "Responses"))

data1 <- data1 %>% relocate(SNN, Documents, TYPE, .after = `PATIENT'S LOCATION (WARD)`)


data1 <- data1 %>% mutate(Values = case_when(
  
  Values == "TIMELINESS" ~ "TIMELY",
  Values %in% c("LEGIBILITY", "LEGBILITY") ~ "LEGIBLE",
  Values == "COMPLETENESS" ~ "COMPLETE",
  Values == "ACCURACY" ~ "ACCURATE",
  
  TRUE ~ Values
  
  ))

data1 <- data1 %>% mutate(across(c(Year, SNN), ~as.numeric(.)))

data1 <- data1 %>% mutate(Month = replace_na(Month, "February"),
                          Year = replace_na(Year, 2023))

data1 <- data1 %>% filter(Month != "January" & Year == 2023) #Removing the January 2023 entry(ies) entered erroneously

data1 <- data1 %>% mutate(Responses = na_if(Responses, "NA")) %>% filter(!is.na(Responses))
data1 <- data1 %>% mutate(Responses = case_when(
  
  Responses %in% c("Y", "M", "D") ~ "Y",
  Responses %in% c("N", "NM", "ND") ~ "N",
  TRUE ~ Responses
  
  ))


## Checking for duplicates across all columns
# data1 %>% group_by(`MR Number`, Month, Year, CONSULTANT, `PATIENT'S LOCATION (WARD)`, Documents,
#                    TYPE, Values, Responses) %>%
#   mutate(Duplicate = n()>1) %>%
#   filter(Duplicate == TRUE) %>%
#   arrange(`MR Number`, Month, Year, CONSULTANT, `PATIENT'S LOCATION (WARD)`, Documents, TYPE,
#           Values, Responses)


# Keeping only distinct values
data1 <- data1 %>% distinct(across(.cols = c(`MR Number`, Month, Year, RMO, NURSE, CONSULTANT,
                                             `PATIENT'S LOCATION (WARD)`, SNN, Documents, TYPE,
                                             Values)), .keep_all = TRUE)

data1 <- data1 %>% pivot_wider(names_from = Values, values_from = Responses)

data1 <- data1 %>% filter(!is.na(DOCUMENTED))

## To get a list of unique values for DOCUMENTED:ACCURATE columns:
# data1 %>% select(DOCUMENTED:last_col()) %>%
#   mutate(across(.cols = everything(), ~replace_na(., "NA"))) %>%
#   map_df(.f = ~str_c(unique(.x), collapse = ", ")) %>%
#   pivot_longer(cols = everything(), names_to = "column", values_to = "unique_entries")


# Adding Departments based on Ward of Admission and/or Consultant
data1 <- data1 %>% rename(Ward = `PATIENT'S LOCATION (WARD)`) %>%
  mutate(CONSULTANT = str_trim(CONSULTANT))
  

data1 <- data1 %>% mutate(Department = case_when(
  
  Ward == "A&E" ~ "A&E",
  
  CONSULTANT %in% c('Abeid, Muzdalifat', 'Abeid, Muzdalfat', 'Bakari, Rahma', 'Rahma Muhammad Bakari', 'Damas, Wilson', 'Jaiswal, Shweta', 'Kaguta, Munawar', 'Kidanto, Hussein', 'Mdachi, Ernest', 'Mgonja, Miriam', 'Moshi, Lynn','Muzo, Jane', 'Ngarina, Matilda', 'Mujumali, Nyasinde') ~ 'Maternity',
  CONSULTANT %in% c('Adatia, Aleesha', 'Adebayo, Philip', 'Amirali, Muzhar', 'Mazhar Amirali', 'Kimberly Annette Craven', 'Lubuva, Neema', 'Muhoka,Peter', 'Virani, Akbarali', 'Bakshi, Fatima', 'Bakshi, Fatma', 'Bakshi,  Fatma', 'Bapumia, Mustafa', 'Craven, Kimberly', 'Chuwa, Harrison', 'Epafra, Emmanuel', 'Shayo, Grace', 'Foi, Andrew', 'Hameed, Kamran', 'Hussain, Bilal', 'Jamal, Nasiruddin', 'Khuzeima, Kaderbhai', 'Khuzeima, Kanbhai', 'Lyuu, Tuzo','Makakala, Mandela', 'Masolwa, Deodatus', 'Mbatina, Consolata', 'Mbithe, Hanifa', 'Mbonea Salehe Yonazi', 'Mwambene, Kissah', 'Nyagori, Harun', 'Mvungi, Robert', 'Robert, Mvungi', 'Rofaeil, Bishoy', 'Sangeti, Saningo', 'Shamji, Munaf', 'Somji, Samina', 'Tungaraza, Kheri', 'Wambura, Casmir', 'Khan, Zahra') ~ 'Medicine',
  CONSULTANT %in% c('Abdallah, Yaser', 'Bulimba, Maria', 'Ebrahim. Mohamedraza', 'Ibrahim, Mohamedraza', 'Hajaj, Salum', 'Kija. Edward', 'Sonal Patel', 'Kubhoja, Sulende', 'Mbise, Roger Lewis', 'Mwamanenge, Naomi', 'Nkya, Deogratis', 'Noorani, Mariam', 'Patel, Sonal', 'Walli, Nahida') ~ 'Paediatrics',
  CONSULTANT %in% c('Ali, Athar', 'Athar, Ali', 'Assey, Anthony', 'Bajsar, Ally', 'Soomro, Hussam', 'Clement, Mughisha', 'Clement, Mugisha', 'Dinda, Julius', 'Nyamuryekung, Masawa', 'Mushi, Fransia', 'Joseph, Alex', 'Khamis Omar Khamis', 'Khan, Muhammad', 'Kimu, Njiku', 'Kumar, Rajeev', 'Laurent, Lemery', 'Liyombo, Edwin', 'Lyimo, Elias', 'Mashamba, Victor', 'Mavura, Maurice', 'Mawalla, Isaac', 'Mcharo, Bryson', 'Moshi, Ndeserua', 'Mrema, Edwin', 'Mrita, Felix', 'Mtanda, Tawakali', 'Mugabo, Rajab', 'Mwanga, Ally', 'Mwansasu, Christopher', 'Mbwambo, John', 'Ngerageza, Japhet', 'Ngiloi, Petronilla', 'Ngiloi, Petronila', 'Njau, Aidan', 'Nyaluke, Paul', 'Nungu, Samwel', 'Padhani, Dilawar', 'Patel, Miten', 'Rajput, Irfan', 'Richard, Enica', 'Ringo, Yona', 'Sianga, William', 'Walter Charles Mbando (ENT)', 'Zehri, AliAkbar', 'Zehri, Ali Akbar') ~ 'Surgery',
  TRUE ~ ''
 
  ))

data1 <- data1 %>%  relocate(Department, .after = Ward)

## To get Ward-department pairs
# data1 %>% distinct(across(c(Ward, Department)))

data1 <- data1 %>% filter((DOCUMENTED == "N") |
                            (DOCUMENTED == "Y" &
                               !is.na(TIMELY) &
                               !is.na(LEGIBLE) &
                               !is.na(COMPLETE))) #Skipped !is.na(ACCURATE)

data1 <- data1 %>% mutate(MET = case_when(
  DOCUMENTED == "Y" & TIMELY == "Y" & LEGIBLE == "Y" & COMPLETE == "Y" &
    (ACCURATE == "Y" | is.na(ACCURATE)) ~ "Y",
  
  TRUE ~ "N"
  
))


# Writing Consultant names as FirstName LastName
data1 <- data1 %>% mutate(CONSULTANT = case_when(Department != "A&E" ~ gsub(",", "", CONSULTANT) %>%
                                                   str_replace_all(string = .,
                                                                   pattern = "(\\w+)(\\s)\\s*(\\w+)",
                                                                   replacement = "\\3\\2\\1"),
                                                 TRUE ~ CONSULTANT))


data1 <- data1 %>% mutate(Quarter = case_when(Month %in% c("January", "February", "March") ~ "Q1",
                                              Month %in% c("April", "May", "June") ~ "Q2",
                                              Month %in% c("July", "August", "September") ~ "Q3",
                                              Month %in% c("October", "November", "December") ~ "Q4")) %>%
  relocate(Quarter, .after = Month)

# Keeping only the chosen month & year
data1 <- data1 %>% filter(Month %in% what_months & Year %in% what_year)

data1 <- data1 %>% mutate(Month = substr(Month, start = 1, stop = 3))


# Identifying the current quarter (highest quarter)
thequarter <- if("Q4" %in% unique(data1$Quarter)){
  
  "Q4"
  
} else if("Q3" %in% unique(data1$Quarter)) {
  
  "Q3"
  
} else if("Q2" %in% unique(data1$Quarter)) {
  
  "Q2"
  
} else {
  
  "Q1"
}


#cShortening document names

data1 <- data1 %>% mutate(Documents = str_trim(Documents),
                          Documents = case_when(

Documents == "Physician Initial Assessment All components" ~ "Physician Initial Assessment",
Documents == "Nursing Initial Assessment All components" ~ "Nursing Initial Assessment",
Documents == "Appropriate pain assessment, intervention and reassessment" ~ "Pain Assessment & Reassessment",
Documents == "Daily Round Notes by Physicians (ward/interdisciplinary ICU round)" ~ "Round Notes (SBAR)",
Documents == "Fall Risk Assessment, Intervention and Reassessment" ~ "Fall Risk Assessment",
Documents == "Multidisciplinary Patient and family education Form from arrival to discharge" ~ "Multidisciplinary Patient and Family Education",
Documents == "Patient care orders on uniform location (Doctor's part)" ~ "Physician Orders",
Documents == "Daily CPOE orders and Nursing signature for administration" ~ "CPOE orders - Nurse's Signature",
Documents == "Informed Consent Form (all procedures requiring consent i.e. anesthesia, sedation, BT, chemo, dialys" ~ "Informed Consent",
Documents == "Check the Discharge summary and DNR forms for any abbreviations" ~ "Discharge summaries without Abbreviations",
Documents == "Check Consent forms for any abbreviations" ~ "Consent forms without Abbreviations",
Documents == "Blood Transfusion Monitoring Form" ~ "Blood Transfusion Monitoring",
Documents == "Discharge/Transfer/Referral Summary at exit" ~ "Discharge summary",
Documents == "Pre-Operative Anesthesia/Sedation Assessment" ~ "Pre-Anaesthesia/Sedation Assessment",
Documents == "Intraoperative form for Site marking, sign in, time out and sign out" ~ "Site Marking, Sign in, Time out & Sign out",
Documents == "Surgical Procedure/ Operative notes" ~ "Procedure/ Operative notes",
Documents == "Pre-Operative (Anaesthesia) Induction Assessment" ~ "Pre-Induction Assessment",
Documents == "Post Anesthesia Recovery Form" ~ "Post-Anaesthesia Recovery",
Documents == "Intraoperative Notes of Anesthesia" ~ "Intraoperative Notes of Anesthesia",
Documents == "Physiotherapy Assessment All components" ~ "Physiotherapy Assessment",
Documents == "Physiotherapy Fall Risk Assessment" ~ "Physiotherapy Fall Risk Assessment",
Documents == "Nutritional Assessment All components" ~ "Nutritional Assessment",
Documents == "STAT medications administration" ~ "STAT medications",
Documents == "Nursing Care Plan" ~ "Nursing Care Plan",
Documents == "Blood Transfusion Order" ~ "Blood Transfusion Order",
Documents == "Daily CPOE orders and Physician signature for administration" ~ "CPOE orders - Doctor's Signature",
Documents == "Nursing Re-Assessment and Appropriate Fall reassessment" ~ "MEWS, NEWS, PEWS & MEOWS",
Documents == "Check Consent, DNR and discharge summary for any abbreviations" ~ "Consent forms without Abbreviations",
Documents == "ICU/CCU discharge/Transfer criteria is documented before discharging/transfer patient" ~ "ICU/CCU discharge/Transfer criteria",
Documents == "Pre & Post cardiac catheterization assessment and orders documented as required" ~ "Pre & Post cardiac catheterization assessment and orders",
Documents == "Procedural sedation assessment and monitoring pre, intra and post procedure documented as required." ~ "Procedural sedation assessment and monitoring",

Documents %in% c("Nursing Re-Assessment (MEWS, PEWS & MEOWS)", "Nursing Re-Assessment and Appropriate Fall reassessment") ~ "MEWS, NEWS, PEWS & MEOWS",
Documents %in% c("Patient care orders - Countersignature by Nurses (Nurse's part)", "Patient care orders on uniform location (Nurse's part)") ~ "Physician Orders - Countersigning",

str_detect(Documents, "work sheet documented as required") ~ "CathLab Worksheet",

TRUE ~ Documents))




data1 <- data1 %>% mutate(Department = case_when(
  
  Department == "Maternity" ~ "Obstetrics & Gynaecology",
  TRUE ~ Department
  
  ))


data1 <- data1 %>% mutate(across(c(Ward, Department),
                        .fns = ~(.x = case_when(.x == "Maternity" ~ "Obstetrics & Gynaecology",
                                                .x == "Maternity Ward" ~ "Obstetrics & Gynaecology",
                                                .x == "Paediatric Ward" ~ "Paediatrics",
                                                .x == "Surgical Ward" ~ "Surgery",
                                                .x == "Medical Ward" ~ "Medicine",
                                                TRUE ~ .x))))



# Assigning responsible department based on department, ward and type

data1 <- data1 %>% mutate(Responsible_Department = case_when(
  (Department == "Medicine" & TYPE %in% c("P", "M")) |
    (Ward == "Medicine" & TYPE == "N") ~ "Medicine",
  
  (Department == "Obstetrics & Gynaecology" & TYPE %in% c("P", "M")) |
    (Ward %in% c("Obstetrics & Gynaecology", "NHDU") & TYPE == "N") ~ "Obstetrics & Gynaecology",
  
  (Department %in% c("Paediatrics", "NHDU") & TYPE %in% c("P", "M")) |
    (Ward == "Paediatrics" & TYPE == "N") ~ "Paediatrics",
  
  (Department == "Surgery" & TYPE %in% c("P", "M")) |
    (Ward == "Surgery" & TYPE == "N") ~ "Surgery",
  
  
  (Ward %in% c("ICU", "HDU", "Paediatric HDU") & TYPE == "N") ~ "ICU",
  
  (Ward == "CCU" & TYPE == "N") ~ "CCU",
  
  (Ward == "Pavilion" & TYPE == "N") ~ "Pavilion",
  
  Department == "A&E" ~ "A&E",
  
  TRUE ~ Ward
    
    )) %>%
  
  relocate(Responsible_Department, .after = Department)



# Functions for later use
## Comparative bar graphs - all departments

bar_graph_all <- function(column){
  
  
  col_quosure <- ensym(column) #rlang::enexpr(column)
  
  ## Creating dataframe with required variables
  intermediate <- data1 %>% filter(!Responsible_Department %in%
                                     c("Pavilion", "ICU", "CCU", "NHDU", "HDU", "Paediatric HDU"))
    group_by(Month, Responsible_Department, !!col_quosure) %>%
    filter(!!col_quosure %in% c("Y", "N")) %>%
    summarise(Count = n()) %>%
    mutate(Percentage = (100 * Count)/sum(Count)) %>% ungroup()
  
  # Setting benchmark based on selected variable (column)
  benchmark = ifelse(ensym(col_quosure) == "DOCUMENTED", 100, 90)
  
  # Graph titles based on selected variable (column)
  title = case_when(ensym(col_quosure) == "DOCUMENTED" ~ "Availability of Required Documents by Department",
                    ensym(col_quosure) == "TIMELY" ~ "Timeliness by Department",
                    ensym(col_quosure) == "LEGIBLE" ~ "Legibility by Department",
                    ensym(col_quosure) == "COMPLETE" ~ "Completeness by Department",
                    ensym(col_quosure) == "ACCURATE" ~ "Accuracy by Department",
                    ensym(col_quosure) == "MET" ~ "Overall Compliance by Department")
  
  # Further wrangling
  intermediate <- intermediate %>%
    filter(!!col_quosure == "N" & Percentage == 100) %>%
    mutate(!!col_quosure := "Y") %>%
    mutate(Count = 0) %>%
    mutate(Percentage = 0) %>%
    rbind(., intermediate %>% filter(!!col_quosure == "Y")) %>%
    arrange(desc(Percentage)) %>%
    mutate_if(is.numeric, round_half_up, 1) %>% 
    
    
    mutate(Category = case_when(

      Percentage < 50 ~ "Below 50",
      (Percentage >= 50 & Percentage < benchmark) ~ "Above 50 but below benchmark",
      Percentage >= benchmark ~ "At or above benchmark"),


          Category = factor(Category, levels = c("At or above benchmark",
                                                 "Above 50 but below benchmark",
                                                  "Below 50"))) %>%
    
    # Drawing ggplot2 graphs (bar graphs)
    ggplot(., aes(x = Month, y = Percentage, fill = Category)) +
    geom_bar(stat = "identity") +
    facet_wrap(~Responsible_Department, nrow = 2) +
    labs(title = paste0("\n", title, "\n"), x = "\nMonth", y = "Compliance (%)\n") +
    theme(legend.position = "bottom",
          legend.title = element_blank(),
          plot.title.position = "plot",
          plot.title = element_text(hjust = 0.5),
          axis.line = element_line(),
          panel.background = element_rect(fill = "gray96"), #"gray96"
          axis.text.x = element_text(angle = 0, hjust = 0.5, vjust = 0.8),
          strip.text = element_text(colour = "black"#, face = "italic"
                                    ),
          strip.background = element_rect(fill = "#B2BEB5")) +
    scale_fill_manual(values = c("Below 50" = "red",
                                 "Above 50 but below benchmark" = "orange",
                                 "At or above benchmark" = "#00b159")) +
    
    guides(fill = guide_legend(title = "Key:")) +
    
    scale_y_continuous(breaks = seq(0, 100, 10))
  
  print(intermediate) %>% 
  assign(value = .,
         x = paste0(tolower(col_quosure),
                    "_bar"),
         envir = .GlobalEnv)
  
}


walk(.x = c("DOCUMENTED", "TIMELY", "LEGIBLE", "COMPLETE", "ACCURATE", "MET"),
     .f = bar_graph_all)


## All departmental charts
departmental_charts <- function(department){
  
  data2 <- data1 %>% filter(Responsible_Department == department)
  
  
  ## Control charts
  control_charts <- function(column){
  
  col_quosure <- ensym(column)
  
  ### Creating dataframe with required variables
  intermediate <- data2 %>%
    select(Month, Responsible_Department, TYPE, !!col_quosure) %>%
    group_by(Month, Responsible_Department, !!col_quosure) %>%
    filter(!!col_quosure %in% c("Y", "N")) %>%
    summarise(Count = n()) %>%
    mutate(Percentage = (100 * Count)/sum(Count)) %>% ungroup()
  
    
  ### Setting benchmark based on selected variable (column)
  benchmark = ifelse(ensym(col_quosure) == "DOCUMENTED", 100, 90)
  
  ### Graph titles based on selected variable (column)
  title = case_when(ensym(col_quosure) == "DOCUMENTED" ~ "Availability of Required Documents",
                    ensym(col_quosure) == "TIMELY" ~ "Timeliness",
                    ensym(col_quosure) == "LEGIBLE" ~ "Legibility",
                    ensym(col_quosure) == "COMPLETE" ~ "Completeness",
                    ensym(col_quosure) == "ACCURATE" ~ "Accuracy",
                    ensym(col_quosure) == "MET" ~ "Overall Compliance")
  
  ### Further wrangling
  intermediate <- intermediate %>%
    filter(!!col_quosure == "N" & Percentage == 100) %>%
    mutate(!!col_quosure := "Y") %>%
    mutate(Count = 0) %>%
    mutate(Percentage = 0) %>%
    rbind(., intermediate %>% filter(!!col_quosure == "Y")) %>%
    arrange(desc(Percentage)) %>%
    mutate_if(is.numeric, round_half_up, 1) %>%
    
    mutate(Mean = mean(Percentage)) %>% 
    mutate(UCL = Mean + 2*(sd(Percentage)),
           LCL = Mean - 2*(sd(Percentage)),
           Benchmark = benchmark) %>% mutate_if(is.numeric, round_half_up, 1)
    
    ### Drawing ggplot2 graphs (line graphs)
    graph <- ggplot(data = intermediate, aes(x = Month, y = Percentage, group = 1)) +
      geom_point(shape = 21, color = "black", fill= "#69b3a2", size = 2) +
      geom_line() + geom_hline(aes(yintercept = UCL, colour = "UCL (2σ)")) +
      geom_hline(aes(yintercept = Benchmark, colour = "Benchmark")) +
      geom_hline(aes(yintercept = Mean, colour = "Mean")) +
      geom_hline(aes(yintercept = LCL, colour = "LCL (2σ)")) +
      scale_y_continuous(breaks = seq(-20,200,10)) +
      
      geom_text(label = intermediate$Percentage,
                hjust = -0.5,
                size = if_else(n_distinct(intermediate$Percentage) >= 9, 2.6, 2.8),
                position = position_dodge(width = 1)) +
      
      theme(plot.title.position = "plot",
            plot.title = element_text(hjust = 0.5),
            axis.line = element_line(),
            panel.background = element_rect(fill = "white"), #gray98
            legend.background = element_rect(),
            legend.title = element_blank(),
            legend.position = "bottom") +
      
      labs(title = paste0("\n", department, " - ", title, "\n"), x = "\nMonth", y = "Compliance (%)\n")
    
      
  print(graph) %>% 
    assign(value = .,
           x = paste0(tolower(col_quosure),
                      "_control_chart"),
           envir = .GlobalEnv)
  
  }
  
  walk(.x = c("DOCUMENTED", "TIMELY", "LEGIBLE", "COMPLETE", "ACCURATE", "MET"), #Control charts walk
       .f = control_charts)
  
  
  
  ## Bar graph per form
  form_charts <- function(column){
    
    col_quosure <- ensym(column)
    
    ## Creating dataframe with required variables
    intermediate <- data2 %>% filter(Quarter == thequarter) %>%                       ###Remember this
      select(Quarter, Year, Documents, TYPE, !!col_quosure) %>%
      group_by(Quarter, Year, Documents, TYPE, !!col_quosure) %>%
      summarise(Count = n()) %>% filter(!!col_quosure %in% c("Y", "N")) %>%
      mutate(Percentage = (100*Count)/sum(Count)) %>% ungroup()
    
    # Setting benchmark based on selected variable (column)
    benchmark = ifelse(ensym(col_quosure) == "DOCUMENTED", 100, 90)
    
    # Graph titles based on selected variable (column)
    title = case_when(ensym(col_quosure) == "DOCUMENTED" ~ "Availability of Required Documents by form",
                      ensym(col_quosure) == "TIMELY" ~ "Timeliness by form",
                      ensym(col_quosure) == "LEGIBLE" ~ "Legibility by form",
                      ensym(col_quosure) == "COMPLETE" ~ "Completeness by form",
                      ensym(col_quosure) == "ACCURATE" ~ "Accuracy by form",
                      ensym(col_quosure) == "MET" ~ "Overall Compliance by form")
    
    # Further wrangling
    intermediate <- intermediate %>%
      filter(!!col_quosure == "N" & Percentage == 100) %>%
      mutate(!!col_quosure := "Y") %>%
      mutate(Count = 0) %>%
      mutate(Percentage = 0) %>%
      rbind(., intermediate %>% filter(!!col_quosure == "Y")) %>%
      arrange(desc(Percentage)) %>%
      mutate_if(is.numeric, round_half_up, 1) %>% 
      
      
      mutate(Category = case_when(
        
        Percentage < 50 ~ "Below 50",
        (Percentage >= 50 & Percentage < benchmark) ~ "Above 50 but below benchmark",
        Percentage >= benchmark ~ "At or above benchmark"),
        
        
        Category = factor(Category, levels = c("At or above benchmark",
                                               "Above 50 but below benchmark",
                                               "Below 50"))) %>%
      inner_join(., {
        
        data2 %>% filter(Quarter == thequarter) %>%                       ###Remember this
          select(Quarter, Year, Documents, TYPE, !!col_quosure) %>%
          filter(!!col_quosure %in% c("Y", "N")) %>%
          group_by(Quarter, Year, Documents, TYPE) %>%
          summarize(Total = n()) %>% ungroup()
          
        }) %>%
      
      relocate(Total, .after = Count) %>%
      arrange(desc(Percentage)) %>% mutate_if(is.numeric, round_half_up, 1) %>%
      mutate(label = glue::glue("*{Percentage}%*  (<sup>{Count}</sup>&frasl;<sub>{Total}</sub>)"),
             Quarter = paste(Quarter, Year)) %>%
      
      # label above is for geom_richtext label
      
      mutate(Documents = reorder(Documents, Percentage))
    
    
    
      # Drawing ggplot2 graphs (bar graphs)
      graph <- ggplot(intermediate,
                      aes(x = Documents,
                          y = Percentage, fill = Category,
                          group = Category,
                          label = label)) +
      geom_col() +
      coord_flip() + 
      labs(title = paste0("\n", department, " - ", title, "\n"),
           x = "\nForm\n", y = "\nCompliance (%)\n") +
      
      theme(legend.position = "bottom", legend.justification = c(0.4, 0),
            #legend.title = element_blank(),
            plot.title.position = "plot",
            plot.title = element_text(family = "Helvetica",
                                      size = 14,
                                      hjust = 0.5),
            axis.line = element_line(),
            panel.background = element_rect(fill = "white"), #gray96
            axis.text.x = element_text(hjust = 0.5, vjust = 0.8),
            
            strip.text = element_text(colour = "black"#, face = "italic"
            ),
            strip.background = element_rect(fill = "#B2BEB5")) +
      
        scale_fill_manual(values = c("At or above benchmark" = "#00b159",
                                   "Above 50 but below benchmark" = "orange",
                                   "Below 50" = "red")) +
        
        guides(fill = guide_legend(title = "Key:")) +
        
      scale_y_continuous(
        expand = c(0, 0),
        limits = c(0, 110),
        breaks = scales::breaks_pretty()
                         ) +
        
        geom_richtext(fill = NA, label.color = NA,
                      hjust = if_else(intermediate$Percentage == 0, 0, 1), #1 originally, moved because 0's start at y axis line then 0.9
                      nudge_y = if_else(intermediate$Percentage == 0, 1, 0), #0.5
                      vjust = 0.5, #0.5
                      size = 2.7) +
      
      facet_wrap(~Quarter)
      
      
    print(graph) %>% 
      assign(value = .,
             x = paste0(tolower(col_quosure),
                        "_form_chart"),
             envir = .GlobalEnv)
    
    
  }
  
  walk(.x = c("DOCUMENTED", "TIMELY", "LEGIBLE", "COMPLETE", "ACCURATE", "MET"), #Form bar chart walk
       .f = form_charts)
  
  
  
  ## Saving the graphs
  save_plots <- function(placeholder){
  
    read_pptx(here("./Open Charts/2. Inpatient/R Code/Template.pptx")) %>%
    ph_with(value = block_list(fpar(ftext(glue("{placeholder} - Open Charts Documentation Compliance for {thequarter} {what_year}"),
                                          prop = fp_text(font.size = 18, font.family = "Helvetica")),
                                    fp_p = fp_par(text.align = "center"))),
            location = ph_location_type(type = "subTitle")) %>% 
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
    ph_with(value = documented_control_chart,
            location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = documented_form_chart,
              location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = documented_bar,
              location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = timely_control_chart,
              location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = timely_form_chart,
              location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = timely_bar,
              location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = legible_control_chart,
              location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = legible_form_chart,
              location = ph_location_fullsize()) %>%
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = legible_bar,
              location = ph_location_fullsize()) %>%  
    
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = complete_control_chart,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = complete_form_chart,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = complete_bar,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = accurate_control_chart,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = accurate_form_chart,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = accurate_bar,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = met_control_chart,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = met_form_chart,
              location = ph_location_fullsize()) %>%
      
    add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = met_bar,
              location = ph_location_fullsize()) %>%  
      
print(glue(here("./Open Charts/2. Inpatient/Output/{what_year}/{thequarter}/{department}/{placeholder} - Open Charts Results.pptx")))
  
    }
  
  walk(.x = department, .f = save_plots) #Walk for saving plots for each department
  
}


# Running trigger for all departments - Setting directory & all!!
# Removing Pavilion, ICU, CCU, NHDU, HDU, Paediatric HDU as they will be run independently below
#departments <- departments[!departments %in% c("Pavilion", "ICU", "CCU", "NHDU", "HDU", "Paediatric HDU")]

departments <- unique(data1$Responsible_Department) %>%
  .[!. %in% c("Pavilion", "ICU", "CCU", "NHDU", "HDU", "Paediatric HDU")]



## Setting where to export pptx (directory)

create_directory <- function(department){
  
  walk(c(glue(here("Open Charts/2. Inpatient/Output", {what_year})),
         glue(here("Open Charts/2. Inpatient/Output", {what_year}, {thequarter})),
         glue(here("Open Charts/2. Inpatient/Output", {what_year}, {thequarter}, {department}))#,
         #glue(here("Open Charts/2. Inpatient/Output", {what_year}, {thequarter}, {department}, "Consultants"))
         ),
       ~dir.create(path = ., showWarnings = FALSE))
  
  }

walk(.x = unique(data1$Responsible_Department), # this is inclusive of departments like even Pavilion
     .f = create_directory)


walk(.x = departments,
     .f = departmental_charts)

rm(list = setdiff(ls(),
                  c("data", "data1", "what_year", "what_months", "thequarter")))



### Summaries for Pavilion, ICU, CCU, NHDU, HDU and Paediatric HDU

departmental_charts_others <- function(other_department){
  
  
  data2 <- data1 %>% filter(Ward == other_department)
  
  
  ## Control charts
  control_charts <- function(column){
    
    col_quosure <- ensym(column)
    
    ### Creating dataframe with required variables
    intermediate <- data2 %>%
      select(Month, Ward, TYPE, !!col_quosure) %>%
      group_by(Month, Ward, !!col_quosure) %>%
      filter(!!col_quosure %in% c("Y", "N")) %>%
      summarise(Count = n()) %>%
      mutate(Percentage = (100 * Count)/sum(Count)) %>% ungroup()
    
    
    ### Setting benchmark based on selected variable (column)
    benchmark = ifelse(ensym(col_quosure) == "DOCUMENTED", 100, 90)
    
    ### Graph titles based on selected variable (column)
    title = case_when(ensym(col_quosure) == "DOCUMENTED" ~ "Availability of Required Documents",
                      ensym(col_quosure) == "TIMELY" ~ "Timeliness",
                      ensym(col_quosure) == "LEGIBLE" ~ "Legibility",
                      ensym(col_quosure) == "COMPLETE" ~ "Completeness",
                      ensym(col_quosure) == "ACCURATE" ~ "Accuracy",
                      ensym(col_quosure) == "MET" ~ "Overall Compliance")
    
    ### Further wrangling
    intermediate <- intermediate %>%
      filter(!!col_quosure == "N" & Percentage == 100) %>%
      mutate(!!col_quosure := "Y") %>%
      mutate(Count = 0) %>%
      mutate(Percentage = 0) %>%
      rbind(., intermediate %>% filter(!!col_quosure == "Y")) %>%
      arrange(desc(Percentage)) %>%
      mutate_if(is.numeric, round_half_up, 1) %>%
      
      mutate(Mean = mean(Percentage)) %>% 
      mutate(UCL = Mean + 2*(sd(Percentage)),
             LCL = Mean - 2*(sd(Percentage)),
             Benchmark = benchmark) %>% mutate_if(is.numeric, round_half_up, 1)
    
    ### Drawing ggplot2 graphs (line graphs)
    graph <- ggplot(data = intermediate, aes(x = Month, y = Percentage, group = 1)) +
      geom_point(shape = 21, color = "black", fill= "#69b3a2", size = 2) +
      geom_line() + geom_hline(aes(yintercept = UCL, colour = "UCL (2σ)")) +
      geom_hline(aes(yintercept = Benchmark, colour = "Benchmark")) +
      geom_hline(aes(yintercept = Mean, colour = "Mean")) +
      geom_hline(aes(yintercept = LCL, colour = "LCL (2σ)")) +
      scale_y_continuous(breaks = seq(-20,200,10)) +
      
      geom_text(label = intermediate$Percentage,
                hjust = -0.5,
                size = if_else(n_distinct(intermediate$Percentage) >= 9, 2.6, 2.8),
                position = position_dodge(width = 1)) +
      
      theme(plot.title.position = "plot",
            plot.title = element_text(hjust = 0.5),
            axis.line = element_line(),
            panel.background = element_rect(fill = "white"), #gray98
            legend.background = element_rect(),
            legend.title = element_blank(),
            legend.position = "bottom") +
      
      labs(title = paste0("\n", other_department, " - ", title, "\n"),
           x = "\nMonth", y = "Compliance (%)\n")
    
    
    print(graph) %>% 
      assign(value = .,
             x = paste0(tolower(col_quosure),
                        "_control_chart"),
             envir = .GlobalEnv)
    
  }
  
  walk(.x = c("DOCUMENTED", "TIMELY", "LEGIBLE", "COMPLETE", "ACCURATE", "MET"), #Control charts walk
       .f = control_charts)
  
  
  
  ## Bar graph per form
  form_charts <- function(column){
    
    col_quosure <- ensym(column)
    
    ## Creating dataframe with required variables
    intermediate <- data2 %>% filter(Quarter == thequarter) %>%                       ###Remember this
      select(Quarter, Year, Documents, TYPE, !!col_quosure) %>%
      group_by(Quarter, Year, Documents, TYPE, !!col_quosure) %>%
      summarise(Count = n()) %>% filter(!!col_quosure %in% c("Y", "N")) %>%
      mutate(Percentage = (100*Count)/sum(Count)) %>% ungroup()
    
    # Setting benchmark based on selected variable (column)
    benchmark = ifelse(ensym(col_quosure) == "DOCUMENTED", 100, 90)
    
    # Graph titles based on selected variable (column)
    title = case_when(ensym(col_quosure) == "DOCUMENTED" ~ "Availability of Required Documents by form",
                      ensym(col_quosure) == "TIMELY" ~ "Timeliness by form",
                      ensym(col_quosure) == "LEGIBLE" ~ "Legibility by form",
                      ensym(col_quosure) == "COMPLETE" ~ "Completeness by form",
                      ensym(col_quosure) == "ACCURATE" ~ "Accuracy by form",
                      ensym(col_quosure) == "MET" ~ "Overall Compliance by form")
    
    # Further wrangling
    intermediate <- intermediate %>%
      filter(!!col_quosure == "N" & Percentage == 100) %>%
      mutate(!!col_quosure := "Y") %>%
      mutate(Count = 0) %>%
      mutate(Percentage = 0) %>%
      rbind(., intermediate %>% filter(!!col_quosure == "Y")) %>%
      arrange(desc(Percentage)) %>%
      mutate_if(is.numeric, round_half_up, 1) %>% 
      
      
      mutate(Category = case_when(
        
        Percentage < 50 ~ "Below 50",
        (Percentage >= 50 & Percentage < benchmark) ~ "Above 50 but below benchmark",
        Percentage >= benchmark ~ "At or above benchmark"),
        
        
        Category = factor(Category, levels = c("At or above benchmark",
                                               "Above 50 but below benchmark",
                                               "Below 50"))) %>%
      inner_join(., {
        
        data2 %>% filter(Quarter == thequarter) %>%                       ###Remember this
          select(Quarter, Year, Documents, TYPE, !!col_quosure) %>%
          filter(!!col_quosure %in% c("Y", "N")) %>%
          group_by(Quarter, Year, Documents, TYPE) %>%
          summarize(Total = n()) %>% ungroup()
        
      }) %>%
      
      relocate(Total, .after = Count) %>%
      arrange(desc(Percentage)) %>% mutate_if(is.numeric, round_half_up, 1) %>%
      mutate(label = glue::glue("*{Percentage}%*  (<sup>{Count}</sup>&frasl;<sub>{Total}</sub>)"),
             Quarter = paste(Quarter, Year)) %>%
      
      # label above is for geom_richtext label
      
      mutate(Documents = reorder(Documents, Percentage))
    
    
    
    # Drawing ggplot2 graphs (bar graphs)
    graph <- ggplot(intermediate,
                    aes(x = Documents,
                        y = Percentage, fill = Category,
                        group = Category,
                        label = label)) +
      geom_col() +
      coord_flip() + 
      labs(title = paste0("\n", other_department, " - ", title, "\n"),
           x = "\nForm\n", y = "\nCompliance (%)\n") +
      
      theme(legend.position = "bottom", legend.justification = c(0.4, 0),
            #legend.title = element_blank(),
            plot.title.position = "plot",
            plot.title = element_text(family = "Helvetica",
                                      size = 14,
                                      hjust = 0.5),
            axis.line = element_line(),
            panel.background = element_rect(fill = "white"), #gray96
            axis.text.x = element_text(hjust = 0.5, vjust = 0.8),
            
            strip.text = element_text(colour = "black"#, face = "italic"
            ),
            strip.background = element_rect(fill = "#B2BEB5")) +
      
      scale_fill_manual(values = c("At or above benchmark" = "#00b159",
                                   "Above 50 but below benchmark" = "orange",
                                   "Below 50" = "red")) +
      
      guides(fill = guide_legend(title = "Key:")) +
      
      scale_y_continuous(
        expand = c(0, 0),
        limits = c(0, 110),
        breaks = scales::breaks_pretty()
      ) +
      
      geom_richtext(fill = NA, label.color = NA,
                    hjust = if_else(intermediate$Percentage == 0, 0, 1), #1 originally, moved because 0's start at y axis line then 0.9
                    nudge_y = if_else(intermediate$Percentage == 0, 1, 0), #0.5
                    vjust = 0.5, #0.5
                    size = 2.7) +
      
      facet_wrap(~Quarter)
    
    
    print(graph) %>% 
      assign(value = .,
             x = paste0(tolower(col_quosure),
                        "_form_chart"),
             envir = .GlobalEnv)
    
    
  }
  
  walk(.x = c("DOCUMENTED", "TIMELY", "LEGIBLE", "COMPLETE", "ACCURATE", "MET"), #Form bar chart walk
       .f = form_charts)
  
  
  
  ## Saving the graphs
  save_plots <- function(placeholder){
    
    read_pptx(here("./Open Charts/2. Inpatient/R Code/Template.pptx")) %>%
      ph_with(value = block_list(fpar(ftext(glue("{placeholder} - Open Charts Documentation Compliance for {thequarter} {what_year}"),
                                            prop = fp_text(font.size = 18, font.family = "Helvetica")),
                                      fp_p = fp_par(text.align = "center"))),
              location = ph_location_type(type = "subTitle")) %>% 
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = documented_control_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = documented_form_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = timely_control_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = timely_form_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = legible_control_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = legible_form_chart,
              location = ph_location_fullsize()) %>%  
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = complete_control_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = complete_form_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = accurate_control_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = accurate_form_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = met_control_chart,
              location = ph_location_fullsize()) %>%
      
      add_slide(layout='Title and Content', master='Office Theme') %>%
      ph_with(value = met_form_chart,
              location = ph_location_fullsize()) %>% 
      
      print(glue(here("./Open Charts/2. Inpatient/Output/{what_year}/{thequarter}/{other_department}/{placeholder} - Open Charts Results.pptx")))
    
  }
  
  walk(.x = other_department, .f = save_plots) #Walk for saving plots for each department
  
  
  
}

# Vector of other departments in the dataset (filtering Responsible_department to only keep the others)
other_departments <- unique(data1$Responsible_Department) %>%
  .[. %in% c("Pavilion", "ICU", "CCU", "NHDU", "HDU", "Paediatric HDU")]


walk(.x = other_departments,
     .f = departmental_charts_others)

