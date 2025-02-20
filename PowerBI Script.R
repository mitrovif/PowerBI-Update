# Load required packages
library(dplyr)
library(readr)
library(openxlsx)
library(isocountry)

# Define file paths
analysis_ready_file <- "C:/Users/mitro/UNHCR/EGRISS Secretariat - 905 - Implementation of Recommendations/01_GAIN Survey/Integration & GAIN Survey/EGRISS GAIN Survey 2024/10 Data/Analysis Ready Files/analysis_ready_group_roster.csv"
data_file <- "C:/Users/mitro/UNHCR/EGRISS Secretariat - 905 - Implementation of Recommendations/01_GAIN Survey/Integration & GAIN Survey/EGRISS GAIN Survey 2024/08 Output/Dashboard Updates/01_Data/data.xlsx"
output_file <- "C:/Users/mitro/UNHCR/EGRISS Secretariat - 905 - Implementation of Recommendations/01_GAIN Survey/Integration & GAIN Survey/EGRISS GAIN Survey 2024/08 Output/Dashboard Updates/01_Data/data_updated.xlsx"
temp_output_file <- "C:/Users/mitro/UNHCR/EGRISS Secretariat - 905 - Implementation of Recommendations/01_GAIN Survey/Integration & GAIN Survey/EGRISS GAIN Survey 2024/08 Output/Dashboard Updates/01_Data/temp_data.xlsx"

# Check if files exist before proceeding
if (!file.exists(analysis_ready_file)) {
  stop("Error: Analysis Ready File does not exist.")
}
if (!file.exists(data_file)) {
  stop("Error: Data File does not exist.")
}

# Load analysis ready data
group_roster <- read_csv(analysis_ready_file, show_col_types = FALSE) %>% 
  filter(ryear == 2024)  # Filter only for ryear 2024

# Ensure PRO06 (Phase) is numeric
group_roster <- group_roster %>%
  mutate(PRO06 = as.numeric(PRO06))

# Transform variables
group_roster <- group_roster %>%
  mutate(
    slicer_year = ryear,
    slicer_region = region,
    slicer_type_of_example = case_when(
      g_conled == 1 ~ "Country Example",
      g_conled %in% c(2, 3) ~ "Institutional Example",
      TRUE ~ NA_character_
    ),
    slicer_population_group = case_when(
      PRO07.A == 1 & PRO07.B == 1 & PRO07.C == 1 ~ "Combined Refugees-IDPs-Stateless",
      PRO07.C == 1 & PRO07.A == 0 & PRO07.B == 0 ~ "Stateless",
      PRO07.A == 1 & PRO07.B == 0 & PRO07.C == 0 ~ "Refugees",
      PRO07.A == 1 & PRO07.B == 1 & PRO07.C == 0 ~ "Combined Refugees-Stateless",
      PRO07.B == 1 & PRO07.A == 0 & PRO07.C == 0 ~ "IDPs",
      PRO07.A == 1 & PRO07.B == 1 & PRO07.C == 0 ~ "Combined Refugees-IDPs",
      TRUE ~ "Any Other"
    ),
    slicer_phase = case_when(
      PRO06 == 2 ~ "Implementation",
      PRO06 == 3 ~ "Completed",
      PRO06 == 1 ~ "Design/Planning",
      PRO06 == 6 ~ "Other",
      PRO06 %in% c(8, NA) ~ "Don't Know/Undetermined",
      TRUE ~ NA_character_
    ),
    Examples = PRO03,
    `Use Recommendation` = PRO09,
    IRRS = PRO10.A,
    IRIS = PRO10.B,
    IROSS = PRO10.C,
    `Type of Example` = slicer_type_of_example,

    slicer_userecs = case_when(
      `Use Recommendation` == 0 ~ "No",
      `Use Recommendation` == 1 ~ "Yes",
      `Use Recommendation` == 2 ~ "No",
      `Use Recommendation` == 8 ~ "No",
      is.na(`Use Recommendation`) ~ "No",
      TRUE ~ NA_character_
    ),
    
    `Only IRRS` = ifelse(PRO10.A == 1 & PRO10.B == 0 & PRO10.C == 0, 1, 0),
    `Only IRIS` = ifelse(PRO10.B == 1 & PRO10.A == 0 & PRO10.C == 0, 1, 0),
    `Only IR0SS` = ifelse(PRO10.C == 1 & PRO10.A == 0 & PRO10.B == 0, 1, 0),
    mixed_recommendation = ifelse(PRO10.A + PRO10.B + PRO10.C >= 2, 1, 0),
    
    Survey = PRO08.A,
    `Administrative Data` = PRO08.B,
    Census = PRO08.C,
    `Data Integration` = PRO08.D,
    `Non-traditional` = PRO08.E,
    Strategy = PRO08.F,
    `Guidance/Toolkit` = PRO08.G,
    `Workshop/Training` = PRO08.H,
    
    iso3_code = mcountry,
    Country = mcountry
  ) %>%
  select(
    slicer_year, slicer_region, slicer_type_of_example, slicer_population_group, slicer_phase,
    Examples, `Use Recommendation`, IRRS, IRIS, IROSS, `Type of Example`, 
    `Only IRRS`, `Only IRIS`, `Only IR0SS`, mixed_recommendation, 
    Survey, `Administrative Data`, Census, `Data Integration`, 
    `Non-traditional`, Strategy, `Guidance/Toolkit`, `Workshop/Training`, 
    iso3_code, Country
  )

group_roster <- group_roster %>%
  left_join(isocountry, by=c("iso3_code"="name")) %>%
  mutate(iso3_code = `alpha_3`) %>%
  select(1:match("Country", names(group_roster))) %>%
  mutate(iso3_code = case_when(Country == "Democratic Republic of the Congo" ~ "COD",
                               Country == "Republic of Moldova" ~ "MDA",
                               Country == "State of Palestine" ~ "PSE",
                               Country == "Turkiye" ~ "TUR",
                               Country == "United Kingdom" ~ "GBR",
                               TRUE ~ iso3_code))

# Save the transformed group roster data separately
write.xlsx(group_roster, temp_output_file)

# Load existing Power BI data file
existing_data <- read.xlsx(data_file, sheet = 1)
colnames(existing_data) <- gsub("\\.", " ", colnames(existing_data)) # replacing the periods (.) with spaces ( )
existing_data <- existing_data[, !duplicated(colnames(existing_data))] # removing the duplicate columns

existing_data <- existing_data %>%
  mutate(Country = case_when(Country == "Kazakhstan." ~ "Kazakhstan",
                             Country == "Palestine, State of" ~ "State of Palestine",
                             Country == "Turkey" ~ "Turkiye",
                             .default = Country))

# # Ensure all columns match before merging
# all_columns <- colnames(existing_data)

# # Ensure group_roster has the same structure by adding missing columns
# for (col in all_columns) {
#   if (!(col %in% colnames(group_roster))) {
#     group_roster[[col]] <- NA  # Add missing columns with NA values
#   }
# }

# # Reorder columns to match existing data file
# group_roster <- group_roster[, all_columns]

# Merge the new transformed data with the existing data
updated_data <- bind_rows(existing_data, group_roster)

# Save the merged file
write.xlsx(updated_data, output_file)

# Print completion message
print("Data has been successfully updated and saved to: data_updated.xlsx")

