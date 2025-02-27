# Load required packages
library(dplyr)
library(readr)
library(openxlsx)
library(isocountry)

# Define file paths
setwd("C:/Users/mitro/UNHCR/EGRISS Secretariat - 905 - Implementation of Recommendations/01_GAIN Survey/Integration & GAIN Survey/EGRISS GAIN Survey 2024")

analysis_ready_file <- "10 Data/Analysis Ready Files/analysis_ready_group_roster.csv"
data_file <- "08 Output/Dashboard Updates/01_Data/data.xlsx"
output_file <- "08 Output/Dashboard Updates/01_Data/data_updated.xlsx"
temp_output_file <- "08 Output/Dashboard Updates/01_Data/temp_data.xlsx"

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

# Recode the 2s and 8s to be 0s (not using recs) & 'Other' regions as blanks
group_roster <- group_roster %>%
  mutate(
    # Update `Use Recommendation` first
    `Use Recommendation` = case_when(
      `Use Recommendation` == 2 | `Use Recommendation` == 8 ~ 0,
      TRUE ~ `Use Recommendation`
    )
  ) %>%
  mutate(
    # Update `slicer_region`
    slicer_region = case_when(
      slicer_region == "Other" ~ NA_character_,
      TRUE ~ slicer_region
    )
  )

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
updated_data <- bind_rows(existing_data, group_roster) %>%
  mutate(slicer_userecs = case_when( # Create `slicer_userecs` based on the updated `Use Recommendation`
  `Use Recommendation` == 1 ~ "Yes",
  TRUE ~ "No"  # This should handle all 0s, 2s, 8s, and NAs as "No"
))

# Save the merged file
write.xlsx(updated_data, output_file)

#==================================================
# Title: Index and Update ISO Country Codes with Latitude and Longitude
#==================================================

library(dplyr)
library(readr)
library(readxl)
library(writexl)
library(countrycode)

# File Paths
setwd("C:/Users/mitro/UNHCR/EGRISS Secretariat - 905 - Implementation of Recommendations/01_GAIN Survey/Integration & GAIN Survey/EGRISS GAIN Survey 2024")

file1 <- "08 Output/Dashboard Updates/01_Data/data_updated.xlsx"
file2 <- "08 Output/Dashboard Updates/01_Data/world_country_and_usa_states_latitude_and_longitude_values.csv"
output_file <- "01_Data/data_updated_with_latlon.xlsx"

# Load the datasets
data1 <- read_excel(file1)
data2 <- read_csv(file2)

# Generate ISO-3 codes and ISO-2 codes in data2 and data1 based on country names
data2 <- data2 %>%
  mutate(iso3c = countrycode(country, "country.name", "iso3c"))

data1 <- data1 %>%
  mutate(iso2c = countrycode(Country, "country.name", "iso2c"))

# Ensure we only keep rows with valid ISO-3 codes that exist in data1
data2 <- data2 %>% filter(!is.na(iso3c) & iso3c %in% data1$iso3_code)

# Merge datasets based on ISO-3 codes, copying all columns from data2 to data1 but ensuring NA values are preserved in data1
data1 <- data1 %>%
  left_join(data2, by = c("iso2c" = "country_code"))

# Save the updated dataset
write_xlsx(data1, output_file)

# Count the number of matched ISO-3 codes
num_matches <- sum(!is.na(data1$latitude))

# Print completion message
message("Updated dataset saved successfully at: ", output_file)
message("Number of matched ISO-3 codes: ", num_matches)
