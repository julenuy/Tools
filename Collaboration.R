

# Install necessary libraries if not already installed
if (!require("tidyverse")) install.packages("tidyverse")
if (!require("readxl")) install.packages("readxl")
library(tidyverse)
library(readxl)

# List of authors of interest with original umlauts
authors_of_interest <- c("dude1","dude2","dude3","dude4")

# Function to extract last name
extract_last_name <- function(author) {
  str_trim(str_split(author, ",")[[1]][1])
}

# Function to read and combine multiple Excel files from a directory
read_and_combine_excel_files <- function(directory_path) {
  file_paths <- list.files(directory_path, pattern = "*.xls", full.names = TRUE)
  combined_data <- data.frame()
  for (file_path in file_paths) {
    data <- read_excel(file_path) %>%
      select(Authors, `Article Title`, DOI) %>%
      mutate(Authors = str_replace_all(Authors, '"', '')) %>%
      mutate(Authors = strsplit(Authors, "; ")) %>%
      unnest(Authors) %>%
      mutate(Last_Name = sapply(Authors, extract_last_name))
    combined_data <- bind_rows(combined_data, data)
  }
  return(combined_data)
}

# Path to the directory containing the Excel files
directory_path <- "/Path/to/directory/folder" # Replace with the actual path to your directory

# Read and combine the data from all Excel files in the directory
combined_data <- read_and_combine_excel_files(directory_path)

# Remove duplicates based on Article_Title and DOI
combined_data <- combined_data %>% distinct(Last_Name, `Article Title`, DOI, .keep_all = TRUE)

# Filter based on last names of authors of interest
expanded_data <- combined_data %>%
  filter(Last_Name %in% authors_of_interest)

# Debugging: Print the first few rows of expanded_data to ensure filtering is correct
print(head(expanded_data))

# Create an empty matrix for the collaboration counts
collab_matrix <- matrix(0, nrow = length(authors_of_interest), ncol = length(authors_of_interest))
rownames(collab_matrix) <- authors_of_interest
colnames(collab_matrix) <- authors_of_interest


# Fill the matrix with collaboration counts
for (doi in unique(expanded_data$DOI)) {
  pub_authors <- expanded_data %>% filter(DOI == doi) %>% pull(Last_Name)
  pub_authors <- intersect(pub_authors, authors_of_interest)
  if (length(pub_authors) > 1) {
    for (i in 1:(length(pub_authors) - 1)) {
      for (j in (i + 1):length(pub_authors)) {
        a1 <- pub_authors[i]
        a2 <- pub_authors[j]
        if (a1 %in% authors_of_interest && a2 %in% authors_of_interest) {
          collab_matrix[a1, a2] <- collab_matrix[a1, a2] + 1
          collab_matrix[a2, a1] <- collab_matrix[a2, a1] + 1
        }
      }
    }
  }
}


# Save the collaboration matrix and author matrix to a CSV file
write.csv(collab_matrix, file = "/Path/to/folder/collaboration_matrix.csv") # Replace with the actual path to save the file
write.csv(expanded_data, file = "/Path/to/folder/Desktop/author_matrix.csv") # Replace with the actual path to save the file