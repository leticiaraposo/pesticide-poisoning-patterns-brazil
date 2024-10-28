################################################################################
# Script: Pesticide Poisoning Patterns Analysis in Brazil - Clustering and Heatmap
# Author: Bruna Faria; Letícia Raposo
# Date: 10/28/2024
# 
# Description:
# This script performs an analysis of distinct pesticide poisoning patterns in Brazil 
# using clustering and heatmap visualization. The analysis includes data preprocessing, 
# reordering of factor levels, Multiple Correspondence Analysis (MCA), and 
# hierarchical clustering on principal components (HCPC). A heatmap is then generated 
# to visualize the identified clusters and patterns across categories.
#
# Requirements:
# - R libraries: readxl, dplyr, gtsummary, flextable, officer, pheatmap, NbClust, 
#   FactoMineR, summarytools, tidyverse, stringdist, RColorBrewer
#
# Outputs:
# - Table1.docx: Summary table by region saved as a Word document
# - Table3.docx: Summary table of clusters by region saved as a Word document
# - Table2.docx: Formatted table of clustering results with conditional formatting
# - Heatmap.png: Heatmap visualization of cluster categories, saved as a high-resolution PNG
#
# Notes:
# - Ensure that all required libraries are installed before running the script.
# - Customize file paths and document names as needed.
# - Adjust heatmap settings in the `pheatmap` function for specific visualization needs.
################################################################################

# Load necessary libraries
library(readxl)
library(gtsummary)
library(flextable)
library(summarytools)
library(tidyverse)
library(cluster)
library(NbClust)
library(factoextra)
library(FactoMineR)
library(officer)
library(stringdist)
library(pheatmap)

# Load the data
data <- read_excel("identification_distinct_pesticide_poisoning_patterns_Brazil_cluster_analysis_data.xlsx")

# Reorder factor levels for various variables based on logical or conventional order
data$Region <- factor(data$Region, levels = c("North", "Northeast", "Central-West", "Southeast", "South"))
data$Age <- factor(data$Age, levels = c("≤ 9", "10-19", "20-64", "≥ 65"))
data$Sex <- factor(data$Sex, levels = c("Female", "Male"))
data$`Race/ethnicity` <- factor(data$`Race/ethnicity`, levels = c("Asian", "Black", "Brown", "Indigenous", "White"))
data$`Employment status` <- factor(data$`Employment status`, levels = c("Unemployed", "Informally employed", "Formally employed", "Self-employed", "Other"))
data$`Work-related contamination` <- factor(data$`Work-related contamination`, levels = c("No", "Yes"))
data$`Exposure location` <- factor(data$`Exposure location`, levels = c("Residence", "Other", "Workplace"))
data$`Toxic agent group` <- factor(data$`Toxic agent group`, levels = c("Agricultural", "Domestic", "Public health-related"))
data$`Usage purpose` <- factor(data$`Usage purpose`, levels = c("Acaricide", "Herbicide", "Fungicide", "Insecticide", "Rodenticide", "Other", "Not applicable"))
data$`Exposure route` <- factor(data$`Exposure route`, levels = c("Cutaneous", "Respiratory", "Digestive", "Other"))
data$`Exposure circumstance` <- factor(data$`Exposure circumstance`, levels = c("Habitual use", "Environmental", "Other", "Accidental", "Suicide attempt"))
data$`Exposure type` <- factor(data$`Exposure type`, levels = c("Acute - single", "Acute - repeated", "Acute on chronic", "Chronic"))
data$Hospitalization <- factor(data$Hospitalization, levels = c("No", "Yes"))
data$`Care type` <- factor(data$`Care type`, levels = c("None", "Home care", "Outpatient", "Hospital"))
data$`Case outcome` <- factor(data$`Case outcome`, levels = c("Cure without sequela", "Cure with sequela", "Death by other cause", "Death by exogenous intoxication"))
data$`Educational level` <- factor(data$`Educational level`, levels = c("Illiterate", "Elementary school", "High school", "Higher education", "Not applicable"))

# Generate summary table by region, save as Word document
list("tbl_summary-fn:percent_fun" = function(x) sprintf(x * 100, fmt='%#.1f')) %>%
  set_gtsummary_theme() 

tbl_summary(data[, -1], by = "Region") %>%
  add_p() %>%
  add_overall() %>%
  as_flex_table() %>%
  save_as_docx(path = "Table1.docx")

# Prepare data for clustering: Remove 'Educational level' column and filter complete cases
data_cluster <- data %>% select(-`Educational level`) %>% drop_na()

# Consolidate levels in certain variables to simplify categories
levels(data_cluster$`Race/ethnicity`) <- c("Not white", "Not white", "Not white", "Not white", "White")
levels(data_cluster$`Case outcome`) <- c("Cure", "Cure", "Death", "Death")
levels(data_cluster$`Care type`) <- c("None/Home care", "None/Home care", "Outpatient/Hospital", "Outpatient/Hospital")
levels(data_cluster$`Employment status`) <- c("Unemployed/Informally employed", "Unemployed/Informally employed", "Formally employed/Self-employed", "Formally employed/Self-employed", "Other")

# Display frequency table for updated data
freq(data_cluster)

# Perform Multiple Correspondence Analysis (MCA) and visualize results
res.mca <- MCA(data_cluster[,-c(1,2)], graph = FALSE)
fviz_screeplot(res.mca, addlabels = TRUE)
fviz_mca_var(res.mca, choice = "mca.cor", repel = TRUE, ggtheme = theme_minimal())
fviz_mca_var(res.mca, repel = TRUE, ggtheme = theme_minimal())

# Hierarchical Clustering on Principal Components (HCPC) analysis
res.hcpc <- HCPC(res.mca, nb.clust = -1, graph = FALSE)

# Extract clusters and save regional summary with clusters
resul_cluster <- res.hcpc$data.clust
levels(resul_cluster$clust) <- c("Cluster 1", "Cluster 2", "Cluster 3")

# Modify factor levels in specific columns within the result data frame
levels(resul_cluster$`Employment status`) <- c("Unemployed/Informally employed", "Formally employed/Self-employed", "Other employment status")
levels(resul_cluster$`Exposure location`) <- c("Residence", "Other exposure location", "Workplace")

# Bind cluster information with regional data, then save summary as Word document
regiao_cluster <- data.frame(Region = data_cluster$Region, Year = data_cluster$Year, Cluster = resul_cluster$clust)
tbl_summary(regiao_cluster, by = Cluster, percent = "r") %>%
  add_p() %>%
  as_flex_table() %>%
  save_as_docx(path = "Table3.docx")

# Extract clusters and add variable names and cluster identifiers
cluster_12 <- data.frame(res.hcpc$desc.var$category$`1`)
cluster_12$Variavel <- rownames(cluster_12)
cluster_12$Cluster <- 1

cluster_22 <- data.frame(res.hcpc$desc.var$category$`2`)
cluster_22$Variavel <- rownames(cluster_22)
cluster_22$Cluster <- 2

cluster_32 <- data.frame(res.hcpc$desc.var$category$`3`)
cluster_32$Variavel <- rownames(cluster_32)
cluster_32$Cluster <- 3

# Concatenate all cluster data into a single DataFrame
data_clusters2 <- bind_rows(cluster_12, cluster_22, cluster_32)

# Round numeric columns in each cluster for consistency
round_columns <- function(df) {
  df %>%
    mutate(across(c(`Cla/Mod`, `Mod/Cla`, Global), round, 1),
           p.value = round(p.value, 3),
           v.test = round(v.test, 2))
}

cluster_12 <- round_columns(cluster_12)
cluster_22 <- round_columns(cluster_22)
cluster_32 <- round_columns(cluster_32)

# Merge clusters by "Variavel"
cluster_juntos <- cluster_12 %>%
  left_join(cluster_22, by = "Variavel", suffix = c("_Cluster1", "_Cluster2")) %>%
  left_join(cluster_32, by = "Variavel", suffix = c("", "_Cluster3"))

# Order dataframe by string similarity on the "Variavel" column
dist_matrix <- stringdistmatrix(cluster_juntos$Variavel, cluster_juntos$Variavel, method = "lv")
hclust_result <- hclust(as.dist(dist_matrix))
cluster_juntos <- cluster_juntos[hclust_result$order, ]

# Select and rename columns for final output
cluster_juntos <- cluster_juntos[, c(6, 3, 1, 2, 5, 8, 9, 12, 14, 15, 18)]
colnames(cluster_juntos) <- c("Categories", "Global", 
                              "Cla/Mod_Cluster1", "Mod/Cla_Cluster1", "v.test_Cluster1", 
                              "Cla/Mod_Cluster2", "Mod/Cla_Cluster2", "v.test_Cluster2", 
                              "Cla/Mod_Cluster3", "Mod/Cla_Cluster3", "v.test_Cluster3")

# Remove unwanted prefix in "Categories" using a regular expression
cluster_juntos$Categories <- sub(".*=", "", cluster_juntos$Categories)

# Create a flextable and apply conditional formatting based on v.test values
ft <- flextable(cluster_juntos)
for (i in 1:nrow(cluster_juntos)) {
  if (!is.na(cluster_juntos$v.test_Cluster1[i]) && cluster_juntos$v.test_Cluster1[i] > 0) {
    ft <- bold(ft, i = i, j = c("Cla/Mod_Cluster1", "Mod/Cla_Cluster1"))
  }
  if (!is.na(cluster_juntos$v.test_Cluster2[i]) && cluster_juntos$v.test_Cluster2[i] > 0) {
    ft <- bold(ft, i = i, j = c("Cla/Mod_Cluster2", "Mod/Cla_Cluster2"))
  }
  if (!is.na(cluster_juntos$v.test_Cluster3[i]) && cluster_juntos$v.test_Cluster3[i] > 0) {
    ft <- bold(ft, i = i, j = c("Cla/Mod_Cluster3", "Mod/Cla_Cluster3"))
  }
}

# Export the formatted flextable to a Word document
doc <- read_docx()
doc <- body_add_flextable(doc, value = ft)
print(doc, target = "Table2.docx")

# Prepare the data matrix for the heatmap
# Selecting specific columns for the heatmap and removing unnecessary ones
data_matrix <- as.matrix(cluster_juntos[,-c(1, 2, 3, 5, 6, 8, 9, 11)])
rownames(data_matrix) <- c(
  "Usage purpose: Not applicable",
  "Usage purpose: Other",
  "Usage purpose: Fungicide",
  "Usage purpose: Rodenticide",
  "Usage purpose: Insecticide",
  "Race/ethnicity: Not white",
  "Race/ethnicity: White",
  "Case outcome: Death",
  "Case outcome: Cure",
  "Sex: Female",
  "Sex: Male",
  "Age: ≥ 65 yr",
  "Age: ≤ 9 yr",
  "Age: 20-64 yr",
  "Age: 10-19 yr",
  "Exposure type: Chronic",
  "Exposure type: Acute on chronic",
  "Exposure type: Acute - single",
  "Exposure type: Acute - repeated",
  "Care type: Outpatient / Hospital",
  "Care type: None / Home care",
  "Hospitalization: Yes",
  "Hospitalization: No",
  "Toxic agent group: Public health-related",
  "Toxic agent group: Agricultural",
  "Toxic agent group: Domestic",
  "Employment status: Other",
  "Employment status: Unemployed / Informally employed",
  "Employment status: Formally employed / Self-employed",
  "Work-related contamination: No",
  "Work-related contamination: Yes",
  "Exposure circumstance: Suicide attempt",
  "Exposure circumstance: Environmental",
  "Exposure circumstance: Accidental",
  "Exposure circumstance: Other",
  "Exposure circumstance: Habitual use",
  "Exposure location: Other",
  "Exposure location: Residence",
  "Exposure location: Workplace",
  "Exposure route: Respiratory",
  "Exposure route: Digestive",
  "Exposure route: Other",
  "Exposure route: Cutaneous"
)

# Set column names to represent clusters
colnames(data_matrix) <- c("Cluster 1", "Cluster 2", "Cluster 3")

# Create and save the heatmap with specific visual and clustering settings
png("Heatmap.png", width = 10, height = 10, units = "in", res = 1200)
pheatmap(
  data_matrix,
  scale = "row",  # Normalizes values across rows
  clustering_distance_rows = "euclidean",  # Sets row clustering distance metric
  clustering_distance_cols = "euclidean",  # Sets column clustering distance metric
  clustering_method = "ward.D",  # Clustering method
  cluster_cols = FALSE,  # Disables column clustering
  cutree_rows = 7,  # Number of clusters to cut rows into
  fontsize_row = 10,  # Row font size for better readability
  treeheight_row = 200,  # Height of row dendrogram
  border_color = "black",  # Adds border for clarity
  legend_breaks = c(-1, 0, 1),  # Custom legend for normalized values
  legend_labels = c("Low", "Medium", "High"),  # Legend labels for interpretation
  display_numbers = FALSE  # Disables display of numbers on heatmap
)
dev.off()