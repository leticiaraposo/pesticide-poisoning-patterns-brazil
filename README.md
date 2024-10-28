# Pesticide Poisoning Patterns in Brazil

## Overview
This project analyzes pesticide poisoning patterns across Brazil, identifying distinct patterns through clustering and visualizing them via heatmap. The analysis uses data preprocessing, Multiple Correspondence Analysis (MCA), and Hierarchical Clustering on Principal Components (HCPC).

## Project Structure
- **data/**: Contains the input data file in `.xlsx` format.
- **scripts/**: Stores the analysis script in R (`clustering_heatmap_analysis.R`).
- **outputs/**: Directory for all generated outputs, including Word tables and a heatmap PNG.

## Setup Instructions
1. **Clone the Repository**:
    ```bash
    git clone https://github.com/your-username/pesticide-poisoning-patterns-brazil.git
    cd pesticide-poisoning-patterns-brazil
    ```

2. **Install Required Libraries**: Ensure all required R libraries are installed.
    ```R
    install.packages(c("readxl", "dplyr", "gtsummary", "flextable", "officer", "pheatmap", 
                       "NbClust", "FactoMineR", "summarytools", "tidyverse", "stringdist", 
                       "RColorBrewer"))
    ```

3. **Run the Script**:
    Open `scripts/identification_distinct_pesticide_poisoning_patterns_Brazil_cluster_analysis_script.R` in R or RStudio and execute the script.

## Outputs
- **Table1.docx**: Summary table by region
- **Table2.docx**: Formatted clustering table with conditional formatting
- **Table3.docx**: Clusters by region summary table
- **Heatmap.png**: High-resolution heatmap of cluster categories

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
