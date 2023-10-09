# Stratified Sampling from Excel Data
This script reads data from an Excel file, performs stratified sampling based on a specified column, and then writes the sampled data along with its distribution to a new Excel file. Additionally, it plots a comparison of the distributions of the original and sampled data.

## Requirements
**Python 3.x**
- **pandas**
- **matplotlib**
- **scikit-learn**
You can install the required packages using pip:

```bash pip install pandas matplotlib scikit-learn ```

## Usage

1. Ensure your data is in an Excel file format (e.g., `test_file.xlsx`).
2. Modify the `file_path` variable in the script to point to your Excel file.
3. If you wish to stratify based on a different column, modify the strata_column variable.
4. Run the script:
   
```bash python stratified_sampling_script.py ```

1. The script will create an output Excel file (sampled_data_SK_Employment.xlsx by default) with two sheets:
- Sampled Data: Contains the stratified samples.
- Distribution: Compares the distribution of the original dataset to the sampled data.
2. A bar chart comparing the distributions will also be displayed.

## Overview of the Code
1. **Reading the Excel File**: The script starts by reading data from an Excel file into a pandas DataFrame.
2. **Stratified Sampling**: The data is stratified based on the column 'Q11 (Employment Status)' by default. The stratified sampling ensures each category in the column is proportionally represented in the sample.
3. **Writing to a New Excel File**: The sampled data and its distribution are written to a new Excel file.
4. **Plotting the Distribution Comparison**: A bar chart is generated to visually compare the distribution of the original dataset to the sampled data.


## Customization
- To change the input or output file names, modify the file_path and output_path variables, respectively.
- To stratify based on a different column, modify the strata_column variable.
