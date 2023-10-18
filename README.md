# Stratified Sampling from Excel Data
This script reads data from an Excel file, performs stratified sampling based on a specified column, and then writes the sampled data along with its distribution to a new Excel file. Additionally, it plots a comparison of the distributions of the original and sampled data.

## Requirements
- **Python 3.x**
- **pandas**
- **matplotlib**
- **scikit-learn**
You can install the required packages using pip:

```bash pip install pandas matplotlib scikit-learn ```

## Overview of the Code
1. **Reading the Excel File**: The script starts by reading data from an Excel file into a pandas DataFrame.
2. **Stratified Sampling**: The data is stratified based on the column 'Q11 (Age)' by default. The stratified sampling ensures each category in the column is proportionally represented in the sample.
3. **Writing to a New Excel File**: The sampled data and its distribution are written to a new Excel file.
4. **Plotting the Distribution Comparison**: A bar chart is generated to visually compare the distribution of the original dataset to the sampled data.

## Usage
To use the script, call the sample_and_plot_from_excel function with appropriate parameters.
```bash sample_and_plot_from_excel(input_file, output_file, strata_column, num_samples, sheet_name=None)```

### Parameters
`input_file`: Path to the input Excel file.
`output_file`: Path where the output Excel file will be saved.
`strata_column`: Name of the column in the dataframe to use for stratification.
`num_samples`: Number of samples to retrieve.
`sheet_name` (Optional): Name of the sheet in the Excel file to read. Defaults to the first sheet if none is specified.

### Example
```bash sample_and_plot_from_excel('test_file.xlsx', 'output.xlsx', 'Q1 (Age)', 15) ```
1. The script will create an output Excel file `output.xlsx` with two sheets:
- Sampled Data: Containing the 15 stratified samples. In this case Q1 (Age).
- Distribution: Compares the distribution of the original dataset to the sampled data.
2. A bar chart comparing the distributions will also be displayed.
  
## Customization
- To change the input, output file names, strata_column, num_samples, modify the `input_file`, `output_file`, `strata_column` and `num_samples` variables, respectively.

## Limitations
- Ensure the number of samples requested does not exceed the number of available records.
- The specified `strata_column` should exist in the Excel file.


## License
This project is licensed under the MIT License.
