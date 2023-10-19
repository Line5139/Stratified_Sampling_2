import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split

def sample_and_plot_from_excel(input_file, output_file, strata_column, num_samples, sheet_name=None):
    """
    Read data from an Excel file, perform stratified sampling, save the results to a new Excel file, and plot the distribution comparison.

    :param input_file: Path to the input Excel file.
    :param output_file: Path where the output Excel file will be saved.
    :param strata_column: Name of the column in the dataframe to use for stratification.
    :param num_samples: Number of samples to retrieve. Make sure 
    :param sheet_name: Name of the sheet in the Excel file to read. If None, defaults to the first sheet.
    """
    # Step 1: Read the specified sheet from the Excel file into a pandas DataFrame
    if sheet_name is None:
        xls = pd.ExcelFile(input_file)
        sheet_name = xls.sheet_names[0]  # We default to the first sheet if none is specified
    
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # Validate inputs
    if strata_column not in df.columns:
        raise ValueError(f"Strata column '{strata_column}' does not exist in the dataframe.")
    
    if num_samples <= 0:
        raise ValueError("Number of samples must be a positive integer.")
    
    if num_samples > len(df):
        raise ValueError("Number of samples requested exceeds the number of available records.")

    # Step 2: Perform proportionate stratified sampling using sklearn
    samples, _ = train_test_split(df, train_size=num_samples, stratify=df[strata_column])

    # Step 3: Write the sampled data and the distribution data to a new Excel file
    with pd.ExcelWriter(output_file) as writer:
        samples.to_excel(writer, sheet_name='Sampled Data', index=False)
        
        # Distribution data
        original_counts = df[strata_column].value_counts(normalize=True).sort_index()
        sample_counts = samples[strata_column].value_counts(normalize=True).sort_index()
        distribution_df = pd.DataFrame({
            strata_column: original_counts.index,
            'Original Distribution': original_counts.values,
            'Sampled Distribution': [sample_counts.get(cat, 0) for cat in original_counts.index]
        })
        distribution_df.to_excel(writer, sheet_name='Distribution', index=False)

    # Step 4: Plot a bar chart comparing the distribution of the original dataset to the samples
    fig, ax = plt.subplots()
    width = 0.35  # the width of the bars
    ind = range(len(original_counts))  # the x locations for the groups
    p1 = ax.bar(ind, original_counts.values, width, label='Original Data')
    p2 = ax.bar([i + width for i in ind], [sample_counts.get(cat, 0) for cat in original_counts.index], width, label='Sampled Data')

    ax.set_title('Distribution comparison between Original and Sampled Data')
    ax.set_xticks([i + width / 2 for i in ind])
    ax.set_xticklabels(original_counts.index)
    ax.legend()

    plt.tight_layout()
    plt.show()

# Usage example (assuming you have an Excel file named 'test_file.xlsx' with a 'Category' column):
# sample_and_plot_from_excel('test_file.xlsx', 'output.xlsx', 'Q1 (Age)', 15)
