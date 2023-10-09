import pandas as pd
import matplotlib.pyplot as plt

def proportional_stratified_sampling(input_file, output_file, sample_size):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)
    
    # Extract columns 1 and 6 as categories
    category_cols = [df.columns[3], df.columns[5]]

    # Calculate the proportions of each combination of the multiple categories
    grouped = df.groupby(list(category_cols))
    proportions = grouped.size() / len(df)

    # Perform proportional stratified sampling based on those proportions
    samples = []
    for categories, proportion in proportions.items():
        n_samples = round(proportion * sample_size)
        subset = df
        for col, cat in zip(category_cols, categories):
            subset = subset[subset[col] == cat]
        samples.append(subset.sample(n=n_samples, replace=True))

    sampled_df = pd.concat(samples, axis=0).reset_index(drop=True)

    # Write the resulting DataFrame to an Excel file
    sampled_df.to_excel(output_file, index=False)

    # Generate pie chart
    generate_pie_chart(sampled_df, category_cols)

def generate_pie_chart(df, category_cols):
    # Group the data by the categories and get the counts
    pie_data = df.groupby(category_cols).size()
    
    # Plot
    pie_data.plot.pie(autopct='%1.1f%%', startangle=90, figsize=(10, 8))
    plt.title("Proportional Stratified Sampling Distribution")
    plt.ylabel('')  # Hide the 'None' ylabel introduced by pandas
    plt.show()

# Use the function
input_filepath = "test_file.xlsx"
output_filepath = "test_file_sampled.xlsx"  
desired_sample_size = 20  # Adjust as needed

proportional_stratified_sampling(input_filepath, output_filepath, desired_sample_size)
