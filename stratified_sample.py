import pandas as pd

def stratified_sample(data, stratify_by, n):
    """
    Returns a stratified sample from a dataframe.
    
    data (DataFrame): the dataframe from which to sample.
    stratify_by (str): the column on which to stratify.
    n (int): the total sample size.
    
    Returns:
    DataFrame: the stratified sample.
    """
    # Get proportions of each group
    group_proportions = data[stratify_by].value_counts(normalize=True)

    # Calculate the size of sample for each group
    sample_sizes = (group_proportions * n).round().astype(int)

    # Sample from each group and concatenate results
    samples = [data[data[stratify_by] == group].sample(n=num_samples, replace=False) for group, num_samples in sample_sizes.items()]
    stratified_sample = pd.concat(samples, axis=0).reset_index(drop=True)

    return stratified_sample

    

# Load the excel data into a dataframe
df = pd.read_excel('test_file.xlsx')

# Get the stratified sample
sampled_df = stratified_sample(df, 'Q1 (Age)', 10)

# Save the sampled data into a new Excel file
sampled_df.to_excel('sampled_data.xlsx', index=False)

print("Sampled data saved to 'sampled_data.xlsx'")