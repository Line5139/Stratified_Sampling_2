import pandas as pd
import matplotlib.pyplot as plt

def nested_proportionate_stratified_sampling(input_file, output_file, benchmark_percentages, total_samples):
    # Load data from the input Excel file
    df = pd.read_excel(input_file)

    # Calculate desired number of samples for each subcategory based on benchmarks
    desired_samples = {}
    for category, subcategories in benchmark_percentages.items():
        category_total_samples = sum(subcategories.values())
        category_desired_samples = {}
        for subcategory, percentage in subcategories.items():
            subcategory_samples = int(total_samples * (percentage / category_total_samples))
            category_desired_samples[subcategory] = subcategory_samples
        desired_samples[category] = category_desired_samples

    # Perform nested proportionate stratified sampling
    sampled_data = pd.DataFrame(columns=df.columns)  # Create an empty DataFrame to store the sampled data

    # Dictionary to store the expected and actual number of samples for each subcategory
    sampling_info = {}

    for category, subcategories in benchmark_percentages.items():
        category_info = {}
        for subcategory, num_samples in desired_samples[category].items():
            # Perform proportionate stratified sampling for the current subcategory
            subcategory_data = df[(df[category] == subcategory)].sample(n=num_samples, random_state=42)
            sampled_data = pd.concat([sampled_data, subcategory_data])

            # Store sampling info for the subcategory
            category_info[subcategory] = {
                'Expected Samples': num_samples,
                'Actual Samples': len(subcategory_data)
            }
        sampling_info[category] = category_info

    # Reset index of the sampled data
    sampled_data.reset_index(drop=True, inplace=True)

    # Save sampling info, percentages data, and sampled data to separate sheets in the output Excel file
    with pd.ExcelWriter(output_file) as writer:
        pd.DataFrame.from_dict(sampling_info, orient='index').to_excel(writer, sheet_name='Sampling Info')
        pd.DataFrame.from_dict(benchmark_percentages).to_excel(writer, sheet_name='Percentages', index=False)
        sampled_data.to_excel(writer, sheet_name='Sampled Data', index=False)

    # Generate a bar chart comparing original data to picked datapoints
    original_counts = df[category].value_counts(normalize=True).sort_index()
    sampled_counts = sampled_data[category].value_counts(normalize=True).sort_index()

    fig, ax = plt.subplots(figsize=(10, 6))
    width = 0.35
    ind = range(len(original_counts))

    p1 = ax.bar(ind, original_counts.values, width, label='Original Data')
    p2 = ax.bar([i + width for i in ind], sampled_counts.values, width, label='Sampled Data')

    ax.set_title(f'Distribution Comparison between Original and Sampled Data ({category} Category)')
    ax.set_xticks([i + width / 2 for i in ind])
    ax.set_xticklabels(original_counts.index)
    ax.legend()

    plt.tight_layout()
    plt.savefig('distribution_comparison.png')
    plt.show()

# Example usage:
# Define benchmark percentages for each category and subcategory
benchmark_percentages = {
    'Age': {'18 - 24 years old': 22, '25 - 34 years old': 26, '35 - 44 years old': 17, '45 - 54 years old': 13, '55 - 64 years old': 13, '65 and above years old': 9},
    'Ethnicity': {'Chinese': 22, 'Malay': 53, 'Indian': 6, 'Non-Malay Bumiputera': 12, 'Sikh': .3 , 'Others (Please Specify)': 7}
}

# Provide input and output file paths
# input_file = 'input_data.xlsx'  # Replace with your input Excel file
# output_file = 'sampled_data.xlsx'  # Output Excel file for sampled data

# Perform nested proportionate stratified sampling and generate the comparison chart
nested_proportionate_stratified_sampling('output_excel/test_file_30K__.xlsx ', 'output_excel/testrun_30K.xlsx', benchmark_percentages, total_samples=3000)
