import pandas as pd
import numpy as np

# Read the original dataset from the Excel file
input_file = 'output_excel/test_file_30K__.xlsx'
data = pd.read_excel(input_file)

# Define the desired distribution percentages
ethnicity_percentages = {
    'Malays': 0.56,
    'Chinese': 0.22,
    'Indians': 0.06,
    'Sikh': 0.003
}

# Calculate the number of data points to be sampled for each ethnicity
total_sample_size = 3000
sample_sizes = {ethnicity: int(total_sample_size * percentage) for ethnicity, percentage in ethnicity_percentages.items()}

# Filter and sample the data based on ethnicity percentages
sampled_data = pd.DataFrame(columns=data.columns)
for ethnicity, size in sample_sizes.items():
    ethnicity_data = data[data['Ethnicity'] == ethnicity]
    if size > len(ethnicity_data):
        sampled_data = pd.concat([sampled_data, ethnicity_data])
    else:
        sampled_data = pd.concat([sampled_data, ethnicity_data.sample(n=size)])

# Output the sampled data to a new Excel file
output_file = 'sampled_data.xlsx'
sampled_data.to_excel(output_file, index=False)

# Output the sampled distribution to a CSV file
sampled_distribution = pd.DataFrame.from_dict(ethnicity_percentages, orient='index', columns=['Target Percentage'])
actual_distribution = sampled_data['Ethnicity'].value_counts(normalize=True).to_frame().rename(columns={'Ethnicity': 'Sampled Percentage'})
distribution_comparison = pd.concat([actual_distribution, sampled_distribution], axis=1)
distribution_comparison.to_csv('distribution_comparison.csv', index_label='Ethnicity')

print(f'Sampled data saved to {output_file}')
print(f'Distribution comparison saved to distribution_comparison.csv')
