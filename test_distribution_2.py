import pandas as pd

def filter_data_by_percentage(data, column_name, percentages):
    total_data_points = len(data)
    filtered_data = pd.DataFrame()

    original_distribution = data[column_name].value_counts(normalize=True).to_dict()

    for category, percentage in percentages.items():
        count = int(total_data_points * percentage)
        category_data = data[data[column_name] == category].head(count)
        filtered_data = filtered_data.append(category_data)

    filtered_distribution = filtered_data[column_name].value_counts(normalize=True).to_dict()

    return filtered_data, original_distribution, filtered_distribution

# Example usage:
# Load the Excel file
data = pd.read_excel('output_excel/test_file_30K__.xlsx')

# Define column name and percentages dictionary
column_name = 'Ethnicity'
percentages = {
    'Malays': 0.56,
    'Chinese': 0.22,
    'Indians': 0.06,
    'Sikh': 0.003
}

# Filter data based on the specified column and percentages
filtered_data, original_distribution, filtered_distribution = filter_data_by_percentage(data, column_name, percentages)

# Save the filtered data to a new Excel file
filtered_data.to_excel('filtered_data.xlsx', index=False)

# Save original and filtered distributions to an Excel file
distributions_data = {
    'Original Distribution': original_distribution,
    'Filtered Distribution': filtered_distribution
}

distributions_df = pd.DataFrame(distributions_data)
distributions_df.to_excel('distributions.xlsx', index=True)
