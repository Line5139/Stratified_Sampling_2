import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split

# Load the Excel file
file_path = 'output_excel/test_file_30K__.xlsx'
data = pd.read_excel(file_path)

# Define the target percentages for each category's subcategories
target_distributions = {
    'Age': {'18 - 24 years old': 22, '25 - 34 years old': 26, '35 - 44 years old': 17, '45 - 54 years old': 13, '55 - 64 years old': 13, '65 and above years old': 9},
    'Ethnicity': {'Chinese': 22, 'Malay': 53, 'Indian': 6, 'Non-Malay Bumiputera': 12, 'Sikh': 0.3, 'Others (Please Specify)': 7},
    'Gender': {'Male': 48, 'Female': 52},
    'State': {'Kedah': 6.73, 'Kelantan': 5.5, 'Perlis': 0.92, 'Terengganu': 3.67, 'WP Labuan': 0.31, 'Perak': 7.65, 'Negeri Sembilan': 3.67, 'Penang': 5.20, 'I am not in Malaysia at this moment': 0, 'Sabah': 10.40, 'WP Kuala Lumpur': 6.12,
              'WP Putrajaya': 0.31, 'Melaka': 3.06, 'Pahang': 4.89, 'Selangor': 21.41, 'Sarawak': 7.65, 'Johor': 12.23},
    'Area': {'Big city': 48, 'Small Town': 43, 'Rural': 9},
    'Education': {'No formal education': 3, 'Studied from Secondary up to SPM': 61, 'Completed University, College or Vocational': 36},
    'Employment': {'Retiree': 11, 'Self employed / Gig job': 18, 'In between employment': 3, 'Not yet employed': 3, 'Permanently employed': 65}
}

# Total samples to draw
total_samples = 3000

# Calculate the number of samples for each subcategory
subcategory_samples = {category: {subcategory: int(total_samples * (percentage / 100))
                                  for subcategory, percentage in subcategories.items()}
                       for category, subcategories in target_distributions.items()}

def adjusted_stratified_sampling(data, category, subcategory_counts):
    sampled_data = pd.DataFrame()
    for subcategory, count in subcategory_counts.items():
        subcat_data = data[data[category] == subcategory]
        if subcat_data.empty or count == 0:
            continue
        if len(subcat_data) <= count:
            sampled_subset = subcat_data
        else:
            sampled_subset, _ = train_test_split(subcat_data, train_size=count, random_state=1, stratify=subcat_data[category])
        sampled_data = pd.concat([sampled_data, sampled_subset])
    return sampled_data

# Perform the adjusted stratified sampling for each category
sampled_data = pd.DataFrame()
for category, subcat_counts in subcategory_samples.items():
    sampled_df = adjusted_stratified_sampling(data, category, subcat_counts)
    sampled_data = pd.concat([sampled_data, sampled_df]).drop_duplicates().reset_index(drop=True)

if len(sampled_data) > total_samples:
    sampled_data = sampled_data.sample(n=total_samples, random_state=1).reset_index(drop=True)

# Calculate the distribution of the sampled data for comparison
sampled_distributions = {category: (sampled_data[category].value_counts(normalize=True) * 100).to_dict()
                         for category in target_distributions.keys()}

# Create the Excel file with two sheets
output_file_path = 'sampled_data_comparison.xlsx'
writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
sampled_data.to_excel(writer, sheet_name='Sampled Data', index=False)

# Write the distribution comparison data to subsequent sheets
for category, dist_comp_df in distribution_comparison.items():
    dist_comp_df.to_excel(writer, sheet_name=f'{category} Comparison', index=False)

writer.save()

# Function to plot the distribution comparison
def plot_distribution_comparison(distributions, category):
    target_dist_df = pd.DataFrame(list(distributions[0].items()), columns=[category, 'Target Percentage'])
    sampled_dist_df = pd.DataFrame(list(distributions[1].items()), columns=[category, 'Sampled Percentage'])
    comparison_df = pd.merge(target_dist_df, sampled_dist_df, on=category)
    fig, ax = plt.subplots(figsize=(10, 6))
    comparison_df.set_index(category).plot(kind='bar', ax=ax)
    plt.title(f'Distribution Comparison for {category}')
    plt.ylabel('Percentage')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    return fig

# Plot and save the distribution comparison bar charts for each category
for category, dist_comp_df in distribution_comparison.items():
    fig = plot_distribution_comparison(
        [target_distributions[category], sampled_distributions[category]], 
        category
    )
    fig_file = f"{category}_distribution_comparison.png"
    fig.savefig(fig_file)
    plt.close(fig)
