import pandas as pd
from sklearn.model_selection import train_test_split

# Load the Excel file
file_path = 'output_excel/test_file_30K__.xlsx'  # Replace with your file path
data = pd.read_excel(file_path)

# Define the target distributions
target_distributions = {
    'Age': {'18 - 24 years old': 22, '25 - 34 years old': 26, '35 - 44 years old': 17, '45 - 54 years old': 13, '55 - 64 years old': 13, '65 and above years old': 9},
    'Ethnicity': {'Chinese': 22, 'Malay': 53, 'Indian': 6, 'Non-Malay Bumiputera': 12, 'Sikh': .3 , 'Others (Please Specify)': 7},
    'Gender':{'Male': 48, 'Female': 52},
    'State': {'Kedah': 6.73, 'Kelantan': 5.5,  'Perlis': .92, 'Terengganu': 3.67 , 'WP Labuan': .31 ,'Perak': 7.65,'Negeri Sembilan': 3.67,'Penang': 5.20 ,'I am not in Malaysia at this moment': 0 ,'Sabah': 10.40,'WP Kuala Lumpur': 6.12,
               'WP Putrajaya': 0.31,'Melaka': 3.06,'Pahang': 4.89,'Selangor': 21.41,'Sarawak': 7.65,'Johor': 12.23},
    'Area' : {'Big city': 48,'Small Town': 43,'Rural': 9},
    'Education':{'No formal education': 3 ,'Studied from Secondary up to SPM': 61 ,'Completed University, College or Vocational ': 36},
    'Employment':{'Retiree': 11 ,'Self employed / Gig job': 18 ,'In between employment': 3,'Not yet employed': 3 ,'Permanently employed': 65}
}

# Total number of samples to be drawn
total_samples = 3200

# Function to calculate the number of samples needed for each stratum in each category
def calculate_sample_sizes(target_distributions, total_samples):
    sample_sizes = {category: {group: int(total_samples * percent / 100) 
                               for group, percent in groups.items()} 
                    for category, groups in target_distributions.items()}
    return sample_sizes

sample_sizes = calculate_sample_sizes(target_distributions, total_samples)

# Function to perform stratified sampling for a given category
def stratified_sampling(data, category, sample_sizes):
    sampled_data = []
    for group, size in sample_sizes[category].items():
        stratum = data[data[category] == group]
        sampled_stratum, _ = train_test_split(stratum, train_size=size, random_state=42, shuffle=True)
        sampled_data.append(sampled_stratum)
    return pd.concat(sampled_data)

# Initial stratified sampling based on Age
sampled_by_age = stratified_sampling(data, 'Age', sample_sizes)

# Function to adjust the distribution for Ethnicity and Area
def adjust_distribution(sampled_data, category, target_distribution, total_samples):
    adjusted_samples = pd.DataFrame()
    current_distribution = sampled_data[category].value_counts(normalize=True) * total_samples
    for group, target_percent in target_distribution.items():
        required_samples = int(total_samples * target_percent / 100)
        current_samples = int(current_distribution.get(group, 0))
        if current_samples < required_samples:
            additional_samples = sampled_data[sampled_data[category] == group].sample(
                n=(required_samples - current_samples), random_state=42, replace=True)
            adjusted_samples = pd.concat([adjusted_samples, additional_samples])
    return pd.concat([sampled_data, adjusted_samples])

# Adjusting the distribution for Ethnicity and Area
adjusted_by_ethnicity = adjust_distribution(sampled_by_age, 'Ethnicity', target_distributions['Ethnicity'], total_samples)
adjusted_final = adjust_distribution(adjusted_by_ethnicity, 'Area', target_distributions['Area'], total_samples)

# Function to prepare comparison data for a category
def prepare_comparison_data(category, sampled_data, target_distributions):
    benchmark_dist = pd.Series(target_distributions.get(category, {}), name=f'{category} - Benchmark')
    sampled_dist = sampled_data[category].value_counts(normalize=True) * 100
    sampled_dist.name = f'{category} - Sampled'
    return pd.concat([benchmark_dist, sampled_dist], axis=1)

# Creating the Excel file with separate sheets for each category
output_file_path = 'output_excel_v2/sampled_data.xlsx'  # Replace with your desired output path
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # Sheet for Picked Datapoints
    adjusted_final.to_excel(writer, sheet_name='Picked Datapoints', index=False)
    
    # Sheets for each distribution comparison
    for category in target_distributions.keys():
        comparison_data = prepare_comparison_data(category, adjusted_final, target_distributions)
        comparison_data.to_excel(writer, sheet_name=f'{category} Distribution', index=True)

    # Adding other categories (Employment, Education, State, Gender) to separate sheets
    additional_categories = ['Employment', 'Education', 'State', 'Gender']
    for category in additional_categories:
        comparison_data = prepare_comparison_data(category, adjusted_final, {})
        comparison_data.to_excel(writer, sheet_name=f'{category} Distribution', index=True)
