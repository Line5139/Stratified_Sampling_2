import pandas as pd
from sklearn.model_selection import train_test_split

def load_data(file_path):
    """Load data from an Excel file."""
    return pd.read_excel(file_path)

def calculate_sample_sizes(target_distributions, total_samples):
    """Calculate sample sizes for each stratum in each category."""
    return {category: {group: int(total_samples * percent / 100) 
                       for group, percent in groups.items()} 
            for category, groups in target_distributions.items()}

def stratified_sampling(data, category, sample_sizes, random_state=42):
    """Perform stratified sampling for a given category."""
    sampled_data = []
    for group, size in sample_sizes[category].items():
        stratum = data[data[category] == group]
        sampled_stratum, _ = train_test_split(stratum, train_size=size, random_state=random_state, shuffle=True)
        sampled_data.append(sampled_stratum)
    return pd.concat(sampled_data)

def adjust_distribution_test(sampled_data, category, target_distribution, total_samples, random_state=42):
    """Adjust the distribution for a given category while maintaining unique data points."""
    adjusted_samples = pd.DataFrame()
    current_distribution = sampled_data[category].value_counts()

    for group, target_percent in target_distribution.items():
        required_samples = int(total_samples * target_percent / 100)
        current_samples = current_distribution.get(group, 0)

        group_data = sampled_data[sampled_data[category] == group]

        # If current samples are less than required, add as many as possible up to the limit of unique data points
        if current_samples < required_samples and len(group_data) > current_samples:
            additional_samples = group_data.sample(n=(required_samples - current_samples), 
                                                   random_state=random_state, replace=False)
            adjusted_samples = pd.concat([adjusted_samples, additional_samples])
        elif current_samples > required_samples:
            # Reduce samples if more than required
            reduced_samples = group_data.sample(n=required_samples, random_state=random_state, replace=False)
            adjusted_samples = pd.concat([adjusted_samples, reduced_samples])
        else:
            # If the samples match the requirement, add them as is
            adjusted_samples = pd.concat([adjusted_samples, group_data])

    # Ensure the total count matches the intended sample size
    if len(adjusted_samples) > total_samples:
        adjusted_samples = adjusted_samples.sample(n=total_samples, random_state=random_state)

    return adjusted_samples

#This method ensures that the datapoints taken are 3200. By filling it the rest of the datapool with random samples
def adjust_distribution(sampled_data, original_data, category, target_distribution, total_samples, random_state=42):
    """Adjust the distribution for a given category while maintaining unique data points and total sample size."""
    # First, remove all samples of the specified category from the original dataset
    remaining_data = original_data[~original_data.index.isin(sampled_data.index)]

    # Adjust the distribution within the sampled data
    adjusted_samples = pd.DataFrame()
    for group, target_percent in target_distribution.items():
        required_samples = int(total_samples * target_percent / 100)
        group_data = sampled_data[sampled_data[category] == group]
        current_samples = len(group_data)

        if current_samples > required_samples:
            # Reduce samples if more than required
            adjusted_group_data = group_data.sample(n=required_samples, random_state=random_state, replace=False)
        else:
            # If the samples are less or equal to the requirement, add them as is
            adjusted_group_data = group_data

        adjusted_samples = pd.concat([adjusted_samples, adjusted_group_data])

    # Add additional unique samples from remaining_data if total is less than 3200
    if len(adjusted_samples) < total_samples:
        additional_samples_needed = total_samples - len(adjusted_samples)
        additional_samples = remaining_data.sample(n=additional_samples_needed, random_state=random_state, replace=False)
        adjusted_samples = pd.concat([adjusted_samples, additional_samples])

    return adjusted_samples



def adjust_distribution_in_batches(sampled_data, category, target_distribution, total_samples, batch_size=500, random_state=42):
    """Adjust the distribution for a given category using batch processing."""
    # Calculate the number of batches
    num_batches = len(sampled_data) // batch_size + (0 if len(sampled_data) % batch_size == 0 else 1)

    adjusted_samples = pd.DataFrame()
    
    for batch_num in range(num_batches):
        # Get the current batch
        batch_start = batch_num * batch_size
        batch_end = min((batch_num + 1) * batch_size, len(sampled_data))
        batch_data = sampled_data.iloc[batch_start:batch_end]

        # Adjust the distribution within this batch
        current_distribution = batch_data[category].value_counts(normalize=True) * total_samples
        for group, target_percent in target_distribution.items():
            required_samples = int(total_samples * target_percent / 100)
            current_samples = int(current_distribution.get(group, 0))
            if current_samples < required_samples:
                additional_samples = batch_data[batch_data[category] == group].sample(
                    n=(required_samples - current_samples), random_state=random_state, replace=True)
            elif current_samples > required_samples:
                additional_samples = batch_data[batch_data[category] == group].sample(
                    n=required_samples, random_state=random_state)
            else:
                additional_samples = batch_data[batch_data[category] == group]

            adjusted_samples = pd.concat([adjusted_samples, additional_samples])

    return adjusted_samples


def prepare_comparison_data(category, sampled_data, target_distributions):
    """Prepare comparison data for a category."""
    benchmark_dist = pd.Series(target_distributions.get(category, {}), name=f'{category} - Benchmark')
    sampled_dist = sampled_data[category].value_counts(normalize=True) * 100
    sampled_dist.name = f'{category} - Sampled'
    return pd.concat([benchmark_dist, sampled_dist], axis=1)

def save_to_excel(sampled_data, target_distributions, output_file_path):
    """Save the sampled data and distribution comparisons to an Excel file."""
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        sampled_data.to_excel(writer, sheet_name='Picked Datapoints', index=False)
        for category in target_distributions.keys():
            comparison_data = prepare_comparison_data(category, sampled_data, target_distributions)
            comparison_data.to_excel(writer, sheet_name=f'{category} Distribution', index=True)
        additional_categories = ['Employment', 'Education', 'State', 'Gender']
        for category in additional_categories:
            comparison_data = prepare_comparison_data(category, sampled_data, {})
            comparison_data.to_excel(writer, sheet_name=f'{category} Distribution', index=True)

# Usage Example
file_path = 'output_excel/test_file_30K__.xlsx'
output_file_path = 'output_excel_v2/Modular_1_.xlsx'
data = load_data(file_path)
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
total_samples = 3200
sample_sizes = calculate_sample_sizes(target_distributions, total_samples)
sampled_by_age = stratified_sampling(data, 'Age', sample_sizes)
data = load_data('output_excel/test_file_30K__.xlsx')
# for _ in range(100):  # For example, 100 iterations
sampled_by_age = adjust_distribution(sampled_by_age, data , 'Ethnicity', target_distributions['Ethnicity'], total_samples)
sampled_by_age = adjust_distribution(sampled_by_age, data , 'Area', target_distributions['Area'], total_samples)
adjusted_final = sampled_by_age
save_to_excel(adjusted_final, target_distributions, output_file_path)
