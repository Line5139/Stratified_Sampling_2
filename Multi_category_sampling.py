import pandas as pd
import matplotlib.pyplot as plt

def prioritize_subcategories(df, categories):
    prioritized_subcategories = {}
    for category in categories:
        counts = df.groupby(category)[category].transform('count')
        sorted_subcategories = counts.sort_values(ascending=False).index
        prioritized_subcategories[category] = sorted_subcategories.tolist()
    return prioritized_subcategories

def proportionate_stratified_sampling(df, prioritized_subcategories, total_samples):
    sampled_data = pd.DataFrame(columns=df.columns)
    samples_count = 0

    for category, subcategories in prioritized_subcategories.items():
        for subcategory in subcategories:
            subcategory_data = df[df[category] == subcategory]
            subcategory_samples = min(len(subcategory_data), total_samples - samples_count)
            sampled_data = pd.concat([sampled_data, subcategory_data.sample(n=subcategory_samples, random_state=42)])
            samples_count += subcategory_samples

            if samples_count >= total_samples:
                break
        
        if samples_count >= total_samples:
            break

    return sampled_data

def perform_stratified_sampling(input_file, output_file, benchmark_percentages, total_samples=3000):
    try:
        # Read data from input Excel file
        df = pd.read_excel(input_file)

        # Prioritize subcategories based on overlapping counts
        prioritized_subcategories = prioritize_subcategories(df, benchmark_percentages.keys())

        # Perform proportionate stratified sampling
        sampled_data = proportionate_stratified_sampling(df, prioritized_subcategories, total_samples)

        # Save sampled data to output Excel file
        sampled_data.to_excel(output_file, index=False, sheet_name='Picked Datapoints')

        # Visualizations (bar chart comparing distributions)
        plt.figure(figsize=(10, 6))
        for category, subcategories in benchmark_percentages.items():
            original_counts = df[category].value_counts(normalize=True).sort_index()
            sampled_counts = sampled_data[category].value_counts(normalize=True).sort_index()
            plt.bar(original_counts.index, original_counts.values, alpha=0.5, label=f'Original {category}')
            plt.bar(sampled_counts.index, sampled_counts.values, alpha=0.5, label=f'Sampled {category}')

        plt.xlabel('Subcategories')
        plt.ylabel('Distribution')
        plt.title('Comparison between Original and Sampled Data')
        plt.xticks(rotation=45)
        plt.legend()
        plt.tight_layout()
        plt.savefig('distribution_comparison.png')
        plt.show()

        # Print a message indicating successful execution
        print(f'Picked {len(sampled_data)} data points matching the benchmark. Output saved to {output_file}.')

    except Exception as e:
        print(f"Error: {e}")

# Example usage:
# Define benchmark percentages for each category and subcategory
benchmark_percentages = {
    'Age': {'18 - 24 years old': 22, '25 - 34 years old': 26, '35 - 44 years old': 17, '45 - 54 years old': 13, '55 - 64 years old': 13, '65 and above years old': 9},
    'Ethnicity': {'Chinese': 22, 'Malay': 53, 'Indian': 6, 'Non-Malay Bumiputera': 12, 'Sikh': .3 , 'Others (Please Specify)': 7},
    'Gender':{'Male': 48, 'Female': 52},
    'State': {'Kedah': 6.73, 'Kelantan': 5.5,  'Perlis': .92, 'Terengganu': 3.67 , 'WP Labuan': .31 ,'Perak': 7.65,'Negeri Sembilan': 3.67,'Penang': 5.20 ,'I am not in Malaysia at this moment': 0 ,'Sabah': 10.40,'WP Kuala Lumpur': 6.12,
               'WP Putrajaya': 0.31,'Melaka': 3.06,'Pahang': 4.89,'Selangor': 21.41,'Sarawak': 7.65,'Johor': 12.23},
    'Area' : {'Big city': 48,'Small Town': 43,'Rural': 9},
    'Education':{'No formal education': 3 ,'Studied from Secondary up to SPM': 61 ,'Completed University, College or Vocational ': 36},
    'Employment':{'Retiree': 11 ,'Self employed / Gig job': 18 ,'In between employment': 3,'Not yet employed': 3 ,'Permanently employed': 65}
}

# Perform nested proportionate stratified sampling and generate the comparison chart
perform_stratified_sampling('output_excel/test_file_30K__.xlsx ', 'testrun_.xlsx', benchmark_percentages)