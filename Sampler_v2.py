import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split


class DataSampler:
    def __init__(self, file_path, total_samples, target_distributions):
        self.file_path = file_path
        self.total_samples = total_samples
        self.target_distributions = target_distributions
        self.data = pd.read_excel(file_path)
        self.sampled_data = None
        self.distribution_comparison = {}
        self.sampled_distributions = {}

    def calculate_subcategory_samples(self):
        return {
            category: {
                subcategory: int(self.total_samples * (percentage / 100))
                for subcategory, percentage in subcategories.items()
            }
            for category, subcategories in self.target_distributions.items()
        }


    def perform_stratified_sampling(self):
        subcategory_samples = self.calculate_subcategory_samples()
        sampled_data = pd.DataFrame()

        for category, subcat_counts in subcategory_samples.items():
            for subcategory, count in subcat_counts.items():
                subcat_data = self.data[self.data[category] == subcategory]
                if subcat_data.empty or count == 0:
                    continue
                if len(subcat_data) <= count:
                    sampled_subset = subcat_data
                else:
                    sampled_subset, _ = train_test_split(
                        subcat_data, train_size=count, random_state=1, stratify=subcat_data[category])
                sampled_data = pd.concat([sampled_data, sampled_subset])

        self.sampled_data = sampled_data.drop_duplicates().reset_index(drop=True)
        if len(self.sampled_data) > self.total_samples:
            self.sampled_data = self.sampled_data.sample(
                n=self.total_samples, random_state=1).reset_index(drop=True)

    def calculate_distribution_comparison(self):
        for category in self.target_distributions.keys():
            self.sampled_distributions[category] = (
                self.sampled_data[category].value_counts(normalize=True) * 100).to_dict()
            target_dist_df = pd.DataFrame(list(self.target_distributions[category].items(
            )), columns=[category, 'Target Percentage'])
            sampled_dist_df = pd.DataFrame(list(self.sampled_distributions[category].items(
            )), columns=[category, 'Sampled Percentage'])
            comparison_df = pd.merge(
                target_dist_df, sampled_dist_df, on=category)
            self.distribution_comparison[category] = comparison_df

    def save_to_excel(self, output_file_path):
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            self.sampled_data.to_excel(writer, sheet_name='Sampled Data', index=False)
            for category, dist_comp_df in self.distribution_comparison.items():
                dist_comp_df.to_excel(writer, sheet_name=f'{category} Comparison', index=False)
    

    def plot_distribution_comparisons(self):
        for category, dist_comp_df in self.distribution_comparison.items():
            fig, ax = plt.subplots(figsize=(10, 6))
            dist_comp_df.set_index(category).plot(kind='bar', ax=ax)
            plt.title(f'Distribution Comparison for {category}')
            plt.ylabel('Percentage')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.show()


# Example usage
if __name__ == "__main__":
    file_path = 'output_excel/test_file_30K__.xlsx'  # Path to the Excel file
    total_samples = 3000  # Number of samples to draw
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

    sampler = DataSampler(file_path, total_samples, target_distributions)
    sampler.perform_stratified_sampling()
    sampler.calculate_distribution_comparison()
    output_file_path = 'sample_datapoints.xlsx'  # Output Excel file
    sampler.save_to_excel(output_file_path)
    sampler.plot_distribution_comparisons()
