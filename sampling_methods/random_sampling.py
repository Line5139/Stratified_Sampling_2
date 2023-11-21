import pandas as pd
import numpy as np
from collections import defaultdict

# Load your dataset
# data = pd.read_excel('your_dataset.xlsx')

# Define your benchmark percentages
benchmark_percentages = {
    # Your benchmark percentages for each category
}

# Calculate sample sizes for each category
total_samples = 4000
sample_sizes = {category: {key: int((percentage / 100) * total_samples) for key, percentage in benchmarks.items()} for category, benchmarks in benchmark_percentages.items()}

# Initialize a dictionary to hold the sampled data
sampled_data = defaultdict(list)

# Random quota sampling function
def random_quota_sampling(category, subcategories, data):
    for subgroup, size in subcategories.items():
        subgroup_data = data[data[category] == subgroup]
        if len(subgroup_data) >= size:
            sampled_subgroup = subgroup_data.sample(n=size, random_state=1)
        else:
            sampled_subgroup = subgroup_data
        sampled_data[category].extend(sampled_subgroup.index.tolist())

# Apply random quota sampling for each category
for category, subcategories in sample_sizes.items():
    random_quota_sampling(category, subcategories, data)

# Combine the indices from each category
unique_indices = set()
for indices in sampled_data.values():
    unique_indices.update(indices)

# Select unique samples
unique_sample_indices = np.random.choice(list(unique_indices), size=total_samples, replace=False)
final_sampled_data = data.loc[unique_sample_indices]
