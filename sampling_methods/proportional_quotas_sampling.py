# Proportional Quota Sampling
quota_sampled_data = pd.DataFrame()

# Proportional quota sampling function
def proportional_quota_sampling(category, subcategories, data, total_quota):
    global quota_sampled_data
    quota_counts = defaultdict(int)

    for index, row in data.iterrows():
        subgroup = row[category]
        if subgroup in subcategories and quota_counts[subgroup] < subcategories[subgroup]:
            quota_sampled_data = quota_sampled_data.append(row)
            quota_counts[subgroup] += 1
        if len(quota_sampled_data) >= total_quota:
            break

# Perform proportional quota sampling for each category
for category, subcategories in sample_sizes.items():
    proportional_quota_sampling(category, subcategories, data, total_samples)

# Remove duplicates and add additional samples if necessary
quota_sampled_data = quota_sampled_data.drop_duplicates()
additional_samples_needed = total_samples - len(quota_sampled_data)
if additional_samples_needed > 0:
    additional_samples = data.drop(quota_sampled_data.index).sample(n=additional_samples_needed, random_state=1)
    quota_sampled_data = pd.concat([quota_sampled_data, additional_samples])

quota_sampled_data = quota_sampled_data.reset_index(drop=True)
