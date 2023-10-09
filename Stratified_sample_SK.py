import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split

# Step 1: Read the Excel file into a pandas DataFrame
file_path = 'test_file.xlsx'
df = pd.read_excel(file_path)

# Assuming you want to stratify based on a column named 'category'
strata_column = 'Q1 (Age)'

# Step 2: Perform proportionate stratified sampling to obtain 10 samples using sklearn
sample_size = len(df) // 10
samples, _ = train_test_split(df, train_size=sample_size, stratify=df[strata_column])

# Step 3: Write the 10 samples and the distribution data to a new Excel file
output_path = 'sampled_data_SK_1.xlsx'
with pd.ExcelWriter(output_path) as writer:
    samples.to_excel(writer, sheet_name='Sampled Data', index=False)
    
    # Distribution data
    original_counts = df[strata_column].value_counts(normalize=True).sort_index()
    sample_counts = samples[strata_column].value_counts(normalize=True).sort_index()
    distribution_df = pd.DataFrame({
        'Category': original_counts.index,
        'Original Distribution': original_counts.values,
        'Sampled Distribution': [sample_counts.get(cat, 0) for cat in original_counts.index]
    })
    distribution_df.to_excel(writer, sheet_name='Distribution', index=False)

# Step 4: Plot a bar chart comparing the distribution of the original dataset to the 10 samples
fig, ax = plt.subplots()
width = 0.35
ind = range(len(original_counts))
p1 = ax.bar(ind, original_counts.values, width, label='Original Data')
p2 = ax.bar([i + width for i in ind], [sample_counts.get(cat, 0) for cat in original_counts.index], width, label='Sampled Data')

ax.set_title('Distribution comparison between Original and Sampled Data')
ax.set_xticks([i + width / 2 for i in ind])
ax.set_xticklabels(original_counts.index)
ax.legend()

plt.tight_layout()
plt.show()
