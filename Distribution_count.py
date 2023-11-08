import pandas as pd

def calculate_category_occurrences(input_file, output_file, category_columns):
    # Read data from the input Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)

    # Calculate the number of occurrences for each unique category in the specified columns
    category_counts = {}
    for column in category_columns:
        category_counts[column] = df[column].value_counts().reset_index()
        category_counts[column].columns = ['Category', 'Count']
        category_counts[column]['Decimal'] = category_counts[column]['Count'] / len(df)

    # Write the results to a new Excel file
    with pd.ExcelWriter(output_file) as writer:
        for column in category_columns:
            category_counts[column].to_excel(writer, sheet_name=f'{column} Occurrences', index=False)

    print(f"Category occurrences (decimal representation) for columns {', '.join(category_columns)} calculated and saved to '{output_file}'.")

# Example usage:
# calculate_category_occurrences('input.xlsx', 'output.xlsx', ['CategoryColumn1', 'CategoryColumn2'])
calculate_category_occurrences('output_excel/test_file_30K__.xlsx', 'output_count_ethnicity.xlsx', ['Ethnicity'])
