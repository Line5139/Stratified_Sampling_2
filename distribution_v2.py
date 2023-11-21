import pandas as pd
import math

# Benchmark percentages
benchmark_percentages = {
    'Age': {'18 - 24 years old': 22, '25 - 34 years old': 26, '35 - 44 years old': 17, '45 - 54 years old': 13, '55 - 64 years old': 13, '65 and above years old': 9},
    'Ethnicity': {'Chinese': 22, 'Malay': 53, 'Indian': 6, 'Non-Malay Bumiputera': 12, 'Sikh': 0.3, 'Others (Please Specify)': 7},
    'Gender': {'Male': 48, 'Female': 52},
    'State': {'Kedah': 6.73, 'Kelantan': 5.5, 'Perlis': 0.92, 'Terengganu': 3.67, 'WP Labuan': 0.31, 'Perak': 7.65, 'Negeri Sembilan': 3.67, 'Penang': 5.20, 'I am not in Malaysia at this moment': 0, 'Sabah': 10.40, 'WP Kuala Lumpur': 6.12,
              'WP Putrajaya': 0.31, 'Melaka': 3.06, 'Pahang': 4.89, 'Selangor': 21.41, 'Sarawak': 7.65, 'Johor': 12.23},
    'Area': {'Big city': 48, 'Small Town': 43, 'Rural': 9},
    'Education': {'No formal education': 3, 'Studied from Secondary up to SPM': 61, 'Completed University, College or Vocational': 36},
    'Employment': {'Retiree': 11, 'Self employed / Gig job': 18, 'In between employment': 3, 'Not yet employed': 3, 'Permanently employed': 65}
}

def calculate_distance(point, benchmark):
    distance = 0
    for category, subcategories in benchmark.items():
        if point[category] in subcategories:
            distance += (point[category] - subcategories[point[category]]) ** 2
    return math.sqrt(distance)

def calculate_overlap(category, data):
    category_data = data[data[category] == category]
    subcategories = set(category_data['Subcategory'])
    overlap_scores = {}
    for other_category in categories:
        if other_category != category:
            other_category_data = data[data[other_category] == other_category]
            other_subcategories = set(other_category_data['Subcategory'])
            overlap_score = len(subcategories.intersection(other_subcategories)) / len(subcategories.union(other_subcategories))
            overlap_scores[other_category] = overlap_score
    sorted_overlaps = sorted(overlap_scores.items(), key=lambda x: x[1], reverse=True)
    return [x[0] for x in sorted_overlaps]

def filter_data_by_subcategory(data, category, subcategory):
    category_data = data[data[category] == subcategory]
    return category_data.to_dict('records')

def select_datapoints(input_file, num_datapoints, categories):
    # Read data from input Excel file
    data = pd.read_excel(input_file)

    # Calculate subcategory overlaps and select datapoints
    subcategories_overlap = {}
    selected_datapoints = []
    total_selected = 0
    for category in categories:
        subcategories_overlap[category] = calculate_overlap(category, data)
        for subcategory in subcategories_overlap[category]:
            subcategory_data = filter_data_by_subcategory(data, category, subcategory)
            subcategory_data.sort(key=lambda x: calculate_distance(x, benchmark_percentages[category][subcategory]), reverse=True)
            num_to_select = min(len(subcategory_data), num_datapoints - total_selected)
            selected_datapoints.extend(subcategory_data[:num_to_select])
            total_selected += num_to_select
            if total_selected >= num_datapoints:
                break
        if total_selected >= num_datapoints:
            break

    # Create DataFrame for selected datapoints
    selected_df = pd.DataFrame(selected_datapoints)

    # Write selected datapoints to Excel file
    output_file = 'test_run_Dis.xlsx'
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    selected_df.to_excel(writer, sheet_name='Selected Datapoints', index=False)

    # Calculate distribution comparison
    original_distribution = data['Category'].value_counts(normalize=True)
    selected_distribution = selected_df['Category'].value_counts(normalize=True)
    distribution_comparison = pd.DataFrame({
        'Original Distribution': original_distribution,
        'Selected Distribution': selected_distribution
    })

    # Write distribution comparison to Excel file
    distribution_comparison.to_excel(writer, sheet_name='Distribution Comparison')

    # Create a bar chart for distribution comparison
    workbook = writer.book
    worksheet = writer.sheets['Distribution Comparison']
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({'values': '=\'Distribution Comparison\'!$B$2:$B$' + str(len(selected_distribution) + 1)})
    chart.add_series({'values': '=\'Distribution Comparison\'!$C$2:$C$' + str(len(selected_distribution) + 1)})
    chart.set_x_axis({'name': 'Categories'})
    chart.set_y_axis({'name': 'Proportion'})
    worksheet.insert_chart('E2', chart)

    # Save the Excel file
    writer.save()

categories = ['Age', 'Ethnicity', 'State', 'Employment', 'Area', 'Gender', 'Education']  # List of categories
# Example usage
if __name__ == "__main__":
    input_file = 'input_data.xlsx'
    num_datapoints = 3000
    select_datapoints('output_excel/test_file_30K__.xlsx', num_datapoints, categories)

