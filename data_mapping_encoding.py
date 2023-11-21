import pandas as pd

def create_sequential_encoding_mappings(target_distributions):
    """
    Creates sequential encoding mappings for each characteristic based on the target distributions.
    """
    encoding_mappings = {}
    for characteristic, values in target_distributions.items():
        mapping = {value: str(i+1).zfill(2) for i, value in enumerate(values.keys())}
        encoding_mappings[characteristic] = mapping
    return encoding_mappings

def apply_encodings(data, encoding_mappings):
    """
    Applies the encodings to the data and creates a single encoded column, handling NaN values.
    """
    for characteristic, mapping in encoding_mappings.items():
        if characteristic in data.columns:
            # Apply encoding and fill NaN values with a placeholder (e.g., '00')
            data[f'{characteristic}_Encoded'] = data[characteristic].map(mapping).fillna('00')
    
    encoded_columns = [f'{char}_Encoded' for char in encoding_mappings.keys() if f'{char}_Encoded' in data.columns]
    # Ensure all values are strings before concatenation
    data['Encoded Characteristics'] = data[encoded_columns].astype(str).agg(''.join, axis=1)
    return data

def encode_and_export_excel(input_file_path, output_file_path, target_distributions):
    """
    Reads an Excel file, applies sequential encoding based on target distributions, 
    and exports the encoded data to a new Excel file.
    """
    data = pd.read_excel(input_file_path)
    encoding_mappings = create_sequential_encoding_mappings(target_distributions)
    encoded_data = apply_encodings(data, encoding_mappings)
    encoded_data.to_excel(output_file_path, index=False)

# Example usage
input_file_path = 'output_excel/test_file_30K__.xlsx'
output_file_path = 'output_excel_v2/output_excel_file.xlsx'
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

encode_and_export_excel(input_file_path, output_file_path, target_distributions)
