import re
import pandas as pd
import os

# Define paths for the input Excel files and the directory containing the filter files
input_excel_path = r'C:\CP360Automation\CP360 Performance Outliers.xlsx'
psr_excel_path = r'C:\CP360Automation\PSRNumber.xlsx'
real_filter_names_path = r'C:\CP360Automation\RealFilterNames.xlsx'
output_where_clause_path = r'C:\CP360Automation\where_clause.txt'
filters_directory = r'C:\CP360Automation'  # Directory containing the .txt filter files

# Function to remove JOIN-like conditions (e.g., table1.column = table2.column) from SQL
def remove_join_conditions(sql_text):
    # This regex identifies equality conditions that look like JOIN criteria
    pattern = r'\b\w+\.\w+\s*=\s*\w+\.\w+\s*(AND\s*)?'
    cleaned_sql = re.sub(pattern, '', sql_text)
    return cleaned_sql

# Function to extract the WHERE clause
def extract_where_clause(sql_text):
    where_clause_match = re.search(r'WHERE\s+(.+?)(GROUP BY|ORDER BY|$)', sql_text, re.DOTALL | re.IGNORECASE)
    if where_clause_match:
        return where_clause_match.group(1).strip()
    return None

# Function to load filters from a specified filter file
def load_filters(filter_file_path):
    with open(filter_file_path, 'r') as filter_file:
        filters = [line.strip().lower() for line in filter_file if line.strip()]
    return filters

# Function to categorize filters based on their existence in the WHERE clause
def categorize_filters(where_clause, mandatory_filters):
    exists = []
    does_not_exist = []
    filter_counts = {}  # Dictionary to store the count of matches for each filter
    extra_filters = {}  # Dictionary to store non-mandatory filters with more than 10 entries

    # Check each mandatory filter in the WHERE clause
    for filter_name in mandatory_filters:
        pattern = rf'\b{re.escape(filter_name)}\b'
        if re.search(pattern, where_clause, re.IGNORECASE):
            exists.append(filter_name)

            # Check for 'IN' clause and count items if present
            in_clause_match = re.search(rf"{re.escape(filter_name)}\s+IN\s*\((.*?)\)", where_clause, re.IGNORECASE)
            if in_clause_match:
                item_count = len(in_clause_match.group(1).split(','))
                filter_counts[filter_name] = item_count
            else:
                filter_counts[filter_name] = 1
        else:
            does_not_exist.append(filter_name)

    # Check for non-mandatory filters in WHERE clause, ensuring mandatory filters are excluded
    mandatory_filter_set = {filter.lower() for filter in mandatory_filters}
    for match in re.finditer(r'\b(\w+)\b\s+IN\s*\((.*?)\)', where_clause, re.IGNORECASE):
        filter_name = match.group(1).lower()
        item_count = len(match.group(2).split(','))

        # Ensure the filter name is not a mandatory filter before adding to extra_filters
        if filter_name not in mandatory_filter_set and item_count > 10:
            extra_filters[filter_name] = item_count

    return exists, does_not_exist, filter_counts, extra_filters

# Load the PSR data and real filter names data
psr_df = pd.read_excel(psr_excel_path, sheet_name='PSRNumber')
real_filter_df = pd.read_excel(real_filter_names_path, sheet_name='Filters')

# Create a dictionary for real filter names lookup
real_filter_names = dict(zip(real_filter_df['SQL filter'].str.lower(), real_filter_df['Filter Name']))

# Load the data from the CP360 Excel file, Sheet "CP360 Performance Outliers Anal"
df = pd.read_excel(input_excel_path, sheet_name='CP360 Performance Outliers Anal')

# Process each row in the DataFrame
for index, row in df.iterrows():
    sql_content = row['PSQL']  # Get the SQL content from the "PSQL" column
    path_value = row['Path']  # Get the file path from the "Path" column

    # Skip if PSQL is empty or contains only whitespace
    if not isinstance(sql_content, str) or not sql_content.strip():
        continue  # Skip updating "Dev Comments" if PSQL is empty

    # Check if path_value is valid and extract the DV name from it
    if not isinstance(path_value, str) or '/' not in path_value:
        df.at[index, 'Dev Comments'] = "Invalid Path provided."
        continue

    dv_name = path_value.split('/')[-1]  # Extract the DV name from the Path column data
    filter_filename = dv_name + '.txt'
    filter_file_path = os.path.join(filters_directory, filter_filename)

    # Check if the filter file exists
    if not os.path.exists(filter_file_path):
        df.at[index, 'Dev Comments'] = f"Filter file '{filter_filename}' not found."
        continue

    # Load mandatory filters from the specified filter file
    mandatory_filters = load_filters(filter_file_path)

    # Remove JOIN-like conditions from SQL content before extracting WHERE clause
    cleaned_sql_content = remove_join_conditions(sql_content)

    # Extract the WHERE clause from the cleaned SQL content
    where_clause = extract_where_clause(cleaned_sql_content)
    comments = ""

    # Check and categorize filters for this SQL
    if where_clause:
        # Optional: Save the WHERE clause to a separate file for reference
        with open(output_where_clause_path, 'w') as output_file:
            output_file.write("WHERE " + where_clause)

        # Categorize filters into those that exist and those that do not
        exists, does_not_exist, filter_counts, extra_filters = categorize_filters(where_clause, mandatory_filters)

        # Generate comments to write into the Excel file for this row
        if not does_not_exist:  # All filters exist
            comments = "The customer is executing a query that includes all mandatory filters, detailed as follows.\n\n"
            for filter_name in exists:
                count = filter_counts.get(filter_name, 1)
                real_name = real_filter_names.get(filter_name.lower(), filter_name)
                filter_label = "filter" if count == 1 else "filters"
                comments += f"{real_name} ({filter_name}): {count} {filter_label}\n"
        else:  # Some filters are missing
            comments = "The customer is executing a query utilizing the following filters:\n\n"
            for filter_name in exists:
                count = filter_counts.get(filter_name, 1)
                real_name = real_filter_names.get(filter_name.lower(), filter_name)
                filter_label = "filter" if count == 1 else "filters"
                comments += f"{real_name} ({filter_name}): {count} {filter_label}\n"
            comments += "\nHowever, not all required filters are being used. The following filters are missing:\n"
            for filter_name in does_not_exist:
                real_name = real_filter_names.get(filter_name.lower(), filter_name)
                comments += f"{real_name} ({filter_name})\n"

        # Add extra filters with more than 10 entries that are not mandatory
        if extra_filters:
            comments += "\n\nApart from the mandatory filters, the customer is using the following additional filters:\n\n"
            for filter_name, count in extra_filters.items():
                real_name = real_filter_names.get(filter_name, filter_name)
                comments += f"{real_name}: {count} filters\n"

    else:
        comments = "No WHERE clause found in the SQL query."

    # Add the note about 20.4 workbook redesign if the path ends with "Inventory Turns Analysis"
    if dv_name == "Inventory Turns Analysis":
        comments += "\nNote: In the 20.4 workbook redesign, the Advanced Time Prompt dashboard filter was replaced with " \
                    "standard filters such as Fiscal Calendar, Year, Quarter, and Period, in addition to the other mentioned filters."

    # Search for the corresponding PSR number in the PSR sheet
    psr_row = psr_df[psr_df['DV Names'] == dv_name]
    if not psr_row.empty:
        psr_number = psr_row.iloc[0]['PSR Number']
        comments += f"\n\nNote: A similar issue has been documented on the PSR Confluence page under PSR #{psr_number}."

    # Write the comments back to the same row under the "Dev Comments" column
    df.at[index, 'Dev Comments'] = comments

# Save the updated DataFrame back to the Excel file
df.to_excel(input_excel_path, sheet_name='CP360 Performance Outliers Anal', index=False)

print(f"The filter check results have been saved to the 'Dev Comments' column in the Excel file for all rows.")
