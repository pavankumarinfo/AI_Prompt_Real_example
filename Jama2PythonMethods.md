# Excel to Python Test Method Converter

## Overview

This Python script converts test cases from Excel files (both `.xls` and `.xlsx` formats) into formatted Python methods. It processes all Excel files from a specified folder, extracts relevant information, and generates corresponding Python methods based on the data. The script also generates a summary report in HTML format and logs the process.

## Features

1. **Read Excel Files from a Folder:**
   - Reads all Excel files from a folder named `import`.
   - Supports both `.xls` and `.xlsx` file types.

2. **Process Each Excel File:**
   - Groups test steps under a single method if the `ID` is the same.
   - Generates method names based on the `Name` column, replacing spaces with underscores.

3. **Format Each Test Case Method:**
   - Adds the following decorators above each method:
     ```python
     @pytest.mark.adhoc
     @pytest.mark.intest
     ```
   - Method content includes:
     - **Description:** From the `Description` column.
     - **Prerequisites:** From the `Prerequisites` column.
     - **Steps:** For each step under the same `ID`, formats as:
       ```
       Step #: Step Action
       ER: Step Expected Result
       Notes: Step Notes
       ```

4. **Clean and Adjust Text:**
   - Replaces all occurrences of non-breaking spaces (` `) with regular spaces.
   - Converts `NaN` values to `NA`.
   - Removes any line breaks within column text.
   - If a line exceeds 130 characters, inserts a line break with a `\` character and adds two additional tab spaces for the continuation lines.

5. **Output Files:**
   - Saves each output Python file in a folder named `converted_files`, preserving the original Excel file name (e.g., `test_data.xlsx` results in `test_data.py`).
   - Creates a summary report in HTML format with a pie chart showing the number of test cases per file. Saves this report as `summary_report.html`.
   - Logs the process, including the number of methods created per file and the total number of new methods, into a log file.

6. **Error Handling:**
   - If no Excel file is provided or if an invalid file is encountered, prints and logs the message: "Error: Invalid or missing Excel file."

## Script

```python
import pandas as pd
import os
import sys
import matplotlib.pyplot as plt
from collections import defaultdict

def clean_text(text):
    if pd.isna(text):
        return 'NA'
    if isinstance(text, str):
        text = text.replace(' ', '').replace('\n', ' ').replace('\r', ' ')
        return text.strip()
    return text

def wrap_text(text, width=130, indent='        '):  
    if isinstance(text, str) and len(text) > width:
        lines = []
        while len(text) > width:
            wrap_at = text.rfind(' ', 0, width)
            if wrap_at == -1:
                wrap_at = width
            lines.append(text[:wrap_at] + ' \\')
            text = text[wrap_at:].strip()
        lines.append(text)
        return ('\n' + indent).join(lines)
    return text

# Define the folder containing the input files
folder_path = 'import_excel_files'

# Check if the folder exists
if not os.path.isdir(folder_path):
    print("Error: Folder 'import' not found.")
    sys.exit(1)

# Get a list of all files in the folder
input_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.xls') or file.lower().endswith('.xlsx')]

# Check if any valid Excel files exist in the folder
if not input_files:
    print("Error: No Excel files found in the 'import' folder.")
    sys.exit(1)

# Initialize variables for summary
method_info = defaultdict(int)
total_excel_files = len(input_files)
total_new_methods = 0

# Create a log file to write the summary
log_file = 'converted_files\summary_log.txt'
html_file = 'converted_files\summary_report.html'

# Create the output directory if it doesn't exist
output_dir = 'converted_files'
os.makedirs(output_dir, exist_ok=True)

with open(log_file, 'w') as log, open(html_file, 'w') as html:
    html.write("<html><head><title>Summary Report</title></head><body>")
    html.write("<h1>Summary Report</h1>")
    
    # Iterate over each input file
    for input_file in input_files:
        # Read data from the Excel file
        try:
            df = pd.read_excel(os.path.join(folder_path, input_file))
        except Exception as e:
            print(f"Error reading Excel file '{input_file}': {e}")
            log.write(f"Error reading Excel file '{input_file}': {e}\n")
            continue

        # Check if the DataFrame is empty
        if df.empty:
            print(f"Warning: Excel file '{input_file}' contains no data.")
            log.write(f"Warning: Excel file '{input_file}' contains no data.\n")
            continue

        # Function to generate the method content
        def generate_test_method():
            method_content = ""

            # Iterate over unique IDs and group rows with the same ID
            for id_value, group in df.groupby('ID'):
                name = clean_text(group['Name'].iloc[0]).replace(' ', '_').replace('-','')
                description = wrap_text(clean_text(group['Description'].iloc[0]))
                prerequisites = wrap_text(clean_text(group['Prerequisites'].iloc[0]))

                method_content += f"@pytest.mark.adhoc\n"
                method_content += f"@pytest.mark.intest\n"
                method_content += f"def test_NEW_{name}(logger):\n"
                method_content += f" \"\"\"\n"
                method_content += f"    Description:\n"
                method_content += f"     {description}\n"
                method_content += f"    Prerequisites:\n"
                method_content += f"     {prerequisites}\n"
                method_content += f"    Steps:\n"

                for index, row in group.iterrows():
                    step_number = clean_text(row['Step #'])
                    step_action = wrap_text(clean_text(row['Step Action']))
                    step_expected_result = wrap_text(clean_text(row['Step Expected Result']))
                    step_notes = wrap_text(clean_text(row['Step Notes']))

                    method_content += f"        Step {step_number}: {step_action}\n"
                    method_content += f"         ER: {step_expected_result}\n"
                    method_content += f"         Notes: {step_notes}\n\n"

                method_content += f"    Project: Phoenix\n"
                method_content += f" \"\"\"\n\n"

                # Update summary variables
                method_info[input_file] += 1

            return method_content

        # Generate test method
        method_content = generate_test_method()

        # Write method content to a Python file with the same name as the Excel file
        output_file = os.path.join(output_dir, f'test_method_{os.path.splitext(input_file)[0]}.py')
        with open(output_file, 'w') as f:
            f.write(method_content)

        total_new_methods += method_info[input_file]

        print(f"Formatted test method has been generated for '{input_file}' in '{output_file}'")
        log.write(f"Formatted test method has been generated for '{input_file}' in '{output_file}'\n")

    # Write summary to console, log file, and HTML report
    summary_lines = [
        "\nSummary:",
        f"Total number of Excel files converted: {total_excel_files}",
        f"Total number of new methods created: {total_new_methods}\n",
        f"{'File Name':<40} {'Total Methods':<15}"
    ]

    for line in summary_lines:
        print(line)
        log.write(line + "\n")
        html.write("<p>" + line + "</p>")

    for file_name, count in method_info.items():
        line = f"{file_name:<40} {count:<15}"
        print(line)
        log.write(line + "\n")
        html.write("<p>" + line + "</p>")

    # Generate pie chart
    labels = list(method_info.keys())
    sizes = list(method_info.values())

    fig, ax = plt.subplots()
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
    ax.axis('equal')
    plt.title('Distribution of Methods per Excel File')

    # Save the pie chart as an image and embed it in the HTML report
    pie_chart_file = os.path.join(output_dir, 'pie_chart.png')
    plt.savefig(pie_chart_file)
    html.write(f'<img src="pie_chart.png" alt="Pie Chart">')

    html.write("</body></html>")

print(f"\nSummary log has been saved in '{log_file}'")
print(f"Summary report has been saved in '{html_file}'")

# Move the log and HTML report to the output directory
os.rename(log_file, os.path.join(output_dir, log_file))
os.rename(html_file, os.path.join(output_dir, html_file))

