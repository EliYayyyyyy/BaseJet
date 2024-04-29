import os
import openpyxl

# Function to check the conditions and update the Excel file
def process_text_file(file_path, worksheet, partial_name):
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
            headers = lines[0].strip().split('\t')

            aligned_sequence_index = headers.index("Aligned_Sequence")
            reference_sequence_index = headers.index("Reference_Sequence")
            percent_reads_index = headers.index("%Reads")

            # Flag to track if any condition matches
            condition_matched = False

            for line in lines[1:]:
                data = line.strip().split('\t')
                reference_sequence = data[reference_sequence_index]
                aligned_sequence = data[aligned_sequence_index]
                percent_reads = data[percent_reads_index]

                # Define conditions for matching
                condition_1 = (
                    reference_sequence[14] == "A" and aligned_sequence[14] == "G" and
                    reference_sequence[31] == aligned_sequence[31] == "G" and
                    reference_sequence[32] == aligned_sequence[32] == "G" and
                    all(reference_sequence[i] == aligned_sequence[i] for i in range(10, 33) if i not in [14, 31, 32])
                )

                condition_2 = (
                    reference_sequence[15] == "A" and aligned_sequence[15] == "G" and
                    reference_sequence[31] == aligned_sequence[31] == "G" and
                    reference_sequence[32] == aligned_sequence[32] == "G" and
                    all(reference_sequence[i] == aligned_sequence[i] for i in range(10, 33) if i not in [15, 31, 32])
                )
                # be extremely careful that the index of number 25 bas is 24
                condition_3 = (
                    reference_sequence[24] == "T" and aligned_sequence[24] == "C" and
                    reference_sequence[7] == aligned_sequence[7] == "C" and
                    reference_sequence[8] == aligned_sequence[8] == "C" and
                    all(reference_sequence[i] == aligned_sequence[i] for i in range(10, 33) if i not in [24, 7, 8])
                )
                # be extremely careful that the index of number 26 bas is 25
                condition_4 = (
                    reference_sequence[25] == "T" and aligned_sequence[25] == "C" and
                    reference_sequence[7] == aligned_sequence[7] == "C" and
                    reference_sequence[8] == aligned_sequence[8] == "C" and
                    all(reference_sequence[i] == aligned_sequence[i] for i in range(10, 33) if i not in [25, 7, 8])
                )

                # Check conditions and append to worksheet
                if condition_1 or condition_2 or condition_3 or condition_4:
                    worksheet.append({"A": partial_name, "B": percent_reads})  # Format to float with 2 digits
                    condition_matched = True
                    break  # No need to check further conditions

            # If no condition matched, write N/A
            if not condition_matched:
                worksheet.append({"A": partial_name, "B": "N/A"})

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

# Set the path to the main folder, Manually iterate to process each file, consider automation in the future
main_folder = "/Users/qichenyuan/Desktop/NGS-088/CRISPRessoBatch_on_192"

# Extract the last two parts of the main folder path to create a sheet name
sheet_name = os.path.basename(main_folder).split("_")[-1:]

# Set the path for the Excel file
excel_file_path = "path"

workbook = openpyxl.load_workbook(excel_file_path)

# Create a new worksheet with the extracted sheet name
worksheet = workbook.create_sheet(title='_'.join(sheet_name))

# Set headers in columns A and B
worksheet['A1'] = 'Name'
worksheet['B1'] = 'Value'

# Iterate through subfolders
for folder_name in os.listdir(main_folder):
    if "CRISPResso_on" in folder_name:
        subfolder_path = os.path.join(main_folder, folder_name)

        # Check if the item is a directory before processing
        if os.path.isdir(subfolder_path):
            partial_name = folder_name.replace("CRISPResso_on_", "")
            text_file_name = [file for file in os.listdir(subfolder_path) if
                              "Alleles_frequency_table_around_sgRNA_" in file and file.endswith(".txt")]

            if text_file_name:
                text_file_path = os.path.join(subfolder_path, text_file_name[0])
                process_text_file(text_file_path, worksheet, partial_name)

# Save the Excel workbook
workbook.save(excel_file_path)

