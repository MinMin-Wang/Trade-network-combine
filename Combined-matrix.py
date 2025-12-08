import pandas as pd
import os

# Define the directory containing the files (assumes current directory)
directory = '.'

# Initialize an empty DataFrame to store the final result
final_matrix = None

# Iterate through all xlsx files starting with "matrix_2010_" in the directory
for filename in os.listdir(directory):
    if filename.startswith("matrix_2010_") and filename.endswith(".xlsx"):
        # Read the current file
        file_path = os.path.join(directory, filename)
        df = pd.read_excel(file_path, index_col=0)  # Use first column as index

        # If final_matrix is empty, assign the current DataFrame to it
        if final_matrix is None:
            final_matrix = df
        else:
            # Add corresponding elements
            final_matrix = final_matrix.add(df, fill_value=0)

# Save the final result to a new xlsx file
final_matrix.to_excel("combined_matrix_2010.xlsx")
print("Combined matrix successfully saved to combined_matrix_2010.xlsx")
