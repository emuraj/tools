import os

# Replace this with your folder path
folder_path = r"C:\Users\e.muraj\Downloads\Trackwise_Exports\Trackwise_Exports_12APR\Unzipped\WE_00D6g000005ilBVEAY_1"

# List all CSV files in the folder
csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]

# Print the results
print("CSV files found:")
for file in csv_files:
    print(file)
