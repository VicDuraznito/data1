import pandas as pd
import os
import time

# Define the path to your Excel file
file_path = 'C:\Users\Usuario\Desktop\MATERIAS\migr'  # Update with your actual file path
# Define the sheet name and the column to monitor
sheet_name = 'Sheet1'  # Update with your actual sheet name
monitor_column = 'Germany'  # Update with the column name you want to monitor

# Function to read the Excel file and return the specified column as a Series
def read_migration_data():
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df[monitor_column]

# Initialize the previous data to None
previous_data = None

# Main loop to check for changes
while True:
    try:
        # Read the current data from the Excel file
        current_data = read_migration_data()
        
        # Check if previous data is None (first run) or if the data has changed
        if previous_data is None:
            print("Monitoring started. Initial data loaded.")
            previous_data = current_data
        elif not current_data.equals(previous_data):
            print("Migration statistics have changed!")
            print("Previous Data:\n", previous_data)
            print("Current Data:\n", current_data)
            
            # Update previous_data to the current data after detecting a change
            previous_data = current_data
        
        # Wait for a specified interval before checking again (e.g., 10 minutes)
        time.sleep(600)  # Check every 600 seconds (10 minutes)
    
    except Exception as e:
        print(f"An error occurred: {e}")
        break

