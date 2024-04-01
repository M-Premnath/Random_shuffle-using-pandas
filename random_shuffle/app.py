import pandas as pd
import random 


# Function to read Excel file
def read_excel(file_path):
    return pd.read_excel(file_path)

# Function to shuffle the data based on the second column
def shuffle_data(data):
    # Extract the second column
    column2 = data.iloc[:, 1]

    
    # Shuffle the second column
    random.shuffle(column2)
    
    return data

# Function to write data to Excel file
def write_excel(data, file_path, sheet_name):
    # Create a deep copy of the data to avoid modifying the original data
    data = data.copy(deep=True)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)

# Sample usage
if __name__ == '__main__':
    # Read the input Excel file
    data = read_excel('random_shuffle\input.xlsx')

    # Shuffle the data
    shuffled_data = shuffle_data(data)

    # Write the shuffled data to the output Excel file
    write_excel(shuffled_data, 'output.xlsx', 'Sheet1')