Python 3.12.4 (v3.12.4:8e8a4baf65, Jun  6 2024, 17:33:18) [Clang 13.0.0 (clang-1300.0.29.30)] on darwin
Type "help", "copyright", "credits" or "license()" for more information.
!pip install xlsxwriter
import pandas as pd
from google.colab import files

# Upload the ICICI statement file manually
uploaded = files.upload()

# Assuming the file is named 'icici_statement.xlsx'
#file_path = 'icici_statement.xlsx'

# Get the uploaded file's name dynamically
file_path = next(iter(uploaded))  # Extract the first file's name

# Load the file using the dynamically obtained file name
df = pd.read_excel(file_path)

# Define a function to clean the 'INR' prefix, commas, and convert to numeric
#def clean_currency(value):
#    try:
#        if isinstance(value, str):
#            # Remove 'INR ' prefix and commas, then convert to float
#            return float(value.replace('INR ', '').replace(',', ''))
#        return float(value)  # Handle non-string values directly
#    except ValueError:
#        print(f"Unable to convert value: {value}")
#        return 0.0  # Default to 0.0 for any problematic values

# Define a function to clean commas and convert to numeric
def clean_currency(value):
    try:
        if pd.isna(value):  # Check if the value is NaN (empty cell)
            return 0.0
        if isinstance(value, str):
            # Remove commas and convert to float
            return float(value.replace(',', ''))
        return float(value)  # Handle non-string values directly
    except ValueError:
        print(f"Unable to convert value: {value}")
        return 0.0  # Default to 0.0 for any problematic values

# Apply the function to 'Withdrawal Amt (INR)', 'Deposit Amt (INR)', and 'Balance (INR)' columns
df['Withdrawal Amt (INR)'] = df['Withdrawal Amt (INR)'].apply(clean_currency)
df['Deposit Amt (INR)'] = df['Deposit Amt (INR)'].apply(clean_currency)
df['Balance (INR)'] = df['Balance (INR)'].apply(clean_currency)

# Calculate initial balance based on the first transaction
first_withdrawal_amount = df['Withdrawal Amt (INR)'].iloc[0]
first_deposit_amount = df['Deposit Amt (INR)'].iloc[0]
first_balance = df['Balance (INR)'].iloc[0]

initial_balance = first_balance + first_withdrawal_amount - first_deposit_amount

# Categorize transactions by 'Remark'
grouped = df.groupby('Remark').agg({
    'Withdrawal Amt (INR)': 'sum',
    'Deposit Amt (INR)': 'sum'
... })
... 
... # Separate inflows (deposits) and outflows (withdrawals)
... credits = grouped['Deposit Amt (INR)']
... debits = grouped['Withdrawal Amt (INR)']  # Keep debits as positive
... 
... # Sum up the cash inflow and outflow
... total_inflow = credits.sum()
... total_outflow = debits.sum()
... 
... # Calculate ending balance
... ending_balance = initial_balance + total_inflow - total_outflow
... 
... # Prepare the output data
... output_data = {
...     'Initial Balance': [initial_balance],
...     'Total Cash Inflow': [total_inflow],
...     'Total Cash Outflow': [total_outflow],
...     'Ending Balance': [ending_balance],
... }
... 
... # Add categorized inflows and outflows
... for category, amount in credits.items():
...     output_data[f'Inflow - {category}'] = [amount]
... 
... for category, amount in debits.items():
...     output_data[f'Outflow - {category}'] = [amount]
... 
... # Convert to DataFrame
... output_df = pd.DataFrame(output_data)
... 
... # Write to an Excel file in Colab's local environment
... output_excel_path = 'icici_statement_summary.xlsx'
... with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
...     output_df.to_excel(writer, sheet_name='Summary', index=False)
...     credits.to_excel(writer, sheet_name='Credits')
...     debits.to_excel(writer, sheet_name='Debits')
... 
... # Download the file
