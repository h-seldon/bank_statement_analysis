Python 3.12.4 (v3.12.4:8e8a4baf65, Jun  6 2024, 17:33:18) [Clang 13.0.0 (clang-1300.0.29.30)] on darwin
Type "help", "copyright", "credits" or "license()" for more information.
import pandas as pd
from google.colab import files

# Upload the file manually
uploaded = files.upload()

# Assuming the file is named 'hsbc_statement.xlsx'
#file_path = 'hsbc_statement.xlsx'

# Get the uploaded file's name dynamically
file_path = next(iter(uploaded))  # Extract the first file's name

# Load the file using the dynamically obtained file name
df = pd.read_excel(file_path)

# Reverse the DataFrame to get it in chronological order
df = df.iloc[::-1].reset_index(drop=True)

# Convert empty cells to 0 for the relevant columns
df['Debit amount'] = df['Debit amount'].fillna(0)
df['Credit amount'] = df['Credit amount'].fillna(0)
df['Balance'] = df['Balance'].fillna(0)

# Convert negative debit amounts to positive
df['Debit amount'] = df['Debit amount'].abs()

# Extract the necessary columns
debit_amounts = df['Debit amount']
credit_amounts = df['Credit amount']
balance = df['Balance']
remarks = df['Remark']

# Calculate initial balance based on the first transaction
first_debit_amount = debit_amounts.iloc[0]
first_credit_amount = credit_amounts.iloc[0]
first_balance = balance.iloc[0]

initial_balance = first_balance - first_credit_amount + first_debit_amount

# Categorize transactions by 'Remark'
grouped = df.groupby('Remark').agg({
    'Debit amount': 'sum',
    'Credit amount': 'sum'
... })
... 
... # Separate inflows (deposits) and outflows (withdrawals)
... credits = grouped['Credit amount']
... debits = grouped['Debit amount']  # Keep debits as positive
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
... # Write to an Excel file in the local environment
... output_excel_path = 'hsbc_statement_summary.xlsx'
... with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
...     output_df.to_excel(writer, sheet_name='Summary', index=False)
...     credits.to_excel(writer, sheet_name='Credits')
...     debits.to_excel(writer, sheet_name='Debits')
... 
... # File path to access the output
