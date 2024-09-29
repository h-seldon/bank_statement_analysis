Python 3.12.4 (v3.12.4:8e8a4baf65, Jun  6 2024, 17:33:18) [Clang 13.0.0 (clang-1300.0.29.30)] on darwin
Type "help", "copyright", "credits" or "license()" for more information.
>>> !pip install xlsxwriter
... import pandas as pd
... from google.colab import files
... 
... # Upload the HDFC statement file manually
... uploaded = files.upload()
... 
... 
... # Assuming the file is named 'hdfc_statement.xlsx'
... file_path = 'hdfc_statement.xlsx'
... df = pd.read_excel(file_path)
... 
... # Get the uploaded file's name dynamically
... #file_path = next(iter(uploaded))  # Extract the first file's name
... 
... # Convert the 'Withdrawal Amt.', 'Deposit Amt.', and 'Closing Balance' columns to numeric
... #df['Withdrawal Amt.'] = pd.to_numeric(df['Withdrawal Amt.'], errors='coerce')
... #df['Deposit Amt.'] = pd.to_numeric(df['Deposit Amt.'], errors='coerce')
... #df['Closing Balance'] = pd.to_numeric(df['Closing Balance'], errors='coerce')
... 
... # Convert the 'Withdrawal Amt.', 'Deposit Amt.', and 'Closing Balance' columns to numeric
... df['Withdrawal Amt.'] = pd.to_numeric(df['Withdrawal Amt.'], errors='coerce').fillna(0)
... df['Deposit Amt.'] = pd.to_numeric(df['Deposit Amt.'], errors='coerce').fillna(0)
... df['Closing Balance'] = pd.to_numeric(df['Closing Balance'], errors='coerce').fillna(0)
... 
... # Calculate initial balance based on the first transaction
... first_withdrawal_amount = df['Withdrawal Amt.'].iloc[0]
... first_deposit_amount = df['Deposit Amt.'].iloc[0]
... first_balance = df['Closing Balance'].iloc[0]
... 
... initial_balance = first_balance + first_withdrawal_amount - first_deposit_amount
... 
# Categorize transactions by 'Remark'
grouped = df.groupby('Remark').agg({
    'Withdrawal Amt.': 'sum',
    'Deposit Amt.': 'sum'
})

# Separate inflows (deposits) and outflows (withdrawals)
credits = grouped['Deposit Amt.']
debits = grouped['Withdrawal Amt.']  # Keep debits as positive

# Sum up the cash inflow and outflow
total_inflow = credits.sum()
total_outflow = debits.sum()

# Calculate ending balance
ending_balance = initial_balance + total_inflow - total_outflow

# Prepare the output data
output_data = {
    'Initial Balance': [initial_balance],
    'Total Cash Inflow': [total_inflow],
    'Total Cash Outflow': [total_outflow],
    'Ending Balance': [ending_balance],
}

# Add categorized inflows and outflows
for category, amount in credits.items():
    output_data[f'Inflow - {category}'] = [amount]

for category, amount in debits.items():
    output_data[f'Outflow - {category}'] = [amount]

# Convert to DataFrame
output_df = pd.DataFrame(output_data)

# Write to an Excel file in Colab's local environment
output_excel_path = 'hdfc_statement_summary.xlsx'
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    output_df.to_excel(writer, sheet_name='Summary', index=False)
    credits.to_excel(writer, sheet_name='Credits')
    debits.to_excel(writer, sheet_name='Debits')

# Download the file
