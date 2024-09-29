Python 3.12.4 (v3.12.4:8e8a4baf65, Jun  6 2024, 17:33:18) [Clang 13.0.0 (clang-1300.0.29.30)] on darwin
Type "help", "copyright", "credits" or "license()" for more information.
>>> !pip install xlsxwriter
... import pandas as pd
... from google.colab import files
... 
... # Upload the file manually
... uploaded = files.upload()
... 
... # Assuming the file is named 'axis_statement.xlsx'
... #file_path = 'axis_statement.xlsx'
... 
... # Get the uploaded file's name dynamically
... file_path = next(iter(uploaded))  # Extract the first file's name
... 
... # Load the file using the dynamically obtained file name
... df = pd.read_excel(file_path)
... 
... # Define a function to clean the 'INR' prefix and convert to numeric
... def clean_currency(value):
...     if isinstance(value, str):
...         # Remove 'INR ' prefix and commas, then convert to float
...         return float(value.replace('INR ', '').replace(',', ''))
...     return value
... 
... # Apply the function to 'Amount' and 'Balance' columns
... df['Amount'] = df['Amount'].apply(clean_currency)
... df['Balance'] = df['Balance'].apply(clean_currency)
... 
... # Extract the necessary columns
... amounts = df['Amount']
... transaction_type = df['Transaction Type']
... balance = df['Balance']
... remarks = df['Remark']
... 
... # Calculate initial balance based on the first transaction
... first_transaction_type = transaction_type.iloc[0]
... first_transaction_amount = amounts.iloc[0]
... first_balance = balance.iloc[0]

if first_transaction_type == 'CR':
    initial_balance = first_balance - first_transaction_amount
elif first_transaction_type == 'DR':
    initial_balance = first_balance + first_transaction_amount

# Categorize transactions by 'Remark'
grouped = df.groupby('Remark').agg({
    'Amount': 'sum',
    'Transaction Type': lambda x: ','.join(x)
})

# Separate inflows (credits) and outflows (debits)
credits = grouped[grouped['Transaction Type'].str.contains('CR')]
debits = grouped[grouped['Transaction Type'].str.contains('DR')]

# Sum up the cash inflow and outflow
total_inflow = credits['Amount'].sum()
total_outflow = debits['Amount'].sum()

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
for category, amount in credits['Amount'].items():
    output_data[f'Inflow - {category}'] = [amount]

for category, amount in debits['Amount'].items():
    output_data[f'Outflow - {category}'] = [amount]

# Convert to DataFrame
output_df = pd.DataFrame(output_data)

# Write to an Excel file in Colab's local environment
output_excel_path = 'axis_statement_summary.xlsx'
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    output_df.to_excel(writer, sheet_name='Summary', index=False)
    credits.to_excel(writer, sheet_name='Credits')
    debits.to_excel(writer, sheet_name='Debits')

# Download the file
