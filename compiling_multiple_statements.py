Python 3.12.4 (v3.12.4:8e8a4baf65, Jun  6 2024, 17:33:18) [Clang 13.0.0 (clang-1300.0.29.30)] on darwin
Type "help", "copyright", "credits" or "license()" for more information.
import pandas as pd
from google.colab import files

# Function to combine all data
def combine_excel_files():
    # Upload multiple Excel files
    uploaded_files = files.upload()

    # Initialize empty lists to store data
    summary_data = {'Initial Balance': 0, 'Total Cash Inflow': 0, 'Total Cash Outflow': 0, 'Ending Balance': 0}
    credits_combined = pd.DataFrame(columns=['Remark', 'Amounts'])
    debits_combined = pd.DataFrame(columns=['Remark', 'Amounts'])

    # Loop through each uploaded file
    for file_name in uploaded_files.keys():
        # Load the file into a pandas ExcelFile object
        excel_file = pd.ExcelFile(file_name)

        # Load the 'Summary' sheet and update the total balances
        summary_df = pd.read_excel(excel_file, sheet_name='Summary')
        summary_data['Initial Balance'] += summary_df['Initial Balance'][0]
        summary_data['Total Cash Inflow'] += summary_df['Total Cash Inflow'][0]
        summary_data['Total Cash Outflow'] += summary_df['Total Cash Outflow'][0]
        summary_data['Ending Balance'] += summary_df['Ending Balance'][0]

        # Load the 'Credits' sheet
        credits_df = pd.read_excel(excel_file, sheet_name='Credits')
        # Combine relevant columns into 'Amounts'
        credits_columns = [col for col in credits_df.columns if 'Amount' in col or 'Deposit Amt' in col or 'Deposit Amt.' in col or 'Credit amount' in col or 'Deposit Amt (INR)' in col or 'Credit' in col]
        credits_df['Amounts'] = credits_df[credits_columns].sum(axis=1, skipna=True)
        # Remove rows where Amounts are 0 or NaN
        credits_df = credits_df[credits_df['Amounts'] > 0]
        # Group by 'Remark' and sum 'Amounts'
        credits_df = credits_df.groupby('Remark', as_index=False)['Amounts'].sum()
        credits_combined = pd.concat([credits_combined, credits_df], ignore_index=True)

        # Load the 'Debits' sheet
        debits_df = pd.read_excel(excel_file, sheet_name='Debits')
        # Combine relevant columns into 'Amounts'
        debits_columns = [col for col in debits_df.columns if 'Amount' in col or 'Withdrawal Amt' in col or 'Withdrawal Amt.' in col or 'Debit amount' in col or 'Withdrawal Amt (INR)' in col or 'Debit' in col]
        debits_df['Amounts'] = debits_df[debits_columns].sum(axis=1, skipna=True)
...         # Remove rows where Amounts are 0 or NaN
...         debits_df = debits_df[debits_df['Amounts'] > 0]
...         # Group by 'Remark' and sum 'Amounts'
...         debits_df = debits_df.groupby('Remark', as_index=False)['Amounts'].sum()
...         debits_combined = pd.concat([debits_combined, debits_df], ignore_index=True)
... 
...     # After concatenation, group credits and debits by 'Remark' again and sum
...     credits_combined = credits_combined.groupby('Remark', as_index=False)['Amounts'].sum()
...     debits_combined = debits_combined.groupby('Remark', as_index=False)['Amounts'].sum()
... 
...     # Sort Credits and Debits by Amounts in descending order
...     credits_combined = credits_combined.sort_values(by='Amounts', ascending=False)
...     debits_combined = debits_combined.sort_values(by='Amounts', ascending=False)
... 
...     # Prepare the final summary DataFrame
...     summary_final = pd.DataFrame([summary_data])
... 
...     # Create DataFrames for Credits and Debits section headers
...     credits_header = pd.DataFrame([{'Remark': '--- CREDITS ---'}])
...     debits_header = pd.DataFrame([{'Remark': '--- DEBITS ---'}])
... 
...     # Concatenate summary, credits, and debits data into one DataFrame
...     summary_final = pd.concat([summary_final, credits_header, credits_combined, debits_header, debits_combined], ignore_index=True)
... 
...     # Write the combined data to a new Excel file
...     output_file = 'combined_statement_summary_with_credits_debits_sorted.xlsx'
...     with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
...         # Write the Summary sheet, which now includes Credits and Debits
...         summary_final.to_excel(writer, sheet_name='Summary', index=False)
... 
...     # Download the final combined Excel file
...     files.download(output_file)
... 
... # Run the function to combine Excel files
