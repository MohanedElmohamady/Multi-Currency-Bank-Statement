import pandas as pd
from datetime import datetime

# File paths
input_file = r"C:\Users\Learn\VS code\Senior Financial system analyst Western Union Task\Visa Account Statement 2024-10-02.xlsx"
output_file = f"C:\\Users\\Learn\\VS code\\Senior Financial system analyst Western Union Task\\Visa34BCD86630_{datetime.now().strftime('%Y%m%d%H%M')}.csv"

# Step 1: Load the data, skipping metadata rows
df = pd.read_excel(input_file, engine='openpyxl', header=23)

# Step 2: Extract Account Number
account_number = "34BCD86630"  # Extracted from file metadata manually

# Step 3: Split Transactions by Currency
# Identify empty rows that separate different currencies
df['is_empty'] = df.isnull().all(axis=1)
currency_groups = df.loc[df['is_empty'].shift(1, fill_value=False) | df['is_empty'], :].index.tolist()
currency_groups = [0] + currency_groups + [len(df)]  # Add start and end points

# Initialize the final DataFrame
all_transactions = pd.DataFrame()

# Process each currency group
for i in range(len(currency_groups) - 1):
    start = currency_groups[i]
    end = currency_groups[i + 1]
    subset = df.iloc[start:end].drop(columns=['is_empty']).dropna(subset=["Date Time"], how='all').copy()

    # Skip empty subsets or invalid groups
    if subset.empty or "Currency" not in subset.columns:
        continue

    # Extract the currency for the current group
    currency = subset["Currency"].iloc[0] if pd.notnull(subset["Currency"].iloc[0]) else "Unknown"

    # Add 'DR/CR' column
    subset['DR/CR'] = subset.apply(
        lambda x: 'CR' if pd.notnull(x['Credit Value']) else 'DR', axis=1
    )

    # Add 'Amount' column based on Debit/Credit
    subset['Amount'] = subset.apply(
        lambda x: x['Credit Value'] if x['DR/CR'] == 'CR' else x['Debit Value'], axis=1
    )

    # Add 'BankAcc Number'
    subset['BankAcc Number'] = f"{account_number}_{currency}"

    # Add 'Transaction Line' column
    subset['Transaction Line'] = '0'

    # Select only relevant columns for output
    subset = subset[[
        'Transaction Line', 'BankAcc Number', 'Date Time', 'Description',
        'Amount', 'DR/CR', 'Balance'
    ]]

    # Add Ending Balance Line for the current currency
    if not subset.empty:
        ending_balance = subset.iloc[-1]["Balance"]
        ending_balance_row = pd.DataFrame([{
            'Transaction Line': '9',
            'BankAcc Number': f"{account_number}_{currency}",
            'Date Time': subset.iloc[-1]["Date Time"],
            'Description': 'Ending Balance',
            'Amount': ending_balance,
            'DR/CR': 'CR' if ending_balance > 0 else 'DR',
            'Balance': None
        }])

        # Align `ending_balance_row` columns with `subset`
        ending_balance_row = ending_balance_row.reindex(columns=subset.columns, fill_value=None)

        # Remove all-NA columns from `ending_balance_row`
        ending_balance_row = ending_balance_row.dropna(axis=1, how='all')

        # Concatenate ending balance row
        subset = pd.concat([subset, ending_balance_row], ignore_index=True)

    # Append to the final DataFrame, ensuring no empty subsets are concatenated
    if not subset.empty:
        all_transactions = pd.concat([all_transactions, subset], ignore_index=True)

# Step 4: Save all transactions to CSV
all_transactions.to_csv(output_file, index=False)

print(f"File transformed and saved as {output_file}")
