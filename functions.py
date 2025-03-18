import pandas as pd
import numpy as np

def calculate_value(row):
    if row['Value'] <= 2:
        return row['საბოლოო დებეტი'] - row['საბოლოო კრედიტი']
    elif 2 < row['Value'] < 6:
        return row['საბოლოო კრედიტი'] - row['საბოლოო დებეტი']
    elif row['Value'] >= 6:
        return row['ბრუნვა კრედიტი'] - row['ბრუნვა დებეტი']
    return None

def removeAddress(filepath):
    """
    Reads the Excel file at 'filepath' from the sheet named '1', removes the 'Address' column if it exists,
    and filters the DataFrame to only include rows where the first column's value has exactly 4 characters,
    ends with '0', but not with two zeros.
    
    Also adds 'Len', 'Last 1', 'Last 2', 'Mapping', 'Value', and 'Calculated' columns to the DataFrame.
    Returns the modified DataFrame with thousands‐separator formatting applied to all numeric columns 
    except for the very first one.
    """
    df = pd.read_excel(filepath, sheet_name="1")

    # Drop Address column if present
    if 'Address' in df.columns:
        df.drop('Address', axis=1, inplace=True)

    first_column = df.columns[0]
    first_col_stripped = df[first_column].astype(str).str.strip()

    # Filter rows
    df = df[
        (first_col_stripped.str.len() == 4)
        & (first_col_stripped.str.endswith("0"))
        & (~first_col_stripped.str.endswith("00"))
    ]

    # Add helper columns
    df['Len'] = df[first_column].astype(str).apply(len)
    df['Last 1'] = 0
    df['Last 2'] = df[first_column].astype(str).str[-2:].astype(int)
    df['Mapping'] = ""

    # Add 'Value' column if 'ანგარიში +' exists
    if 'ანგარიში +' in df.columns:
        df['Value'] = df['ანგარიში +'].astype(str).str[0].astype(int)

    # Sequential mapping
    library_mapping = [
        "Accounts Payables", "Accounts Receivables", "Accounts Receivables",
        "Accounts Receivables", "Cash & Cash equivalents", "Cash & Cash equivalents",
        "Cash & Cash equivalents", "COGS", "COGS", "FX Gain/Loss",
        "Interest Expense", "Interest Payable", "Inventories", "Inventories",
        "Inventories", "Inventories", "Long-term loan", "Maintenance Expense",
        "Net Intangible Assets", "Net Intangible Assets", "Net PPE", "Net PPE",
        "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE",
        "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE",
        "Other fixed assets", "Other fixed assets", "Other long term liabilities",
        "Other non-operating gain/loss", "Other non-operating gain/loss",
        "Other non-operating gain/loss", "Other non-operating gain/loss",
        "Other operating Expense", "Other operating Expense", "Other reserves",
        "Other Short term liabilities", "Rent Expense", "Retained Earning",
        "Retained Earning", "Retained Earning", "Salary Expense",
        "Salary Payables", "Sales Revenue", "Sales Revenue", "Services cost",
        "Share Capital", "Short-Term Loans", "Taxes Payable", "Taxes Payable",
        "Taxes Payable", "Taxes Payable", "Taxes Payable", "Taxes Payable",
        "Taxes Payable", "Transportation expense", "Utility Expense"
    ]
    library_index = 0
    for index, row in df.iterrows():
        if library_index < len(library_mapping):
            df.at[index, 'Mapping'] = library_mapping[library_index]
            library_index += 1

    # Reorder columns
    cols = [
        df.columns[0],
        'Len', 'Last 1', 'Last 2', 'Mapping'
    ] + [c for c in df.columns 
         if c not in ['Len','Last 1','Last 2','Mapping','Value']] + ['Value']
    df = df[cols]

    # Remove duplicate columns if any
    if 'ანგარიში +' in df.columns:
        df = df.loc[:, ~df.columns.duplicated()]

    # Convert all possible columns to int where feasible
    for col in df.columns[1:]:  # skip the FIRST column
        try:
            df[col] = pd.to_numeric(df[col], errors='ignore').round().astype(int)
        except ValueError:
            pass

    # Add final 'Calculated' column
    df['Calculated'] = df.apply(calculate_value, axis=1)

    # Now apply thousands-separator formatting to all numeric columns except the first column
    # so that we get e.g. 2,314,390 instead of 2314390
    for i, col in enumerate(df.columns):
        if i == 0:
            # This is the very first 4-digit column. Leave it alone.
            continue

        # If the column is numeric, format it. If not, do nothing.
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].apply(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else "")

        # If it’s not numeric (string or mixed), we just leave it.

    return df


def sum_by_mapping(df):
    """
    Takes a DataFrame that has 'Mapping' and 'Calculated' columns.
    Sums the 'Calculated' values for each unique 'Mapping', and then
    returns a custom-ordered financial statement-like DataFrame.
    """
    # Convert 'Calculated' back to numeric (it’s currently string with commas)
    df['Calculated'] = df['Calculated'].replace(",", "", regex=True)  # remove commas if any
    df['Calculated'] = pd.to_numeric(df['Calculated'], errors='coerce')

    grouped_df = df.groupby('Mapping', as_index=False)['Calculated'].sum()

    # Build the final custom-ordered table
    final_df = build_custom_report(grouped_df)
    return final_df


import pandas as pd
import numpy as np

def calculate_value(row):
    if row['Value'] <= 2:
        return row['საბოლოო დებეტი'] - row['საბოლოო კრედიტი']
    elif 2 < row['Value'] < 6:
        return row['საბოლოო კრედიტი'] - row['საბოლოო დებეტი']
    elif row['Value'] >= 6:
        return row['ბრუნვა კრედიტი'] - row['ბრუნვა დებეტი']
    return None

def removeAddress(filepath):
    """
    Reads the Excel file at 'filepath' from the sheet named '1', removes the 'Address' column if it exists,
    and filters the DataFrame to only include rows where the first column's value has exactly 4 characters,
    ends with '0', but not with two zeros.

    Also adds 'Len', 'Last 1', 'Last 2', 'Mapping', 'Value', and 'Calculated' columns to the DataFrame.
    Returns the modified DataFrame with thousands‐separator formatting applied to all numeric columns
    except for the very first one.
    """
    df = pd.read_excel(filepath, sheet_name="1")

    # Drop Address column if present
    if 'Address' in df.columns:
        df.drop('Address', axis=1, inplace=True)

    first_column = df.columns[0]
    first_col_stripped = df[first_column].astype(str).str.strip()

    # Filter rows
    df = df[
        (first_col_stripped.str.len() == 4)
        & (first_col_stripped.str.endswith("0"))
        & (~first_col_stripped.str.endswith("00"))
    ]

    # Add helper columns
    df['Len'] = df[first_column].astype(str).apply(len)
    df['Last 1'] = 0
    df['Last 2'] = df[first_column].astype(str).str[-2:].astype(int)
    df['Mapping'] = ""

    # Add 'Value' column if 'ანგარიში +' exists
    if 'ანგარიში +' in df.columns:
        df['Value'] = df['ანგარიში +'].astype(str).str[0].astype(int)

    # Sequential mapping
    library_mapping = [
        "Accounts Payables", "Accounts Receivables", "Accounts Receivables",
        "Accounts Receivables", "Cash & Cash equivalents", "Cash & Cash equivalents",
        "Cash & Cash equivalents", "COGS", "COGS", "FX Gain/Loss",
        "Interest Expense", "Interest Payable", "Inventories", "Inventories",
        "Inventories", "Inventories", "Long-term loan", "Maintenance Expense",
        "Net Intangible Assets", "Net Intangible Assets", "Net PPE", "Net PPE",
        "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE",
        "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE",
        "Other fixed assets", "Other fixed assets", "Other long term liabilities",
        "Other non-operating gain/loss", "Other non-operating gain/loss",
        "Other non-operating gain/loss", "Other non-operating gain/loss",
        "Other operating Expense", "Other operating Expense", "Other reserves",
        "Other Short term liabilities", "Rent Expense", "Retained Earning",
        "Retained Earning", "Retained Earning", "Salary Expense",
        "Salary Payables", "Sales Revenue", "Sales Revenue", "Services cost",
        "Share Capital", "Short-Term Loans", "Taxes Payable", "Taxes Payable",
        "Taxes Payable", "Taxes Payable", "Taxes Payable", "Taxes Payable",
        "Taxes Payable", "Transportation expense", "Utility Expense"
    ]
    library_index = 0
    for index, row in df.iterrows():
        if library_index < len(library_mapping):
            df.at[index, 'Mapping'] = library_mapping[library_index]
            library_index += 1

    # Reorder columns
    cols = [
        df.columns[0],
        'Len', 'Last 1', 'Last 2', 'Mapping'
    ] + [c for c in df.columns
         if c not in ['Len','Last 1','Last 2','Mapping','Value']] + ['Value']
    df = df[cols]

    # Remove duplicate columns if any
    if 'ანგარიში +' in df.columns:
        df = df.loc[:, ~df.columns.duplicated()]

    # Convert all possible columns to int where feasible (skipping the FIRST column)
    for col in df.columns[1:]:
        try:
            df[col] = pd.to_numeric(df[col], errors='ignore').round().astype(int)
        except ValueError:
            pass

    # Add final 'Calculated' column
    df['Calculated'] = df.apply(calculate_value, axis=1)

    # Now apply thousands-separator formatting to all numeric columns except the first column
    for i, col in enumerate(df.columns):
        if i == 0:
            # This is the very first 4-digit column. Leave it as is.
            continue
        # If the column is numeric, format it with thousands separators
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].apply(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else "")

    return df


def sum_by_mapping(df):
    """
    Takes a DataFrame that has 'Mapping' and 'Calculated' columns.
    Sums the 'Calculated' values for each unique 'Mapping', and then
    returns a custom-ordered financial statement-like DataFrame.
    """
    # Convert 'Calculated' back to numeric (it may be string with commas)
    df['Calculated'] = df['Calculated'].replace(",", "", regex=True)  # remove commas if any
    df['Calculated'] = pd.to_numeric(df['Calculated'], errors='coerce')

    grouped_df = df.groupby('Mapping', as_index=False)['Calculated'].sum()

    # Build the final custom-ordered table
    final_df = build_custom_report(grouped_df)
    return final_df


def build_custom_report(grouped_df):
    """
    Given a DataFrame with columns ['Mapping', 'Calculated'],
    produce a final table in the exact order requested, with
    totals and placeholders (#REF!) where needed.

    Also makes rows with "Total ..." HTML-bold by storing <b>...</b>.
    """
    # 1) Convert the grouped_df to a dict for easy lookup
    sum_dict = dict(zip(grouped_df['Mapping'], grouped_df['Calculated']))

    # 2) Define the exact rows/sequence you want
    row_order = [
        "Cash & Cash equivalents",
        "Accounts Receivables",
        "Tax assets",
        "Inventories",
        "Advances Paid",
        "Other current assets",
        "Total Current Assets",
        "Net PPE",
        "Net Intangible Assets",
        "Other fixed assets",
        "Total Fixed Assets",
        "Total Assets",
        "Accounts Payables",
        "Salary Payables",
        "Short-Term Loans",
        "Taxes Payable",
        "Interest Payable",
        "Other Short term liabilities",
        "Total current Liabilities",
        "Long-term loan",
        "Other long term liabilities",
        "Total long term Liabilities",
        "Total Liabilities",
        "Share Capital",
        "Retained Earning",
        "Other reserves",
        "Shareholder equity",
        "Liabilities & Equity",
        "Check"
    ]

    # 3) Initialize placeholders
    import pandas as pd
import numpy as np

def calculate_value(row):
    if row['Value'] <= 2:
        return row['საბოლოო დებეტი'] - row['საბოლოო კრედიტი']
    elif 2 < row['Value'] < 6:
        return row['საბოლოო კრედიტი'] - row['საბოლოო დებეტი']
    elif row['Value'] >= 6:
        return row['ბრუნვა კრედიტი'] - row['ბრუნვა დებეტი']
    return None

def removeAddress(filepath):
    """
    Reads the Excel file at 'filepath' from the sheet named '1', removes the 'Address' column if it exists,
    and filters the DataFrame to only include rows where the first column's value has exactly 4 characters,
    ends with '0', but not with two zeros.
    
    Also adds 'Len', 'Last 1', 'Last 2', 'Mapping', 'Value', and 'Calculated' columns to the DataFrame.
    Returns the modified DataFrame with thousands‐separator formatting applied to all numeric columns 
    except for the very first one.
    """
    df = pd.read_excel(filepath, sheet_name="1")

    # Drop Address column if present
    if 'Address' in df.columns:
        df.drop('Address', axis=1, inplace=True)

    first_column = df.columns[0]
    first_col_stripped = df[first_column].astype(str).str.strip()

    # Filter rows
    df = df[
        (first_col_stripped.str.len() == 4)
        & (first_col_stripped.str.endswith("0"))
        & (~first_col_stripped.str.endswith("00"))
    ]

    # Add helper columns
    df['Len'] = df[first_column].astype(str).apply(len)
    df['Last 1'] = 0
    df['Last 2'] = df[first_column].astype(str).str[-2:].astype(int)
    df['Mapping'] = ""

    # Add 'Value' column if 'ანგარიში +' exists
    if 'ანგარიში +' in df.columns:
        df['Value'] = df['ანგარიში +'].astype(str).str[0].astype(int)

    # Sequential mapping
    library_mapping = [
        "Accounts Payables", "Accounts Receivables", "Accounts Receivables",
        "Accounts Receivables", "Cash & Cash equivalents", "Cash & Cash equivalents",
        "Cash & Cash equivalents", "COGS", "COGS", "FX Gain/Loss",
        "Interest Expense", "Interest Payable", "Inventories", "Inventories",
        "Inventories", "Inventories", "Long-term loan", "Maintenance Expense",
        "Net Intangible Assets", "Net Intangible Assets", "Net PPE", "Net PPE",
        "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE",
        "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE", "Net PPE",
        "Other fixed assets", "Other fixed assets", "Other long term liabilities",
        "Other non-operating gain/loss", "Other non-operating gain/loss",
        "Other non-operating gain/loss", "Other non-operating gain/loss",
        "Other operating Expense", "Other operating Expense", "Other reserves",
        "Other Short term liabilities", "Rent Expense", "Retained Earning",
        "Retained Earning", "Retained Earning", "Salary Expense",
        "Salary Payables", "Sales Revenue", "Sales Revenue", "Services cost",
        "Share Capital", "Short-Term Loans", "Taxes Payable", "Taxes Payable",
        "Taxes Payable", "Taxes Payable", "Taxes Payable", "Taxes Payable",
        "Taxes Payable", "Transportation expense", "Utility Expense"
    ]
    library_index = 0
    for index, row in df.iterrows():
        if library_index < len(library_mapping):
            df.at[index, 'Mapping'] = library_mapping[library_index]
            library_index += 1

    # Reorder columns
    cols = [
        df.columns[0],
        'Len', 'Last 1', 'Last 2', 'Mapping'
    ] + [c for c in df.columns 
         if c not in ['Len','Last 1','Last 2','Mapping','Value']] + ['Value']
    df = df[cols]

    # Remove duplicate columns if any
    if 'ანგარიში +' in df.columns:
        df = df.loc[:, ~df.columns.duplicated()]

    # Convert all possible columns to int where feasible (skipping first column)
    for col in df.columns[1:]:
        try:
            df[col] = pd.to_numeric(df[col], errors='ignore').round().astype(int)
        except ValueError:
            pass

    # Add final 'Calculated' column
    df['Calculated'] = df.apply(calculate_value, axis=1)

    # Now apply thousands-separator formatting to all numeric columns except the first column
    for i, col in enumerate(df.columns):
        if i == 0:
            # The very first 4-digit column - leave it alone
            continue

        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].apply(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else "")

    return df


def sum_by_mapping(df):
    """
    Takes a DataFrame that has 'Mapping' and 'Calculated' columns.
    Sums the 'Calculated' values for each unique 'Mapping', and then
    returns a custom-ordered financial statement-like Styler object
    so that totals appear in bold in a Jupyter/HTML environment.
    """
    # Convert 'Calculated' back to numeric (remove commas if present)
    df['Calculated'] = df['Calculated'].replace(",", "", regex=True)  # remove commas
    df['Calculated'] = pd.to_numeric(df['Calculated'], errors='coerce')

    grouped_df = df.groupby('Mapping', as_index=False)['Calculated'].sum()

    # Build the final custom-ordered table as a plain DataFrame
    final_df = build_custom_report(grouped_df)

    # Return a Styler that will highlight rows containing 'Total' in bold
    return highlight_totals_in_bold(final_df)


def build_custom_report(grouped_df):
    """
    Given a DataFrame with columns ['Mapping', 'Calculated'],
    produce a final table in the exact order requested, with
    totals and placeholders (#REF!) where needed. Returns a
    plain DataFrame.
    """
    # Convert grouped results to a lookup dict
    sum_dict = dict(zip(grouped_df['Mapping'], grouped_df['Calculated']))

    row_order = [
        "Cash & Cash equivalents",
        "Accounts Receivables",
        "Tax assets",
        "Inventories",
        "Advances Paid",
        "Other current assets",
        "Total Current Assets",
        "Net PPE",
        "Net Intangible Assets",
        "Other fixed assets",
        "Total Fixed Assets",
        "Total Assets",
        "Accounts Payables",
        "Salary Payables",
        "Short-Term Loans",
        "Taxes Payable",
        "Interest Payable",
        "Other Short term liabilities",
        "Total current Liabilities",
        "Long-term loan",
        "Other long term liabilities",
        "Total long term Liabilities",
        "Total Liabilities",
        "Share Capital",
        "Retained Earning",
        "Other reserves",
        "Shareholder equity",
        "Liabilities & Equity",
        "Check"
    ]

    # These remain placeholders for now
    ref_rows = {"Retained Earning", "Shareholder equity", "Liabilities & Equity", "Check"}
    final_vals = {}
    for item in row_order:
        if item in ref_rows:
            final_vals[item] = 0
        else:
            final_vals[item] = sum_dict.get(item, 0.0)

    # Helper function
    def safe_sum(keys):
        total = 0.0
        for k in keys:
            val = final_vals.get(k, 0.0)
            if isinstance(val, (int, float)):
                total += val
        return total

    # Compute the totals
    final_vals["Total Current Assets"] = safe_sum([
        "Cash & Cash equivalents",
        "Accounts Receivables",
        "Tax assets",
        "Inventories",
        "Advances Paid",
        "Other current assets"
    ])
    final_vals["Total Fixed Assets"] = safe_sum([
        "Net PPE",
        "Net Intangible Assets",
        "Other fixed assets"
    ])
    final_vals["Total Assets"] = safe_sum([
        "Total Current Assets",
        "Total Fixed Assets"
    ])
    final_vals["Total current Liabilities"] = safe_sum([
        "Accounts Payables",
        "Salary Payables",
        "Short-Term Loans",
        "Taxes Payable",
        "Interest Payable",
        "Other Short term liabilities"
    ])
    final_vals["Total long term Liabilities"] = safe_sum([
        "Long-term loan",
        "Other long term liabilities"
    ])
    final_vals["Total Liabilities"] = safe_sum([
        "Total current Liabilities",
        "Total long term Liabilities"
    ])

    # Build the final DataFrame
    rows = []
    for item in row_order:
        rows.append({
            "Mapping": item,
            "Calculations": final_vals[item]
        })

    final_df = pd.DataFrame(rows, columns=["Mapping", "Calculations"])
    # Return a plain DataFrame (numeric in the second column)
    return final_df


def highlight_totals_in_bold(df):
    """
    Takes a plain DataFrame with columns ['Mapping', 'Calculations'],
    and returns a Styler that:
      - Bolds any row whose 'Mapping' contains 'Total'
      - Applies thousands separators to 'Calculations' (numeric)
    """
    def bold_totals(row):
        if 'Total' in str(row['Mapping']):
            # Return a style for each column in this row
            return ['font-weight: bold'] * len(row)
        else:
            return [''] * len(row)

    # Create a styler
    styler = df.style.apply(bold_totals, axis=1)

    # Format the numeric values in the 'Calculations' column with commas
    # (only if they are numeric, placeholders or zero remain as is if not numeric)
    styler.format("{:,.0f}", subset=['Calculations'])

    return styler
