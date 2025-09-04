import pandas as pd

# Load the Excel file
xl = pd.ExcelFile('databook.xlsx')
print('Sheets:', xl.sheet_names)

# Examine first few sheets to understand structure
for sheet in xl.sheet_names[:5]:
    print(f'\n=== {sheet} ===')
    df = pd.read_excel('databook.xlsx', sheet_name=sheet)
    print(f'Shape: {df.shape}')
    print('Columns:', df.columns.tolist()[:15])

    # Check for entity-related columns or multi-entity structure
    if len(df.columns) > 0:
        print('First few rows:')
        print(df.head(3))

        # Look for columns that might indicate multiple entities
        entity_cols = [col for col in df.columns if 'entity' in col.lower() or 'company' in col.lower()]
        if entity_cols:
            print(f'Potential entity columns: {entity_cols}')

        # Check if there are multiple tables or sections
        print('Unique values in first column:', df.iloc[:, 0].unique()[:10])
