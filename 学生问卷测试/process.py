import pandas as pd
import numpy as np

# Replace 'your_file.xlsx' with the path to your Excel file
df = pd.read_excel('datastu.xlsx')
#print(df)
# Identify multi-choice columns (you may need to adjust this based on your data)
multi_choice_columns = [11,12,13,15,17,19,20,23]  # Replace with actual multi-choice column names

# Convert multi-choice columns into binary columns for each choice
for col in multi_choice_columns:
    choices = set(''.join(df[col].dropna().astype(str).str.replace('[^0-9]', '').unique())) 
    for choice in choices:
        df[f'{col}_choice_{choice}'] = df[col].astype(str).apply(lambda x: 1 if choice in x else 0)

# Drop the original multi-choice columns
df = df.drop(columns=multi_choice_columns)

# Convert all columns to numeric, coercing errors to NaN
df = df.apply(pd.to_numeric, errors='coerce')

# Initialize a list to store valid correlation pairs
valid_corr_pairs = []

# Calculate the correlation matrix, skipping pairs with NaNs
for i in range(df.shape[1]):
    for j in range(i+1, df.shape[1]):
        col1 = df.iloc[:, i]
        col2 = df.iloc[:, j]
        try:
            correlation = col1.corr(col2)
            if not np.isnan(correlation):
                valid_corr_pairs.append((df.columns[i], df.columns[j], correlation))
        except Exception as e:
            continue

# Convert the list to a DataFrame
corr_pairs_df = pd.DataFrame(valid_corr_pairs, columns=['Index1', 'Index2', 'Correlation'])

# Sort the pairs by their correlation coefficients in descending order
sorted_corr_pairs_df = corr_pairs_df.sort_values(by='Correlation', ascending=False)

# Display the top pairs
print(sorted_corr_pairs_df)
sorted_corr_pairs_df.to_excel("correlation.xlsx")