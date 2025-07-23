import pandas as pd
import re

# Load the Excel file
file_path = "L2 Platform Support 2025 7-22-2025 8-11-17 PM.xlsx"
df = pd.read_excel(file_path, sheet_name='L2 Platform Support 2025')

# Convert date/time columns
df['Entered Queue'] = pd.to_datetime(df['Entered Queue'], errors='coerce')
df['Created On'] = pd.to_datetime(df['Created On'], errors='coerce')

# Normalize Title (case-insensitive, space-insensitive, "+" spacing normalized)
def normalize_title(title):
    if pd.isna(title):
        return ""
    title = title.lower().strip()
    title = re.sub(r'\s+', ' ', title)              # Collapse multiple spaces
    title = re.sub(r'\s*\+\s*', '+', title)         # Normalize spacing around "+"
    return title

df['Normalized Title'] = df['Title'].apply(normalize_title)

# 1. Platform types and their frequencies
platform_counts_df = df['Platform'].value_counts(dropna=False).reset_index()
platform_counts_df.columns = ['Platform', 'Case Count']

# 2. Top 10 standardized case types
common_case_titles_df = df['Normalized Title'].value_counts().head(10).reset_index()
common_case_titles_df.columns = ['Case Title (Standardized)', 'Case Count']

# 3. Case quantity worked by each team member
cases_by_member_df = df['Worked By'].value_counts().reset_index()
cases_by_member_df.columns = ['Team Member', 'Case Count']

# 4. Case quantity by priority
cases_by_priority_df = df['Priority'].value_counts(dropna=False).reset_index()
cases_by_priority_df.columns = ['Priority', 'Case Count']

# 5. Top 10 days with the most cases entered queue
top_days_df = df['Entered Queue'].dt.date.value_counts().head(10).sort_index().reset_index()
top_days_df.columns = ['Date', 'Case Count']

# 6. Number of cases missing platform info
missing_platform_count = df['Platform'].isna().sum()

# 7. Peak hour periods for case entries
peak_hours = df['Entered Queue'].dt.hour.value_counts().sort_index()
peak_hours_df = peak_hours.reset_index(drop=False)
peak_hours_df.columns = ['Hour', 'Cases Entered']

# --- Display Results ---
print("\n1. Platform Counts:")
print(platform_counts_df)

print("\n2. Top Standardized Case Titles:")
print(common_case_titles_df)

print("\n3. Cases by Team Member:")
print(cases_by_member_df)

print("\n4. Cases by Priority:")
print(cases_by_priority_df)

print("\n5. Top 10 Days with Most Cases Entered:")
print(top_days_df)

print("\n6. Missing Platform Count:")
print(missing_platform_count)

print("\n7. Peak Hour Distribution (Entered Queue):")
print(peak_hours_df)
