import pandas as pd
import re
from tabulate import tabulate

# --- Replace with your actual Excel file path ---
file_path = "L2 Platform Support 2025 7-22-2025 8-11-17 PM.xlsx"

# Load data
df = pd.read_excel(file_path, sheet_name='L2 Platform Support 2025')

# Parse datetime columns
df['Entered Queue'] = pd.to_datetime(df['Entered Queue'], errors='coerce')
df['Created On'] = pd.to_datetime(df['Created On'], errors='coerce')


# Normalize case titles: lowercase, trim, collapse spaces, normalize +
def normalize_title(title):
    if pd.isna(title):
        return ""
    title = title.lower().strip()
    title = re.sub(r'\s+', ' ', title)
    title = re.sub(r'\s*\+\s*', '+', title)
    return title


df['Normalized Title'] = df['Title'].apply(normalize_title)

# --- Replace blank priority with "Normal"
df['Priority'] = df['Priority'].fillna('Normal')

# 1. Platform counts
platform_counts_df = df['Platform'].value_counts(dropna=False).reset_index()
platform_counts_df.columns = ['Platform', 'Case Count']
platform_counts_df.index += 1

# 2. Top 10 standardized case titles
common_case_titles_df = df['Normalized Title'].value_counts().head(10).reset_index()
common_case_titles_df.columns = ['Case Title (Standardized)', 'Case Count']
common_case_titles_df.index += 1

# 3. Cases by team member
cases_by_member_df = df['Worked By'].value_counts().reset_index()
cases_by_member_df.columns = ['Team Member', 'Case Count']
cases_by_member_df.index += 1

# 4. Cases by priority (blanks treated as "Normal")
cases_by_priority_df = df['Priority'].value_counts().reset_index()
cases_by_priority_df.columns = ['Priority', 'Case Count']
cases_by_priority_df.index += 1

# 5. Top 10 days with most cases entered
top_days_df = df['Entered Queue'].dt.date.value_counts().head(10).sort_index().reset_index()
top_days_df.columns = ['Date', 'Case Count']
top_days_df.index += 1

# 6. Count of missing platform values
missing_platform_count = df['Platform'].isna().sum()

# 7. Peak hour distribution (leave Hour as is, 0–23)
peak_hours_df = df['Entered Queue'].dt.hour.value_counts().sort_index().reset_index()
peak_hours_df.columns = ['Hour', 'Cases Entered']


# Function to display tables with left-aligned text and right-aligned numbers
def print_table(df, title):
    print(f"\n{title}")
    print(tabulate(df, headers='keys', showindex=True, tablefmt='pretty', stralign='left', numalign='right'))


# --- Print all results cleanly ---
print_table(platform_counts_df, "1. Platform Counts")
print_table(common_case_titles_df, "2. Top Standardized Case Titles")
print_table(cases_by_member_df, "3. Cases by Team Member")
print_table(cases_by_priority_df, "4. Cases by Priority")
print_table(top_days_df, "5. Top 10 Days with Most Cases Entered")

print(f"\n6. Missing Platform Count:\n{missing_platform_count}")

print_table(peak_hours_df, "7. Peak Hour Distribution (Entered Queue)")

