import pandas as pd
import re
from datetime import timedelta
from tabulate import tabulate

# --- Replace with your actual Excel file path ---
file_path = "L2 Platform Support 2025 8-4-2025 12-02-31 PM.xlsx"

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

# --- Get top issues for Normal priority (including blanks)
normal_issues = df[df['Priority'] == 'Normal']
normal_common_titles = normal_issues['Normalized Title'].value_counts().head(10).reset_index()
normal_common_titles.columns = ['Normalized Title', 'Case Count']

# --- Get top issues for High priority
high_issues = df[df['Priority'] == 'High']
high_common_titles = high_issues['Normalized Title'].value_counts().head(10).reset_index()
high_common_titles.columns = ['Normalized Title', 'Case Count']

# 5. Top 10 days with most cases entered
top_days_df = df['Entered Queue'].dt.date.value_counts().head(10).reset_index()
top_days_df.columns = ['Date', 'Case Count']
top_days_df.index += 1

# Get the latest date in the dataset
latest_date = df['Entered Queue'].max().date()

# Define the 10-day window (inclusive of latest date)
start_date = latest_date - timedelta(days=9)

# Filter rows within the last 10 days
recent_cases = df[df['Entered Queue'].dt.date.between(start_date, latest_date)]

# 1. Total count
last_10_days_count = recent_cases.shape[0]
print(f"\n8. Total Case Quantity in the Last 10 Days ({start_date} to {latest_date}): {last_10_days_count}")

last_10_days_df = recent_cases['Entered Queue'].dt.date.value_counts().sort_index(ascending=False).reset_index()
last_10_days_df.columns = ['Date', 'Case Count']
last_10_days_df.index += 1


# 6. Count of missing platform values
missing_platform_count = df['Platform'].isna().sum()

# 7. Peak hour distribution (leave Hour as is, 0–23)
peak_hours_df = df['Entered Queue'].dt.hour.value_counts().sort_index().reset_index()
peak_hours_df.columns = ['Hour', 'Cases Entered']


# Function to display tables with left-aligned text and right-aligned numbers
def print_table(df, title, show_index=True):
    print(f"\n{title}")
    print(tabulate(df, headers='keys', showindex=show_index, tablefmt='pretty', stralign='left', numalign='right'))


# --- Print all results cleanly ---
print_table(platform_counts_df, "1. Platform Counts")
print(f"\n2. Cases without Platform :\n{missing_platform_count}")
print_table(common_case_titles_df, "3. Most Common Case Titles")
print_table(cases_by_member_df, "4. Cases by Team Member")
print_table(cases_by_priority_df, "5. Cases by Priority")
print_table(normal_common_titles, "6. Top 10 Most Common Issues - Normal Priority (including empty)")
print_table(high_common_titles, "7. Top 10 Most Common Issues - High Priority")
print_table(top_days_df, "8. Top 10 Days with Most Cases Entered")
print_table(last_10_days_df, f"9. Daily Case Counts in the Last 10 Days ({start_date} to {latest_date})")
print(f"\n10. Total Case Quantity in the Last 10 Days ({start_date} to {latest_date}): {last_10_days_count}")
print_table(peak_hours_df, "11. Peak Hour Distribution (Entered Queue)", show_index=False)
