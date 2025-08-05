import pandas as pd
import re
from datetime import timedelta
from tabulate import tabulate

file_path = "L2 Platform Support 2025 8-5-2025 10-42-33 AM.xlsx"

df = pd.read_excel(file_path, sheet_name='L2 Platform Support 2025')

# Process datetime
df['Entered Queue'] = pd.to_datetime(df['Entered Queue'], errors='coerce')
df['Created On'] = pd.to_datetime(df['Created On'], errors='coerce')
df['Modified On'] = pd.to_datetime(df['Modified On'], errors='coerce')


# Trim case title
def normalize_title(title):
    if pd.isna(title):
        return ""
    title = title.lower().strip()
    title = re.sub(r'\s+', ' ', title)
    title = re.sub(r'\s*\+\s*', '+', title)
    return title


df['Normalized Title'] = df['Title'].apply(normalize_title)


# 1. Case count by platform
df['Platform'] = df['Platform'].fillna('Empty')
platform_counts_df = df['Platform'].value_counts(dropna=False).reset_index()
platform_counts_df.columns = ['Platform', 'Case Count']
platform_counts_df.index += 1

# 2. Top 10 case types
common_case_titles_df = df['Normalized Title'].value_counts().head(10).reset_index()
common_case_titles_df.columns = ['Case Title (Standardized)', 'Case Count']
common_case_titles_df.index += 1

# 3. Cases worked by team member
cases_by_member_df = df['Worked By'].value_counts().reset_index()
cases_by_member_df.columns = ['Team Member', 'Case Count']
cases_by_member_df.index += 1

# 4. Case count by priority
cases_by_priority_df = df['Priority'].value_counts().reset_index()
cases_by_priority_df.columns = ['Priority', 'Case Count']
cases_by_priority_df.index += 1

# 5. Top issues for Normal priority (including blanks)
normal_issues = df[df['Priority'] == 'Normal']
normal_common_titles = normal_issues['Normalized Title'].value_counts().head(10).reset_index()
normal_common_titles.columns = ['Normalized Title', 'Case Count']

# 6. Top issues for High priority
high_issues = df[df['Priority'] == 'High']
high_common_titles = high_issues['Normalized Title'].value_counts().head(10).reset_index()
high_common_titles.columns = ['Normalized Title', 'Case Count']

# 7. Top 10 days with most cases entered
top_days_df = df['Entered Queue'].dt.date.value_counts().head(10).reset_index()
top_days_df.columns = ['Date', 'Case Count']
top_days_df.index += 1

# Get the latest date in the dataset
latest_date = df['Entered Queue'].max().date()

# Get 10-day window (until latest date)
start_date = latest_date - timedelta(days=9)

# Filter rows within the last 10 days
recent_cases = df[df['Entered Queue'].dt.date.between(start_date, latest_date)]

# 8. Total count
last_10_days_count = recent_cases.shape[0]

last_10_days_df = recent_cases['Entered Queue'].dt.date.value_counts().sort_index(ascending=False).reset_index()
last_10_days_df.columns = ['Date', 'Case Count']
last_10_days_df.index += 1

# 10. Peak hour distribution
peak_hours_df = df['Entered Queue'].dt.hour.value_counts().sort_index().reset_index()
peak_hours_df.columns = ['Hour', 'Cases Entered']

# Filter for "Inactive" status
inactive_cases = df[df['Status'] == 'Inactive'].copy()

# Calculate resolution time
inactive_cases['Average Resolution Time'] = inactive_cases['Modified On'] - inactive_cases['Entered Queue']
inactive_cases['Resolution Hours'] = inactive_cases['Average Resolution Time'].dt.total_seconds() / 3600
inactive_cases['Resolution Days'] = inactive_cases['Resolution Hours'] / 24


# 11. Average resolved time for all cases
avg_resolved_time = inactive_cases['Average Resolution Time'].mean()

# 12. Average resolved time by platform
avg_by_platform = inactive_cases.groupby('Platform')['Average Resolution Time'].mean().reset_index()
avg_by_platform_sorted = avg_by_platform.sort_values(by='Average Resolution Time')

# 13. Average resolved time by team member
avg_by_member = inactive_cases.groupby('Worked By')['Average Resolution Time'].mean().reset_index()
avg_by_member_sorted = avg_by_member.sort_values(by='Average Resolution Time')

# 14. Average resolved time of Normal priority cases
avg_normal_priority = inactive_cases[inactive_cases['Priority'] == 'Normal']['Average Resolution Time'].mean()

# 15. Average resolved time of High priority cases
avg_high_priority = inactive_cases[inactive_cases['Priority'] == 'High']['Average Resolution Time'].mean()

# 16. Duration ranges
under_12_hours = inactive_cases[inactive_cases['Resolution Hours'] <= 12].shape[0]
between_12_24_hours = inactive_cases[(inactive_cases['Resolution Hours'] > 12) & (inactive_cases['Resolution Hours'] <= 24)].shape[0]
between_1_3_days = inactive_cases[(inactive_cases['Resolution Days'] > 1) & (inactive_cases['Resolution Days'] <= 3)].shape[0]
between_3_7_days = inactive_cases[(inactive_cases['Resolution Days'] > 3) & (inactive_cases['Resolution Days'] <= 7)].shape[0]
over_7_days = inactive_cases[inactive_cases['Resolution Days'] > 7].shape[0]


# Display tables
def print_table(df, title, show_index=True, colalign=None):
    print(f"\n{title}")
    print(tabulate(
        df,
        headers='keys',
        showindex=show_index,
        tablefmt='pretty',
        stralign='left',
        numalign='right',
        colalign=colalign  # Custom alignment
    ))


# Print results
print_table(platform_counts_df, "1. Case Count by Platform")
print_table(common_case_titles_df, "2. Most Common Case Titles")
print_table(cases_by_member_df, "3. Cases by Team Member")
print_table(cases_by_priority_df, "4. Cases by Priority")
print_table(normal_common_titles, "5. Top 10 Most Common Issues - Normal Priority (including empty)")
print_table(high_common_titles, "6. Top 10 Most Common Issues - High Priority")
print_table(top_days_df, "7. Top 10 Days with Most Cases Entered")
print_table(last_10_days_df, f"8. Daily Case Counts in the Last 10 Days ({start_date} to {latest_date})")
print(f"\n9. Total Case Quantity in the Last 10 Days ({start_date} to {latest_date}): {last_10_days_count}")
print_table(peak_hours_df, "10. Peak Hour Distribution (Entered Queue)", show_index=False)

print("11. Average resolved time for all cases:", avg_resolved_time)
print_table(avg_by_platform_sorted, "2. Average Resolved Time by Platform", show_index=False, colalign=("left", "right"))
print_table(avg_by_member_sorted, "3. Average Resolved Time by Team Member", show_index=False, colalign=("left", "right"))
print("\n14. Average resolved time of Normal priority cases:", avg_normal_priority)
print("15. Average resolved time of High priority cases:", avg_high_priority)
print("16. Number of cases resolved under 12 hours:", under_12_hours)
print("17. Number of cases resolved between 12 - 24 hours:", between_12_24_hours)
print("18. Number of cases resolved between 1 - 3 days:", between_1_3_days)
print("19. Number of cases resolved between 3 - 7 days:", between_3_7_days)
print("20. Number of cases resolved in over 7 days:", over_7_days)
