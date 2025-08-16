import pandas as pd
import re
from datetime import timedelta
from tabulate import tabulate

file_path = "L2 Platform Support Data.xlsx"

df = pd.read_excel(file_path, sheet_name='Sheet1')

# Process datetime
df['Entered Queue'] = pd.to_datetime(df['Entered Queue'], errors='coerce')
df['Resolution Date'] = pd.to_datetime(df['Resolution Date'], errors='coerce')

# Convert PST to EST
time_columns = ['Entered Queue', 'Resolution Date']
for col in time_columns:
    df[col] = df[col] + pd.Timedelta(hours=3)


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
df['Platform'] = df['Platform'].fillna('Other (Test cases/Undefined)')
platform_counts_df = df['Platform'].value_counts(dropna=False).reset_index()
platform_counts_df.columns = ['Platform', 'Case Count']
platform_counts_df.index += 1

# 1.1. Top 5 most common subjects per platform
platform_subject_counts = df.groupby(['Platform', 'Subject']).size().reset_index(name='Case Count')

# Sort and take top 5 per platform
top5_per_platform = (
    platform_subject_counts
    .sort_values(['Platform', 'Case Count'], ascending=[True, False])
    .groupby('Platform')
    .head(5)
)

# 1.2 Top 10 Customers per Platform
platform_customer_counts = df.groupby(['Platform', 'Customer']).size().reset_index(name='Case Count')

# Sort and take top 10 per platform
top10_per_platform = (
    platform_customer_counts
    .sort_values(['Platform', 'Case Count'], ascending=[True, False])
    .groupby('Platform')
    .head(10)
)

# 2. Top 10 case types
common_case_titles_df = df['Normalized Title'].value_counts().head(10).reset_index()
common_case_titles_df.columns = ['Case Title (Standardized)', 'Case Count']
common_case_titles_df.index += 1

# 3. Cases worked by team member
cases_by_member_df = df['Worked By'].value_counts().reset_index()
cases_by_member_df.columns = ['Team Member', 'Case Count']
cases_by_member_df.index += 1

# 4. Case count by priority
df['Priority'] = df['Priority'].fillna('Normal')
cases_by_priority_df = df['Priority'].value_counts().reset_index()
cases_by_priority_df.columns = ['Priority', 'Case Count']
cases_by_priority_df.index += 1

# 5. Top issues for Normal priority (including blanks)
normal_issues = df[df['Priority'].isna() | (df['Priority'] == 'Normal')]
normal_common_titles = normal_issues['Normalized Title'].value_counts().head(10).reset_index()
normal_common_titles.columns = ['Normalized Title', 'Case Count']
normal_common_titles.index += 1

# 6. Top issues for High priority
high_issues = df[df['Priority'] == 'High']
high_common_titles = high_issues['Normalized Title'].value_counts().head(10).reset_index()
high_common_titles.columns = ['Normalized Title', 'Case Count']
high_common_titles.index += 1

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

# 8. Average cases per week day
df['Day of Week'] = df['Entered Queue'].dt.day_name()

# Count cases per date and day of week
cases_per_day = df.groupby([df['Entered Queue'].dt.date, 'Day of Week']).size().reset_index(name='Case Count')

# Average case count by day of week, rounded to nearest whole number
avg_cases_by_dow = (
    cases_per_day.groupby('Day of Week')['Case Count']
    .mean()
    .round(0)
    .astype(int)  # convert to integer
    .reindex(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])
    .reset_index()
)


# 10. Peak hour distribution
peak_hours_df = df['Entered Queue'].dt.hour.value_counts().sort_index().reset_index()
peak_hours_df.columns = ['Hour', 'Cases Entered']

# Filter cases that have a resolution date
resolved_cases = df[df['Resolution Date'].notna()].copy()

# Calculate resolution time
resolved_cases['Average Resolution Time'] = resolved_cases['Resolution Date'] - resolved_cases['Entered Queue']
resolved_cases['Resolution Hours'] = resolved_cases['Average Resolution Time'].dt.total_seconds() / 3600
resolved_cases['Resolution Days'] = resolved_cases['Resolution Hours'] / 24


# 11. Average resolved time for all cases
avg_resolved_time = resolved_cases['Average Resolution Time'].mean()

# 12. Average resolved time by platform
avg_by_platform = resolved_cases.groupby('Platform')['Average Resolution Time'].mean().reset_index()
avg_by_platform_sorted = avg_by_platform.sort_values(by='Average Resolution Time')

# 13. Average resolved time by team member
avg_by_member = resolved_cases.groupby('Worked By')['Average Resolution Time'].mean().reset_index()
avg_by_member_sorted = avg_by_member.sort_values(by='Average Resolution Time')

# 14. Average resolved time of Normal priority cases
avg_normal_priority = resolved_cases[resolved_cases['Priority'].isna() | (resolved_cases['Priority'] == 'Normal')]['Average Resolution Time'].mean()

# 15. Average resolved time of High priority cases
avg_high_priority = resolved_cases[resolved_cases['Priority'] == 'High']['Average Resolution Time'].mean()

# 16. Duration ranges
under_12_hours = resolved_cases[resolved_cases['Resolution Hours'] <= 12].shape[0]
between_12_24_hours = resolved_cases[(resolved_cases['Resolution Hours'] > 12) & (resolved_cases['Resolution Hours'] <= 24)].shape[0]
between_1_3_days = resolved_cases[(resolved_cases['Resolution Days'] > 1) & (resolved_cases['Resolution Days'] <= 3)].shape[0]
between_3_7_days = resolved_cases[(resolved_cases['Resolution Days'] > 3) & (resolved_cases['Resolution Days'] <= 7)].shape[0]
over_7_days = resolved_cases[resolved_cases['Resolution Days'] > 7].shape[0]


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


# 1. Case count by platform
print_table(platform_counts_df, "1. Case Count by Platform")
# 1.1 Print separate table for each platform
for platform, table in top5_per_platform.groupby('Platform'):
    print_table(table.reset_index(drop=True), f"Top 5 Subjects - {platform}", show_index=False)
# 1.2
for platform, table in top10_per_platform.groupby('Platform'):
    print_table(table.reset_index(drop=True), f"Top 10 Customers - {platform}", show_index=False)
print_table(common_case_titles_df, "2. Most Common Case Titles")
print_table(cases_by_member_df, "3. Cases Count by Team Member")
print_table(cases_by_priority_df, "4. Case Count by Priority")
print_table(normal_common_titles, "5. Top 10 Most Common Issues - Normal Priority")
print_table(high_common_titles, "6. Top 10 Most Common Issues - High Priority")
print_table(top_days_df, "7. Top 10 Days with Most Cases Entered Queue")
print_table(avg_cases_by_dow, "8. Average Case Count by Day of Week", show_index=False)

print_table(peak_hours_df, "10. Case Entered Queue by each hour (Vancouver time)", show_index=False)

print("11. Average resolved time for all cases:", avg_resolved_time)
print_table(avg_by_platform_sorted, "12. Average Resolved Time by Platform", show_index=False, colalign=("left", "right"))
print_table(avg_by_member_sorted, "13. Average Resolved Time by Team Member", show_index=False, colalign=("left", "right"))
print("\n14. Average resolved time of Normal priority cases:", avg_normal_priority)
print("15. Average resolved time of High priority cases:", avg_high_priority)
print("16. Number of cases resolved under 12 hours:", under_12_hours)
print("17. Number of cases resolved between 12 - 24 hours:", between_12_24_hours)
print("18. Number of cases resolved between 1 - 3 days:", between_1_3_days)
print("19. Number of cases resolved between 3 - 7 days:", between_3_7_days)
print("20. Number of cases resolved in over 7 days:", over_7_days)
