import pandas as pd
import re
import math
from datetime import timedelta
from tabulate import tabulate


file_path = "L2 Platform Support Master Data.xlsx"

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
df['Platform'] = df['Platform'].fillna('Other')
platform_counts_df = df['Platform'].value_counts(dropna=False).reset_index()
platform_counts_df.columns = ['Platform', 'Case Count']
platform_counts_df.index += 1

# 2. Top 5 most common subjects per platform
platform_subject_counts = df.groupby(['Platform', 'Subject']).size().reset_index(name='Case Count')

# Sort and take top 5 per platform
top5_per_platform = (
    platform_subject_counts
    .sort_values(['Platform', 'Case Count'], ascending=[True, False])
    .groupby('Platform')
    .head(5)
)

# 3. Top 10 Customers per Platform
platform_customer_counts = df.groupby(['Platform', 'Customer']).size().reset_index(name='Case Count')

# Sort and take top 10 per platform
top10_per_platform = (
    platform_customer_counts
    .sort_values(['Platform', 'Case Count'], ascending=[True, False])
    .groupby('Platform')
    .head(10)
)

# 4. Cases worked by team member
cases_by_member_df = df['Worked By'].value_counts().reset_index()
cases_by_member_df.columns = ['Team Member', 'Case Count']
cases_by_member_df.index += 1

# 5. Case count by priority
df['Priority'] = df['Priority'].fillna('Normal')
cases_by_priority_df = df['Priority'].value_counts().reset_index()
cases_by_priority_df.columns = ['Priority', 'Case Count']
cases_by_priority_df.index += 1

# Group by Priority + Subject
priority_subject_counts = (
    df.groupby(['Priority', 'Subject'])
      .size()
      .reset_index(name='Case Count')
)

# Sort and take top 5 per priority
top5_subjects_per_priority = (
    priority_subject_counts
    .sort_values(['Priority', 'Case Count'], ascending=[True, False])
    .groupby('Priority')
    .head(5)
)

# 9. Top 10 days with most cases entered queue
top_days_df = df['Entered Queue'].dt.date.value_counts().head(10).reset_index()
top_days_df.columns = ['Date', 'Case Count']
top_days_df.index += 1

# Get the latest date in the dataset
latest_date = df['Entered Queue'].max().date()

# 10. Average cases per week day
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


# 11. Peak hour distribution
peak_hours_df = df['Entered Queue'].dt.hour.value_counts().sort_index().reset_index()
peak_hours_df.columns = ['Hour', 'Cases Entered']

# Filter cases that have a resolution date
resolved_cases = df[df['Resolution Date'].notna()].copy()

# Calculate resolution time
resolved_cases['Average Resolution Time'] = resolved_cases['Resolution Date'] - resolved_cases['Entered Queue']
resolved_cases['Resolution Hours'] = resolved_cases['Average Resolution Time'].dt.total_seconds() / 3600
resolved_cases['Resolution Days'] = resolved_cases['Resolution Hours'] / 24

# Round resolution time columns to full seconds
resolved_cases['Average Resolution Time'] = resolved_cases['Average Resolution Time'].dt.round('1s')

# 12.1. Average resolved time for all cases
avg_resolved_time = resolved_cases['Average Resolution Time'].mean()

# 12.2. Average resolved time of Normal priority cases
avg_normal_priority = resolved_cases[resolved_cases['Priority'].isna() | (resolved_cases['Priority'] == 'Normal')]['Average Resolution Time'].mean()

# 12.3. Average resolved time of High priority cases
avg_high_priority = resolved_cases[resolved_cases['Priority'] == 'High']['Average Resolution Time'].mean()

# 13. Average resolved time by platform
avg_by_platform = resolved_cases.groupby('Platform')['Average Resolution Time'].mean().reset_index()
avg_by_platform_sorted = avg_by_platform.sort_values(by='Average Resolution Time')

# 14. Average resolved time by team member
avg_by_member = resolved_cases.groupby('Worked By')['Average Resolution Time'].mean().reset_index()
avg_by_member_sorted = avg_by_member.sort_values(by='Average Resolution Time')

# 15. Duration ranges
under_12_hours = resolved_cases[resolved_cases['Resolution Hours'] <= 12].shape[0]
between_12_24_hours = resolved_cases[(resolved_cases['Resolution Hours'] > 12) & (resolved_cases['Resolution Hours'] <= 24)].shape[0]
between_1_3_days = resolved_cases[(resolved_cases['Resolution Days'] > 1) & (resolved_cases['Resolution Days'] <= 3)].shape[0]
between_3_7_days = resolved_cases[(resolved_cases['Resolution Days'] > 3) & (resolved_cases['Resolution Days'] <= 7)].shape[0]
over_7_days = resolved_cases[resolved_cases['Resolution Days'] > 7].shape[0]

resolution_ranges = pd.DataFrame([
    {"Resolution time": "Under 12 hours", "Case Count": under_12_hours},
    {"Resolution time": "12 - 24 hours", "Case Count": between_12_24_hours},
    {"Resolution time": "1 - 3 days", "Case Count": between_1_3_days},
    {"Resolution time": "3 - 7 days", "Case Count": between_3_7_days},
    {"Resolution time": "Over 7 days", "Case Count": over_7_days},
])

# Calculate percentages (1 decimal place)
total_cases = resolution_ranges["Case Count"].sum()
resolution_ranges["Percentage"] = (
    resolution_ranges["Case Count"] / total_cases * 100
).round(1).astype(str) + "%"

# Add total row
total_row = pd.DataFrame([{
    "Resolution time": "Total",
    "Case Count": total_cases,
    "Percentage": "100%"
}])

resolution_summary = pd.concat([resolution_ranges, total_row], ignore_index=True)

# 16. Filter only escalated cases (where Escalated == "Yes")
escalated_cases = df[df['Escalated'].astype(str).str.strip().str.lower() == 'yes']

# Count number of escalated cases per Subject
subject_escalated_counts = escalated_cases['Subject'].value_counts().reset_index()
subject_escalated_counts.columns = ['Subject', 'Escalated Case Count']

# Average resolved time for escalated cases
escalated_resolved_cases = resolved_cases[
    resolved_cases['Escalated'].astype(str).str.strip().str.lower() == 'yes'
].copy()

if not escalated_resolved_cases.empty:
    avg_escalated_time = escalated_resolved_cases['Average Resolution Time'].mean()
    # Round up to seconds
    total_seconds = math.ceil(avg_escalated_time.total_seconds())
    avg_escalated_time = timedelta(seconds=total_seconds)
else:
    avg_escalated_time = None


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


def format_timedelta(td):
    if pd.isna(td):
        return "N/A"
    total_seconds = int(td.total_seconds())
    return str(timedelta(seconds=total_seconds))


# 1. Case count by platform
print_table(platform_counts_df, "1. Case Count by Platform")
# 2. Top 5 Subjects per Platform
print("2. Top 5 Subjects per Platform")
for platform, table in top5_per_platform.groupby('Platform'):
    print_table(table.reset_index(drop=True), f"Top 5 Subjects - {platform}", show_index=False)
# 3. Top 10 Customers per Platform
print("3. Top 10 Customers per Platform")
for platform, table in top10_per_platform.groupby('Platform'):
    print_table(table.reset_index(drop=True), f"Top 10 Customers - {platform}", show_index=False)

print_table(cases_by_member_df, "4. Cases Count by Team Member")
print_table(cases_by_priority_df, "5. Case Count by Priority")
print("6-7-8. Top 5 Subjects per Priority")
for priority, table in top5_subjects_per_priority.groupby('Priority'):
    print_table(
        table.reset_index(drop=True),
        f"Top 5 Subjects - Priority: {priority}",
        show_index=False
    )

print_table(top_days_df, "9. Top 10 Days with Most Cases Entered Queue")
print_table(avg_cases_by_dow, "10. Average Case Count by Day of Week", show_index=False)
print_table(peak_hours_df, "11. Case Entered Queue by each hour (EST)", show_index=False)

print("12. Average resolved time for all cases:", format_timedelta(avg_resolved_time))
print("12. Average resolved time of Normal priority cases:", format_timedelta(avg_normal_priority))
print("13. Average resolved time of High priority cases:", format_timedelta(avg_high_priority))
print_table(avg_by_platform_sorted.assign(
    **{'Average Resolution Time': avg_by_platform_sorted['Average Resolution Time'].apply(format_timedelta)}),
    "13. Average Resolved Time by Platform", show_index=False, colalign=("left", "right"))
print_table(avg_by_member_sorted.assign(
    **{'Average Resolution Time': avg_by_member_sorted['Average Resolution Time'].apply(format_timedelta)}),
    "14. Average Resolved Time by Team Member", show_index=False, colalign=("left", "right"))

print_table(
    resolution_summary,
    "15 Case Count by Resolution Time Range",
    show_index=False,
    colalign=("left", "right", "right")
)

print_table(subject_escalated_counts, "16: Escalated cases by Subject", show_index=False)
if avg_escalated_time is not None:
    print("17. Average resolved time of Escalated cases:", avg_escalated_time)
else:
    print(". No resolved escalated cases found.")

