import pandas as pd
import re
import math
from datetime import timedelta
from tabulate import tabulate
import numpy as np


file_path = "L2 Platform Support Master Data_PST.xlsx"

df = pd.read_excel(file_path, sheet_name='Sheet1')

# Process datetime
df['Entered Queue'] = pd.to_datetime(df['Entered Queue'], errors='coerce')
df['Resolution Date'] = pd.to_datetime(df['Resolution Date'], errors='coerce')

# Convert PST to EST
# time_columns = ['Entered Queue', 'Resolution Date']
# for col in time_columns:
#     df[col] = df[col] + pd.Timedelta(hours=3)


def business_timedelta(start, end):
    """
    Calculate timedelta excluding weekends (Saturday, Sunday).
    """
    if pd.isna(start) or pd.isna(end):
        return pd.NaT

    start_date = start.date()
    end_date = end.date()

    # Same-day case
    if start_date == end_date:
        if start.weekday() < 5:  # Mon–Fri
            return end - start
        else:
            return timedelta(0)

    # Count weekdays INCLUDING the end date if it's Mon–Fri
    weekdays = np.busday_count(start_date, end_date)
    if end.weekday() < 5:
        weekdays += 1

    # Remove 2 since we’ll handle the first and last day separately
    full_days = max(weekdays - 2, 0)

    # Partial first day
    end_of_start_day = pd.Timestamp.combine(start_date, pd.Timestamp.max.time()).replace(
        hour=23, minute=59, second=59
    )
    partial_first = (end_of_start_day - start).total_seconds()
    if start.weekday() >= 5:  # weekend start
        partial_first = 0

    # Partial last day
    start_of_end_day = pd.Timestamp.combine(end_date, pd.Timestamp.min.time())
    partial_last = (end - start_of_end_day).total_seconds()
    if end.weekday() >= 5:  # weekend end
        partial_last = 0

    total_secs = int(full_days * 86400 + partial_first + partial_last)  # cast to int
    return timedelta(seconds=total_secs)


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

total_platform_cases = platform_counts_df["Case Count"].sum()

platform_counts_df["Percentage"] = (
    platform_counts_df["Case Count"] / total_platform_cases * 100
).round(1).astype(str) + "%"

# Add total row
platform_total_row = pd.DataFrame([{
    "Platform": "Total",
    "Case Count": total_platform_cases,
    "Percentage": "100.0%"
}])

platform_summary = pd.concat([platform_counts_df, platform_total_row], ignore_index=True)

# 2. Case count by platform per month
df['Year-Month'] = df['Entered Queue'].dt.to_period('M').astype(str)

monthly_platform_counts = (
    df.groupby(['Year-Month', 'Platform'])
    .size()
    .reset_index(name='Case Count')
)

# Add percentage relative to month total
monthly_totals = monthly_platform_counts.groupby('Year-Month')['Case Count'].transform('sum')
monthly_platform_counts['Percentage'] = (
    monthly_platform_counts['Case Count'] / monthly_totals * 100
).round(1).astype(str) + "%"


# 3. Top 5 most common subjects per platform
platform_totals = df.groupby('Platform').size()

# Subject counts
platform_subject_counts = (
    df.groupby(['Platform', 'Subject']).size().reset_index(name='Case Count')
)

# % of the platform total
platform_subject_counts['Percentage'] = (
    platform_subject_counts['Case Count'] /
    platform_subject_counts['Platform'].map(platform_totals) * 100
).round(1).astype(str) + '%'

# Top 5 per platform
top5_per_platform = (
    platform_subject_counts
    .sort_values(['Platform', 'Case Count'], ascending=[True, False])
    .groupby('Platform')
    .head(5)
)

# 4. Top 10 Customers per Platform
platform_totals = df.groupby('Platform').size()

# Customer counts, excluding MHS Inc, MHS Case Temp
platform_customer_counts = (
    df[~df['Customer'].isin(["Multi-Health Systems Inc.", "MHS Case Temp"])]
    .groupby(['Platform', 'Customer']).size()
    .reset_index(name='Case Count')
)

# Add percentage relative to platform total
platform_customer_counts['Percentage'] = (
    platform_customer_counts['Case Count'] /
    platform_customer_counts['Platform'].map(platform_totals) * 100
).round(1).astype(str) + '%'

# Sort and take top 10 per platform
top10_per_platform = (
    platform_customer_counts
    .sort_values(['Platform', 'Case Count'], ascending=[True, False])
    .groupby('Platform')
    .head(10)
)

# 5. Cases worked by team member
cases_by_member_df = df['Worked By'].value_counts().reset_index()
cases_by_member_df.columns = ['Team Member', 'Case Count']
cases_by_member_df.index += 1

total_member_cases = cases_by_member_df["Case Count"].sum()
cases_by_member_df["Percentage"] = (
    cases_by_member_df["Case Count"] / total_member_cases * 100
).round(1).astype(str) + "%"

member_total_row = pd.DataFrame([{
    "Team Member": "Total",
    "Case Count": total_member_cases,
    "Percentage": "100.0%"
}])

cases_by_member_summary = pd.concat([cases_by_member_df, member_total_row], ignore_index=True)

# 6. Platform and Case Count by Member
member_platform_counts = (
    df.groupby(['Worked By', 'Platform'])
      .size()
      .reset_index(name='Case Count')
      .sort_values(['Worked By', 'Case Count'], ascending=[True, False])
)

# 7. Case count by priority
df['Priority'] = df['Priority'].fillna('Normal')
cases_by_priority_df = df['Priority'].value_counts().reset_index()
cases_by_priority_df.columns = ['Priority', 'Case Count']
cases_by_priority_df.index += 1

total_priority_cases = cases_by_priority_df["Case Count"].sum()
cases_by_priority_df["Percentage"] = (
    cases_by_priority_df["Case Count"] / total_priority_cases * 100
).round(1).astype(str) + "%"

priority_total_row = pd.DataFrame([{
    "Priority": "Total",
    "Case Count": total_priority_cases,
    "Percentage": "100.0%"
}])

cases_by_priority_summary = pd.concat([cases_by_priority_df, priority_total_row], ignore_index=True)

# 8. Group by Priority + Subject
priority_totals = df.groupby('Priority').size()

priority_subject_counts = (
    df.groupby(['Priority', 'Subject']).size().reset_index(name='Case Count')
)

priority_subject_counts['Percentage'] = (
    priority_subject_counts['Case Count'] /
    priority_subject_counts['Priority'].map(priority_totals) * 100
).round(1).astype(str) + '%'

top5_subjects_per_priority = (
    priority_subject_counts
    .sort_values(['Priority', 'Case Count'], ascending=[True, False])
    .groupby('Priority')
    .head(5)
)

# 9. Top 10 busiest days
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

total_avg_cases = avg_cases_by_dow["Case Count"].sum()
avg_cases_by_dow["Percentage"] = (
    avg_cases_by_dow["Case Count"] / total_avg_cases * 100
).round(1).astype(str) + "%"

dow_total_row = pd.DataFrame([{
    "Day of Week": "Total",
    "Case Count": total_avg_cases,
    "Percentage": "100.0%"
}])

avg_cases_summary = pd.concat([avg_cases_by_dow, dow_total_row], ignore_index=True)

# 11. Case Count by Hour
peak_hours_df = df['Entered Queue'].dt.hour.value_counts().sort_index().reset_index()
peak_hours_df.columns = ['Hour', 'Cases Entered']

total_hourly_cases = peak_hours_df["Cases Entered"].sum()
peak_hours_df["Percentage"] = (
    peak_hours_df["Cases Entered"] / total_hourly_cases * 100
).round(1).astype(str) + "%"

hour_total_row = pd.DataFrame([{
    "Hour": "Total",
    "Cases Entered": total_hourly_cases,
    "Percentage": "100.0%"
}])

hourly_summary = pd.concat([peak_hours_df, hour_total_row], ignore_index=True)

# Filter cases that have a resolution date
resolved_cases = df[df['Resolution Date'].notna()].copy()

# Calculate resolution time
resolved_cases['Average Resolution Time'] = resolved_cases.apply(
    lambda row: business_timedelta(row['Entered Queue'], row['Resolution Date']), axis=1
)

# Recompute Hours & Days based on new business timedelta
resolved_cases['Resolution Hours'] = resolved_cases['Average Resolution Time'].apply(
    lambda td: td.total_seconds() / 3600 if pd.notna(td) else None
)
resolved_cases['Resolution Days'] = resolved_cases['Average Resolution Time'].apply(
    lambda td: td.total_seconds() / 86400 if pd.notna(td) else None
)


# Round resolution time columns to full seconds
resolved_cases['Average Resolution Time'] = resolved_cases['Average Resolution Time'].dt.round('1s')

# 12. Resolution Time
# 12.1. Average resolved time for all cases
avg_resolved_time = resolved_cases['Average Resolution Time'].mean()

# 12.2. Average resolved time of Normal priority cases
avg_normal_priority = resolved_cases[resolved_cases['Priority'].isna() | (resolved_cases['Priority'] == 'Normal')]['Average Resolution Time'].mean()

# 12.3. Average resolved time of High priority cases
avg_high_priority = resolved_cases[resolved_cases['Priority'] == 'High']['Average Resolution Time'].mean()


def format_timedelta(td):
    if pd.isna(td):
        return "N/A"
    total_seconds = int(td.total_seconds())
    return str(timedelta(seconds=total_seconds))


# 13. Average resolved time by platform
avg_by_platform = resolved_cases.groupby('Platform')['Average Resolution Time'].mean().reset_index()
avg_by_platform_sorted = avg_by_platform.sort_values(by='Average Resolution Time')

avg_by_platform_with_days = avg_by_platform_sorted.copy()

# Add days column (1 decimal)
avg_by_platform_with_days["Resolution Days"] = (
    avg_by_platform_with_days["Average Resolution Time"].dt.total_seconds() / 86400
).round(1)

# Format the timedelta column into readable hh:mm:ss
avg_by_platform_with_days["Average Resolution Time"] = (
    avg_by_platform_with_days["Average Resolution Time"].apply(format_timedelta)
)

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

# Total escalated cases
total_escalated_cases = escalated_cases.shape[0]

# Escalated cases with a non-empty subject
escalated_with_subject = escalated_cases[escalated_cases['Subject'].notna() & (escalated_cases['Subject'].astype(str).str.strip() != "")]
escalated_with_subject_count = escalated_with_subject.shape[0]

# 17. Count number of escalated cases per Subject
subject_escalated_counts = escalated_cases['Subject'].value_counts().reset_index()
subject_escalated_counts.columns = ['Subject', 'Escalated Case Count']

total_escalated = subject_escalated_counts["Escalated Case Count"].sum()
subject_escalated_counts["Percentage"] = (
    subject_escalated_counts["Escalated Case Count"] / total_escalated * 100
).round(1).astype(str) + "%"

subject_total_row = pd.DataFrame([{
    "Subject": "Total",
    "Escalated Case Count": total_escalated,
    "Percentage": "100.0%"
}])

subject_escalated_summary = pd.concat([subject_escalated_counts, subject_total_row], ignore_index=True)

# 17.1. Escalated Subjects by Platform
escalated_subject_platform_counts = (
    escalated_cases
    .groupby(['Platform', 'Subject'])
    .size()
    .reset_index(name='Escalated Case Count')
    .sort_values(['Platform', 'Escalated Case Count'], ascending=[True, False])
)

# 18. Average resolved time for escalated cases
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

# ------------------------------------- PRINTING ---------------------------------------------


# 1. Case count by platform
print_table(
    platform_summary,
    "1. CASE COUNT BY PLATFORM",
    show_index=False,
    colalign=("left", "right", "right")
)

print("\n2. CASE COUNT BY PLATFORM (MONTHLY)")
for month, table in monthly_platform_counts.groupby('Year-Month'):
    # Sort descending by Case Count (same as Section 1)
    table = table.sort_values("Case Count", ascending=False).reset_index(drop=True).copy()

    total_cases = table['Case Count'].sum()
    pct_sum = pd.to_numeric(table['Percentage'].str.rstrip('%')).sum()

    total_row = pd.DataFrame([{
        "Year-Month": month,
        "Platform": "Total",
        "Case Count": total_cases,
        "Percentage": f"{pct_sum:.1f}%"
    }])

    table_with_total = pd.concat([table, total_row], ignore_index=True)

    print_table(
        table_with_total.drop(columns=['Year-Month']),
        f"Case Count by Platform - {month}",
        show_index=False,
        colalign=("left", "right", "right")
    )

# 2. Top 5 Subjects per Platform
print("\n3. TOP 5 SUBJECTS BY PLATFORM")
for platform, table in top5_per_platform.groupby('Platform'):
    print_table(
        table.reset_index(drop=True),
        f"Top 5 Subjects - {platform}",
        show_index=False,
        colalign=("left", "left", "right", "right")
    )

# 3. Top 10 Customers per Platform
print("\n4. TOP 10 CUSTOMERS BY PLATFORM")
for platform, table in top10_per_platform.groupby('Platform'):
    print_table(
        table.reset_index(drop=True),
        f"Top 10 Customers - {platform}",
        show_index=False,
        colalign=("left", "left", "right", "right")
    )


# 4. Case Count by Team Member
print_table(
    cases_by_member_summary,
    "\n5. CASE COUNT BY TEAM MEMBER",
    show_index=False,
    colalign=("left", "right", "right")
)

print("\n6. PLATFORMS WORKED BY TEAM MEMBER")
for member, table in member_platform_counts.groupby('Worked By'):
    total_cases = table["Case Count"].sum()
    table["Percentage"] = (table["Case Count"] / total_cases * 100).round(1).astype(str) + "%"

    member_total_row = pd.DataFrame([{
        "Worked By": member,
        "Platform": "Total",
        "Case Count": total_cases,
        "Percentage": "100.0%"
    }])

    table_with_total = pd.concat([table, member_total_row], ignore_index=True)

    print_table(
        table_with_total.reset_index(drop=True).drop(columns=["Worked By"]),
        f"Platforms - {member}",
        show_index=False,
        colalign=("left", "right", "right")
    )

print_table(
    cases_by_priority_summary,
    "\n7. CASE COUNT BY PRIORITY",
    show_index=False,
    colalign=("left", "right", "right")
)

print("\n8. TOP 5 SUBJECTS BY PRIORITY")
for priority, table in top5_subjects_per_priority.groupby('Priority'):
    print_table(
        table.reset_index(drop=True),
        f"Top 5 Subjects - Priority: {priority}",
        show_index=False,
        colalign=("left", "left", "right", "right")
    )

print_table(top_days_df, "\n9. TOP 10 BUSIEST DAYS OF 2025")

print_table(
    avg_cases_summary,
    "\n10. AVERAGE CASE COUNT BY WEEKDAY",
    show_index=False,
    colalign=("left", "right", "right")
)

print_table(
    hourly_summary,
    "\n11. CASE ENTERED QUEUE BY HOUR (EST)",
    show_index=False,
    colalign=("left", "right", "right")
)

print("\n12. AVERAGE RESOLUTION TIME BY PRIORITY")
print("Overall average (all cases):", format_timedelta(avg_resolved_time))
print("Normal priority cases:", format_timedelta(avg_normal_priority))
print("High priority cases:", format_timedelta(avg_high_priority))

print_table(
    avg_by_platform_with_days,
    "\n13. AVERAGE RESOLUTION TIME BY PLATFORM",
    show_index=False,
    colalign=("left", "right", "right")
)

print_table(avg_by_member_sorted.assign(
    **{'Average Resolution Time': avg_by_member_sorted['Average Resolution Time'].apply(format_timedelta)}),
    "\n14. AVERAGE RESOLUTION TIME BY TEAM MEMBER", show_index=False, colalign=("left", "right"))

print_table(
    resolution_summary,
    "\n15. CASE COUNT BY RESOLUTION TIME RANGE",
    show_index=False,
    colalign=("left", "right", "right")
)

print(f"\nESCALATED CASES:")
print(f"Total escalated cases: {total_escalated_cases}")
print(f"Escalated cases with a Subject: {escalated_with_subject_count}")

if avg_escalated_time is not None:
    print("Average resolved time of Escalated cases:", avg_escalated_time)

print_table(subject_escalated_summary, "\n16. ESCALATED CASE COUNT BY SUBJECT", show_index=False)

print("\n17. ESCALATED SUBJECTS BY PLATFORM")
for platform, table in escalated_subject_platform_counts.groupby('Platform'):
    total_platform = table["Escalated Case Count"].sum()
    table["Percentage"] = (table["Escalated Case Count"] / total_platform * 100).round(1).astype(str) + "%"

    platform_total_row = pd.DataFrame([{
        "Platform": platform,
        "Subject": "Total",
        "Escalated Case Count": total_platform,
        "Percentage": "100.0%"
    }])

    table_with_total = pd.concat([table, platform_total_row], ignore_index=True)

    print_table(
        table_with_total.reset_index(drop=True),
        f"Escalated Subjects - {platform}",
        show_index=False
    )

# 17. AVERAGE RESOLUTION TIME FOR ESCALATED CASES BY PLATFORM
if not escalated_resolved_cases.empty:
    escalated_avg_by_platform = (
        escalated_resolved_cases
        .groupby('Platform')['Average Resolution Time']
        .mean()
        .reset_index()
    )

    # Format as d hh:mm:ss
    escalated_avg_by_platform["Avg Resolution"] = (
        escalated_avg_by_platform["Average Resolution Time"].apply(format_timedelta)
    )

    # Days (rounded 1 decimal)
    escalated_avg_by_platform["Avg Resolution Days"] = (
        escalated_avg_by_platform["Average Resolution Time"].dt.total_seconds() / 86400
    ).round(1)

    # Sort from shortest to longest
    escalated_avg_by_platform = escalated_avg_by_platform.sort_values("Average Resolution Time")

    # Drop raw timedelta (keep formatted)
    escalated_avg_by_platform = escalated_avg_by_platform.drop(columns=["Average Resolution Time"])

    # Reorder columns: hh:mm:ss before days
    escalated_avg_by_platform = escalated_avg_by_platform[
        ["Platform", "Avg Resolution", "Avg Resolution Days"]
    ].reset_index(drop=True)

    # Make rank start from 1 instead of 0
    escalated_avg_by_platform.index += 1

    print_table(
        escalated_avg_by_platform,
        "\n18. AVERAGE RESOLUTION TIME FOR ESCALATED CASES BY PLATFORM",
        show_index=True,
        colalign=("left", "left", "right")
    )


# 19. CASE COUNT BY PLATFORM GROUP (Apr 1 - Sep 30, 2025)
start_date = pd.Timestamp('2025-04-01')
end_date = pd.Timestamp('2025-09-30')

# Filter by date range
df_apr_sep = df[(df['Entered Queue'] >= start_date) & (df['Entered Queue'] <= end_date)].copy()

# Define overlapping group membership
group_definitions = {
    'Portals - MAC+ / LMS / GIFR / USB / FAS': ['MAC+', 'LMS', 'GIFR', 'USB', 'FAS'],
    'Education - TAP': ['TAP'],
    'Public Safety - GEARS / CORE PATHWAY': ['GEARS', 'CORE SOLUTIONS'],
    'Gifted - MGI': ['MGI']
}

total_cases = df_apr_sep.shape[0]


def summarize_cases_overlap(escalated_value, title):
    # Filter escalated / non-escalated
    filtered = df_apr_sep[
        df_apr_sep['Escalated'].astype(str).str.strip().str.lower().eq('yes')
        if escalated_value else
        df_apr_sep['Escalated'].astype(str).str.strip().str.lower().ne('yes')
    ]

    results = []
    for group_name, platforms in group_definitions.items():
        count = filtered[filtered['Platform'].astype(str).str.upper().isin(platforms)].shape[0]
        results.append({"Platform Group": group_name, "Case Count": count})

    # Compute percentage of total (5241)
    summary = pd.DataFrame(results)
    summary["% of Total"] = (summary["Case Count"] / total_cases * 100).round(1).astype(str) + "%"

    # Add total row
    total_row = pd.DataFrame([{
        "Platform Group": "Total",
        "Case Count": summary["Case Count"].sum(),
        "% of Total": f"{(summary['Case Count'].sum() / total_cases * 100):.1f}%"
    }])
    summary = pd.concat([summary, total_row], ignore_index=True)

    print_table(summary, f"\n19. {title}", show_index=False, colalign=("left", "right", "right"))


# Run both versions
print(f"\nTotal cases from April 1 to September 30, 2025: {total_cases}")
summarize_cases_overlap(escalated_value=False, title="NON-ESCALATED CASES (Apr 1 - Sep 30, 2025)")
summarize_cases_overlap(escalated_value=True, title="ESCALATED CASES (Apr 1 - Sep 30, 2025)")

# ------------------------------------- EXPORT RESULTS TO EXCEL ---------------------------------------------
output_path = "analysis_output.xlsx"


def safe_sheet_name(name: str) -> str:
    """Truncate/sanitize sheet names to be Excel-safe."""
    name = re.sub(r'[\\/*?:\[\]]', '', name)
    return name[:31]


def concat_with_blank_rows(grouped_df):
    """Combine grouped DataFrames with a blank row between groups, preserving column order."""
    parts = []
    # Keep column order from the first group
    first_cols = None
    for _, subdf in grouped_df:
        if first_cols is None:
            first_cols = list(subdf.columns)
        subdf = subdf.reindex(columns=first_cols)
        parts.append(subdf)
        blank = pd.DataFrame([{col: "" for col in first_cols}])
        parts.append(blank)
    return pd.concat(parts, ignore_index=True)


with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

    # 1. Case count by platform
    platform_summary.to_excel(writer, sheet_name="1_Case_Count_by_PF", index=False)

    # 2. Case count by platform (monthly) — separated by blank rows
    monthly_sorted_parts = []

    for month, subdf in monthly_platform_counts.groupby("Year-Month"):
        # Sort descending by Case Count
        subdf_sorted = subdf.sort_values("Case Count", ascending=False).reset_index(drop=True)

        # Add total row
        total_cases = subdf_sorted["Case Count"].sum()
        pct_sum = pd.to_numeric(subdf_sorted["Percentage"].str.rstrip('%')).sum()
        total_row = pd.DataFrame([{
            "Year-Month": month,
            "Platform": "Total",
            "Case Count": total_cases,
            "Percentage": f"{pct_sum:.1f}%"
        }])

        # Combine this month’s block and a blank row
        month_block = pd.concat([subdf_sorted, total_row], ignore_index=True)
        monthly_sorted_parts.append(month_block)
        monthly_sorted_parts.append(pd.DataFrame([{col: "" for col in month_block.columns}]))

    # Combine all months, preserving column order
    monthly_concat = pd.concat(monthly_sorted_parts, ignore_index=True)[
        ["Year-Month", "Platform", "Case Count", "Percentage"]
    ]

    monthly_concat.to_excel(writer, sheet_name=safe_sheet_name("2_Monthly_Platform_Cases"), index=False)

    # 3. Top 5 subjects per platform
    top5_concat = concat_with_blank_rows(top5_per_platform.groupby("Platform"))
    top5_concat.to_excel(writer, sheet_name=safe_sheet_name("3_Top5_Subjects_by_PF"), index=False)

    # 4. Top 10 customers per platform
    top10_concat = concat_with_blank_rows(top10_per_platform.groupby("Platform"))
    top10_concat.to_excel(writer, sheet_name=safe_sheet_name("4_Top10_Customers_by_PF"), index=False)

    # 5. Case count by team member
    cases_by_member_summary.to_excel(writer, sheet_name="5_Case_by_Member", index=False)

    # 6. Platform & case count by member — separated by blank rows
    member_concat = concat_with_blank_rows(member_platform_counts.groupby("Worked By"))
    member_concat.to_excel(writer, sheet_name=safe_sheet_name("6_Platform_by_Member"), index=False)

    # 7. Case count by priority
    cases_by_priority_summary.to_excel(writer, sheet_name="7_Cases_by_Priority", index=False)

    # 8. Top 5 subjects by priority — separated by blank rows
    top5_priority_concat = concat_with_blank_rows(top5_subjects_per_priority.groupby("Priority"))
    top5_priority_concat.to_excel(writer, sheet_name=safe_sheet_name("8_Top5_Subjects_by_Priority"), index=False)

    # 9. Top 10 busiest days
    top_days_df.to_excel(writer, sheet_name="9_Top10_Busiest_Days", index=False)

    # 10. Average case count by weekday
    avg_cases_summary.to_excel(writer, sheet_name="10_Avg_Cases_by_Weekday", index=False)

    # 11. Case entered queue by hour
    hourly_summary.to_excel(writer, sheet_name="11_Cases_by_Hour", index=False)

    # 12. Average resolution time summary (text only)
    avg_res_summary = pd.DataFrame({
        "Priority": [
            "Overall average (all cases)",
            "Normal priority cases",
            "High priority cases"
        ],
        "Average Resolution Time": [
            format_timedelta(avg_resolved_time),
            format_timedelta(avg_normal_priority),
            format_timedelta(avg_high_priority)
        ]
    })
    avg_res_summary.to_excel(writer, sheet_name="12_Avg_Resolution_Time", index=False)

    # 13. Average resolution time by platform
    avg_by_platform_with_days.to_excel(writer, sheet_name="13_Avg_Res_Time_by_PF", index=False)

    # 14. Average resolution time by team member (formatted)
    avg_by_member_export = avg_by_member_sorted.copy()
    avg_by_member_export["Average Resolution"] = avg_by_member_export["Average Resolution Time"].apply(format_timedelta)
    avg_by_member_export["Resolution Days"] = (
        avg_by_member_export["Average Resolution Time"].dt.total_seconds() / 86400
    ).round(1)
    avg_by_member_export = avg_by_member_export.drop(columns=["Average Resolution Time"])
    avg_by_member_export.to_excel(writer, sheet_name="14_Avg_Res_Time_by_Member", index=False)

    # 15. Resolution time ranges
    resolution_summary.to_excel(writer, sheet_name="15_Res_Time_Range", index=False)

    # 16. Escalated subjects summary
    subject_escalated_summary.to_excel(writer, sheet_name="16_Escalated_Subjects", index=False)

    # 17. Escalated subjects by platform — separated by blank rows
    esc_concat = concat_with_blank_rows(escalated_subject_platform_counts.groupby("Platform"))
    esc_concat.to_excel(writer, sheet_name=safe_sheet_name("17_Escalated_Subjects_by_PF"), index=False)

    # 18. Escalated average resolution time (if exists)
    if 'escalated_avg_by_platform' in locals():
        escalated_avg_by_platform.to_excel(writer, sheet_name="18_Escalated_Avg_Res_Time", index=False)

    # 19. CASE COUNT BY PLATFORM GROUP (Apr 1 - Sep 30, 2025)
    start_date = pd.Timestamp('2025-04-01')
    end_date = pd.Timestamp('2025-09-30')
    df_apr_sep = df[(df['Entered Queue'] >= start_date) & (df['Entered Queue'] <= end_date)].copy()

    group_definitions = {
        'Portals - MAC+ / TAP / USB / FAS': ['MAC+', 'TAP', 'USB', 'FAS'],
        'Public Safety - LMS / GIFR / GEARS / CORE PATHWAY': ['LMS', 'GIFR', 'GEARS', 'CORE SOLUTIONS'],
        'Gifted - MGI': ['MGI']
    }

    total_cases = df_apr_sep.shape[0]


    def build_cases_overlap(escalated_value: bool, title: str):
        """Return a DataFrame matching the printed Section 19 table."""
        filtered = df_apr_sep[
            df_apr_sep['Escalated'].astype(str).str.strip().str.lower().eq('yes')
            if escalated_value else
            df_apr_sep['Escalated'].astype(str).str.strip().str.lower().ne('yes')
        ]

        results = []
        for group_name, platforms in group_definitions.items():
            count = filtered[filtered['Platform'].astype(str).str.upper().isin(platforms)].shape[0]
            results.append({"Platform Group": group_name, "Case Count": count})

        summary = pd.DataFrame(results)
        summary["% of Total"] = (summary["Case Count"] / total_cases * 100).round(1).astype(str) + "%"

        total_row = pd.DataFrame([{
            "Platform Group": "Total",
            "Case Count": summary["Case Count"].sum(),
            "% of Total": f"{(summary['Case Count'].sum() / total_cases * 100):.1f}%"
        }])
        summary = pd.concat([summary, total_row], ignore_index=True)
        return summary


    # Build both Section 19 tables
    non_escalated_19 = build_cases_overlap(False, "NON-ESCALATED CASES (Apr 1 - Sep 30, 2025)")
    escalated_19 = build_cases_overlap(True, "ESCALATED CASES (Apr 1 - Sep 30, 2025)")

    # Export them
    non_escalated_19.to_excel(writer, sheet_name=safe_sheet_name("19.1_L2_Cases_Apr-Sep2025"), index=False)
    escalated_19.to_excel(writer, sheet_name=safe_sheet_name("19.2_L3_Cases_Apr-Sep2025"), index=False)

print(f"\n✅ All tables exported successfully to {output_path}")
