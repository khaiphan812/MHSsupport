import pandas as pd
import matplotlib.pyplot as plt

# --- Load Data ---
file_path = "MHS Platform Support 2025 7-22-2025 12-57-27 PM.xlsx"
sheet_name = "L2 Platform Support 2025"
df = pd.read_excel(file_path, sheet_name=sheet_name)

# --- Convert Dates ---
df['Created On'] = pd.to_datetime(df['Created On (Object) (Case)'], errors='coerce')
df['Modified On'] = pd.to_datetime(df['(Do Not Modify) Modified On'], errors='coerce')

# --- Create Resolution Time ---
df['Resolution Time (hrs)'] = (df['Modified On'] - df['Created On']).dt.total_seconds() / 3600

# --- Summary Metrics ---
total_tickets = len(df)
inactive_tickets = df['Status'].value_counts().get('Inactive', 0)
missing_descriptions = df['Description (Object) (Case)'].isna().sum()
unassigned_tickets = df['Worked By'].isna().sum()
average_resolution = round(df['Resolution Time (hrs)'].mean(), 2)
slow_tickets = len(df[df['Resolution Time (hrs)'] > 48])

# --- Team Productivity ---
top_staff = df['Worked By'].value_counts().head(6)

# --- Platform Analysis ---
platform_tickets = df['Platform (Object) (Case)'].value_counts().head(5)

# --- Recent Ticket Volume ---
ticket_volume = df['Created On'].dt.date.value_counts().sort_index().tail()

# --- Print Results ---
print("\n=== 📊 SUMMARY METRICS ===")
print(f"Total Tickets: {total_tickets}")
print(f"Inactive Tickets: {inactive_tickets}")
print(f"Tickets Missing Description: {missing_descriptions}")
print(f"Unassigned Tickets: {unassigned_tickets}")
print(f"Average Resolution Time: {average_resolution} hours")
print(f"Tickets Resolved > 48 hrs: {slow_tickets}")

print("\n=== 👥 TICKETS HANDLED BY TEAM MEMBERS ===")
print(top_staff.to_string())

print("\n=== 🖥️ TOP 5 PLATFORMS BY TICKETS ===")
print(platform_tickets.to_string())

print("\n=== 📅 RECENT TICKET VOLUME ===")
print(ticket_volume.to_string())

df = pd.read_excel("MHS Platform Support 2025 7-22-2025 12-57-27 PM.xlsx", sheet_name="L2 Platform Support 2025")

# Count tickets by platform
platform_counts = df['Platform (Object) (Case)'].value_counts()

# Draw a bar chart
# platform_counts.plot(kind='bar', color='skyblue')
# plt.title("Tickets per Platform")
# plt.xlabel("Platform")
# plt.ylabel("Number of Tickets")
# plt.tight_layout()
# plt.show()
#
# df['Created On'] = pd.to_datetime(df['Created On (Object) (Case)'], errors='coerce')
# daily_counts = df['Created On'].dt.date.value_counts().sort_index()
#
# daily_counts.plot(kind='line', marker='o')
# plt.title("Tickets Created per Day")
# plt.xlabel("Date")
# plt.ylabel("Tickets")
# plt.xticks(rotation=45)
# plt.tight_layout()
# plt.show()


import pandas as pd

# Load the Excel file
file_path = "MHS Platform Support 2025 7-22-2025 12-57-27 PM.xlsx"
sheet_name = "L2 Platform Support 2025"
df = pd.read_excel(file_path, sheet_name=sheet_name)

# --- Clean Request Type (for accurate counts) ---
df['Request Type Clean'] = df['Title'].astype(str).str.strip().str.lower()

# Top 5 most common request types (case-insensitive)
top_requests = df['Request Type Clean'].value_counts().head(5)
print("\nTop Request Types:")
for req_type, count in top_requests.items():
    print(f"- {req_type.title()} – {count} cases")

# --- Missing Information ---
missing_platform = df['Platform (Object) (Case)'].isna().sum()
missing_worked_by = df['Worked By'].isna().sum()

print("\nMissing Information:")
print(f"- Platform field missing in {missing_platform} entries")
print(f"- Worked By field missing in {missing_worked_by} entries")

# --- Ticket Priority Breakdown ---
print("\nTicket Priority Breakdown:")
priority_counts = df['Priority (Object) (Case)'].value_counts()
for priority, count in priority_counts.items():
    print(f"- {priority}: {count} tickets")

# --- Resolution Time Buckets ---
# Clean and parse 'Resolution Time (Hours)' column
df['Resolution Time (Hours)'] = pd.to_numeric(df['Created On (Object) (Case)'], errors='coerce')

# Define resolution buckets
bins = [0, 24, 48, 72, 168, float('inf')]
labels = ['<24h', '24-48h', '48-72h', '3-7 days', '>7 days']
df['Resolution Bucket'] = pd.cut(df['Resolution Time (Hours)'], bins=bins, labels=labels, right=False)

print("\nResolution Time Buckets:")
bucket_counts = df['Resolution Bucket'].value_counts().reindex(labels)
for label, count in bucket_counts.items():
    print(f"- {label}: {count} tickets")

# --- Performance Metric: % resolved in under 24h ---
total_tickets = len(df)
resolved_under_24h = bucket_counts['<24h']
percent_under_24h = (resolved_under_24h / total_tickets) * 100

print("\nPerformance Metrics:")
print(f"- Percent Resolved in under 24h: {percent_under_24h:.2f}%")
