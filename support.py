import pandas as pd

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
top_staff = df['Worked By'].value_counts().head(5)

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

print("\n=== 👥 TOP 5 STAFF BY TICKETS HANDLED ===")
print(top_staff.to_string())

print("\n=== 🖥️ TOP 5 PLATFORMS BY TICKETS ===")
print(platform_tickets.to_string())

print("\n=== 📅 RECENT TICKET VOLUME ===")
print(ticket_volume.to_string())
