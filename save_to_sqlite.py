import pandas as pd
import sqlite3

# Path to your actual CSV file
csv_file = "C:/Users/ScottPhillips/OneDrive - Affinity Group/Desktop/Python/Fuzzy Match Project/crm_data.csv"

# Load the CRM data from CSV
df = pd.read_csv(csv_file)

# Create or connect to SQLite database
conn = sqlite3.connect("C:/Users/ScottPhillips/OneDrive - Affinity Group/Desktop/Python/Fuzzy Match Project/crm_data.db")

# Save the CRM data to a table named "crm"
df.to_sql("crm", conn, if_exists="replace", index=False)

# Close the connection
conn.close()

print("CRM data successfully saved to SQLite!")
