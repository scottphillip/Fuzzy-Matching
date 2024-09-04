import pandas as pd
import streamlit as st
import sqlite3
from fuzzywuzzy import process, fuzz
from concurrent.futures import ThreadPoolExecutor

# Connect to SQLite database
conn = sqlite3.connect("crm_data.db")
crm_df = pd.read_sql("SELECT * FROM crm", conn)

# Function to clean and prepare text
def clean_text(text):
    return ' '.join(str(text).upper().split())

# Function to get the best match
def get_best_match(row, crm_df):
    query_name = clean_text(row['companyName'])
    query_address = clean_text(row['companyAddress'])
    query_state = clean_text(row['companyState'])

    crm_df['combined'] = crm_df['companyName'] + ' ' + crm_df['companyAddress'] + ' ' + crm_df['companyState']
    choices = crm_df['combined'].tolist()

    best_match_tuple = process.extractOne(f"{query_name} {query_address} {query_state}", choices, scorer=fuzz.token_sort_ratio)
    if best_match_tuple:
        best_match, score = best_match_tuple
        best_match_row = crm_df.loc[crm_df['combined'] == best_match].iloc[0]
        
        if score >= 65 and query_state == clean_text(best_match_row['companyState']):
            return {
                'Match ID': best_match_row['systemId'],
                'Match Score': score,
                'Matched Name': best_match_row['companyName'],
                'Matched Address': best_match_row['companyAddress'],
                'Matched City': best_match_row['companyCity'],
                'Matched State': best_match_row['companyState'],
                'Matched Zip': best_match_row['companyZipCode']
            }
    return {
        'Match ID': '',
        'Match Score': 0,
        'Matched Name': '',
        'Matched Address': '',
        'Matched City': '',
        'Matched State': '',
        'Matched Zip': ''
    }

# Streamlit UI
st.title("Fuzzy Matching Tool")

# Preview CRM data
st.write(f"CRM Data Loaded: {crm_df.shape[0]} rows")
st.dataframe(crm_df.head(500))  # Preview first 500 rows

# File uploader
uploaded_file = st.file_uploader("Upload your file for matching", type=["csv", "xlsx"])
if uploaded_file is not None:
    # Load the uploaded file
    user_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Clean the text data
    user_df['companyName'] = user_df['companyName'].apply(clean_text)
    user_df['companyAddress'] = user_df['companyAddress'].apply(clean_text)
    user_df['companyCity'] = user_df['companyCity'].apply(clean_text)
    user_df['companyState'] = user_df['companyState'].apply(clean_text)

    # Perform fuzzy matching
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(lambda row: get_best_match(row, crm_df), [row for _, row in user_df.iterrows()]))

    # Create results DataFrame
    for key in results[0].keys():
        user_df[key] = [result[key] for result in results]

    # Write the results to a new sheet in Excel
    output_path = "matched_file.xlsx"
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        user_df.to_excel(writer, index=False, sheet_name='Original Data')

    st.success("Fuzzy matching completed! Download the results below:")
    st.download_button(label="Download Matched Results", data=open(output_path, "rb").read(), file_name=output_path)

conn.close()
