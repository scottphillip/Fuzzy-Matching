import os
import pandas as pd
import streamlit as st
import sqlite3
from fuzzywuzzy import process, fuzz
from concurrent.futures import ThreadPoolExecutor

# Load CRM data from SQLite database
db_path = os.path.join(os.path.dirname(__file__), 'crm_data.db')
conn = sqlite3.connect(db_path)
crm_df = pd.read_sql("SELECT * FROM crm", conn)

# Show a preview of the CRM data to ensure it's loaded properly
st.write(f"CRM Data Loaded: {len(crm_df)} rows")
st.dataframe(crm_df.head(500))  # Show first 500 rows for preview

# Function to clean and prepare text for fuzzy matching
def clean_text(text):
    return ' '.join(str(text).upper().split())

# Function to get the best match from the CRM data
def get_best_match(row, crm_df):
    query = f"{row['companyName']} {row['companyAddress']} {row['companyCity']} {row['companyState']}"
    choices_df = crm_df.copy()
    choices_df['combined'] = crm_df['companyName'] + ' ' + crm_df['companyAddress'] + ' ' + crm_df['companyCity'] + ' ' + crm_df['companyState']
    choices = choices_df['combined'].tolist()

    best_match_tuple = process.extractOne(query, choices, scorer=fuzz.token_sort_ratio) or ("", 0)
    best_match, score = best_match_tuple

    if score >= 65 and str(row['companyState']) == str(choices_df.iloc[choices_df['combined'] == best_match].iloc[0]['companyState']):
        match_row = choices_df.iloc[choices_df['combined'] == best_match].iloc[0]
        match_id = match_row['systemId']
        return match_id, score, match_row['companyName'], match_row['companyAddress'], match_row['companyCity'], match_row['companyState'], match_row['companyZipCode']
    else:
        return "", 0, "", "", "", "", ""

# Streamlit UI to upload the user's file
st.title("Fuzzy Matching Tool")

uploaded_file = st.file_uploader("Upload your file for matching", type=["csv", "xlsx"])
if uploaded_file is not None:
    # Load the uploaded file into a dataframe
    user_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Show a preview of the user-uploaded file
    st.write(f"Uploaded File: {len(user_df)} rows")
    st.dataframe(user_df.head(500))  # Show first 500 rows of the uploaded file

    # Clean the text data
    user_df['companyName'] = user_df['companyName'].apply(clean_text)
    user_df['companyAddress'] = user_df['companyAddress'].apply(clean_text)
    user_df['companyCity'] = user_df['companyCity'].apply(clean_text)
    user_df['companyState'] = user_df['companyState'].apply(clean_text)

    # Perform fuzzy matching
    with ThreadPoolExecutor() as executor:
        results = list(executor.map(lambda row: get_best_match(row, crm_df), [row for _, row in user_df.iterrows()]))

    # Add the matching results to the dataframe
    user_df['Match ID'] = [result[0] for result in results]
    user_df['Match Score'] = [result[1] for result in results]
    user_df['Matched Name'] = [result[2] for result in results]
    user_df['Matched Address'] = [result[3] for result in results]
    user_df['Matched City'] = [result[4] for result in results]
    user_df['Matched State'] = [result[5] for result in results]
    user_df['Matched Zip'] = [result[6] for result in results]

    # Create a results dataframe with matched data
    matched_df = user_df[user_df['Match Score'] >= 65]

    # Write the results to a new sheet in the Excel file
    output_path = "matched_file.xlsx"
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        user_df.to_excel(writer, index=False, sheet_name='Original Data')
        matched_df.to_excel(writer, index=False, sheet_name='Matched Results')

    st.success("Fuzzy matching completed! Download the results below:")
    st.download_button(label="Download Matched Results", data=open(output_path, "rb").read(), file_name=output_path)

conn.close()
