import pandas as pd
import streamlit as st
import sqlite3
from fuzzywuzzy import process, fuzz
from concurrent.futures import ThreadPoolExecutor

# Load CRM data from SQLite database
conn = sqlite3.connect("C:/Users/ScottPhillips/OneDrive - Affinity Group/Desktop/Python/Fuzzy Match Project/crm_data.db")
crm_df = pd.read_sql("SELECT * FROM crm", conn)

# Show preview of CRM data
st.write("CRM Data Preview (First 500 rows):")
st.write(crm_df.head(500))  # Displaying 500 rows from CRM for preview
st.write(f"Total CRM records loaded: {len(crm_df)}")

# Function to clean and standardize text for fuzzy matching
def clean_text(text):
    return ' '.join(str(text).upper().replace(".", "").replace(",", "").split())

# Function to get the best match from the CRM data
def get_best_match(row, crm_df):
    query_name = clean_text(row['companyName'])
    query_address = clean_text(row['companyAddress'])
    query_city = clean_text(row['companyCity'])
    query_state = clean_text(row['companyState'])

    # Combine relevant CRM fields into a single string for matching
    crm_df['combined'] = crm_df['companyName'].apply(clean_text) + ' ' + crm_df['companyAddress'].apply(clean_text) + ' ' + crm_df['companyCity'].apply(clean_text) + ' ' + crm_df['companyState'].apply(clean_text)

    # Perform fuzzy matching on company name and address
    choices = crm_df['combined'].tolist()
    best_match_tuple = process.extractOne(f"{query_name} {query_address} {query_city} {query_state}", choices, scorer=fuzz.token_sort_ratio) or ("", 0)
    best_match, score = best_match_tuple

    # Ensure state matches before confirming
    if score >= 65 and query_state == crm_df.iloc[choices.index(best_match)]['companyState']:
        match_row = crm_df.iloc[choices.index(best_match)]
        return match_row['systemId'], score, match_row['companyName'], match_row['companyAddress'], match_row['companyCity'], match_row['companyState'], match_row['companyZipCode']
    else:
        return "", 0, "", "", "", "", ""

# Streamlit UI to upload the user's file
st.title("Fuzzy Matching Tool")

uploaded_file = st.file_uploader("Upload your file for matching", type=["csv", "xlsx"])
if uploaded_file is not None:
    # Load the uploaded file into a dataframe
    user_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Show preview of user-uploaded data
    st.write("Uploaded File Preview:")
    st.write(user_df.head())  # Show the first 5 rows of the uploaded file
    st.write(f"Total records uploaded: {len(user_df)}")

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

