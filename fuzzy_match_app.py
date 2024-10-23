import pandas as pd
import streamlit as st
import sqlite3
import re
from rapidfuzz import process, fuzz
from concurrent.futures import ThreadPoolExecutor
import os

# Prevent system sleep on Windows
import platform
if platform.system() == "Windows":
    import ctypes

def prevent_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

def allow_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

prevent_sleep()  # Prevent system sleep at the start of the script

# Caching the CRM data to reduce reload time
@st.cache_data
def load_crm_data(db_path):
    conn = sqlite3.connect(db_path)
    crm_df = pd.read_sql("SELECT companyName, companyAddress, companyState, companyCity, companyZipCode, systemId FROM crm", conn)
    conn.close()
    return crm_df

@st.cache_data
def standardize_and_clean_crm_data(crm_df):
    crm_df['companyName'] = crm_df['companyName'].apply(clean_text)
    crm_df['companyAddress'] = crm_df['companyAddress'].apply(standardize_address)
    crm_df['companyCity'] = crm_df['companyCity'].apply(clean_text)
    crm_df['companyState'] = crm_df['companyState'].apply(clean_text)
    return crm_df

# Function to standardize and clean address
def standardize_address(address):
    if not isinstance(address, str):
        return ""  # If the address is None or not a string, return an empty string
    
    # Remove punctuation and special characters
    address = re.sub(r'[.,\-()/]', '', address)
    
    # (Rest of standardization logic...)

    return address.strip()

# Function to clean and prepare text (general cleaning for names)
def clean_text(text):
    return ' '.join(str(text).upper().split())

# Function to get the best match with address (exact match) and name (fuzzy match)
def get_best_match(row, crm_df):
    query_name = clean_text(row['companyName'])
    query_address = standardize_address(row['companyAddress'])
    query_city = clean_text(row['companyCity'])
    query_state = clean_text(row['companyState'])

    # Step 1: Exact match on address + city + state
    exact_matches = crm_df[
        (crm_df['companyAddress'].apply(standardize_address) == query_address) &
        (crm_df['companyCity'].apply(clean_text) == query_city) &
        (crm_df['companyState'].apply(clean_text) == query_state)
    ]

    if not exact_matches.empty:
        best_match_row = exact_matches.iloc[0]
        return {
            'Match ID': best_match_row['systemId'],
            'Match Score': 100,
            'Matched Name': best_match_row['companyName'],
            'Matched Address': best_match_row['companyAddress'],
            'Matched City': best_match_row['companyCity'],
            'Matched State': best_match_row['companyState'],
            'Matched Zip': best_match_row['companyZipCode']
        }

    # Step 2: Fuzzy match on name but ensure same city or state matches
    best_match_tuple = process.extractOne(query_name, crm_df['companyName'].tolist(), scorer=fuzz.token_sort_ratio)

    if best_match_tuple and best_match_tuple[1] >= 65:
        best_match_name = best_match_tuple[0]
        best_match_row = crm_df[crm_df['companyName'] == best_match_name].iloc[0]

        if query_city == clean_text(best_match_row['companyCity']) or query_state == clean_text(best_match_row['companyState']):
            return {
                'Match ID': best_match_row['systemId'],
                'Match Score': best_match_tuple[1],
                'Matched Name': best_match_row['companyName'],
                'Matched Address': best_match_row['companyAddress'],
                'Matched City': best_match_row['companyCity'],
                'Matched State': best_match_row['companyState'],
                'Matched Zip': best_match_row['companyZipCode']
            }

    return {'Match ID': '', 'Match Score': 0, 'Matched Name': '', 'Matched Address': '', 'Matched City': '', 'Matched State': '', 'Matched Zip': ''}

# Streamlit UI
st.title("Fuzzy Matching Tool")

# Get the current script directory and set paths
base_path = os.path.dirname(os.path.realpath(__file__))
db_path = os.path.join(base_path, "crm_data.db")
output_path = os.path.join(base_path, "matched_file.xlsx")

# Load and cache CRM data
crm_df = load_crm_data(db_path)
crm_df = standardize_and_clean_crm_data(crm_df)

# Preview CRM data
st.write(f"CRM Data Loaded: {crm_df.shape[0]} rows")
st.dataframe(crm_df.head(500))

# Initialize session state for results and progress
if 'results' not in st.session_state:
    st.session_state['results'] = []
if 'processed_rows' not in st.session_state:
    st.session_state['processed_rows'] = 0

# File uploader for user's file
uploaded_file = st.file_uploader("Upload your file for matching", type=["csv", "xlsx"])
if uploaded_file is not None:
    # Load the uploaded file (CSV or Excel)
    user_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Standardize and clean user data
    user_df['companyName'] = user_df['companyName'].apply(clean_text)
    user_df['companyAddress'] = user_df['companyAddress'].apply(standardize_address)
    user_df['companyCity'] = user_df['companyCity'].apply(clean_text)
    user_df['companyState'] = user_df['companyState'].apply(clean_text)

    # Only process the rows that haven't been processed yet
    unprocessed_rows = user_df.iloc[st.session_state['processed_rows']:]

    # Progress bar for matching
    st.write("Fuzzy matching in progress, please wait...")
    progress_bar = st.progress(0)

    # Perform fuzzy matching
    with ThreadPoolExecutor(max_workers=4) as executor:
        for result in executor.map(lambda row: get_best_match(row, crm_df), [row for _, row in unprocessed_rows.iterrows()]):
            st.session_state['results'].append(result)

            # Update the number of processed rows in session state
            st.session_state['processed_rows'] += 1
            progress_bar.progress(st.session_state['processed_rows'] / len(user_df))

    # Add results to user_df
    for key in st.session_state['results'][0].keys():
        user_df[key] = [result[key] for result in st.session_state['results']]

    # Save results to Excel file
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        user_df.to_excel(writer, index=False, sheet_name='Original Data')

    # Streamlit download button
    st.success("Fuzzy matching completed! Download the results below:")
    st.download_button(label="Download Matched Results", data=open(output_path, "rb").read(), file_name="matched_file.xlsx")

# Allow system sleep after execution
allow_sleep()


