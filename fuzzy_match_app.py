import pandas as pd
import streamlit as st
import sqlite3
import re
from rapidfuzz import process, fuzz  # Use rapidfuzz for name matching
from concurrent.futures import ThreadPoolExecutor
import time

# Remove system sleep prevention for non-Windows environments
import platform
if platform.system() == "Windows":
    import ctypes

# Function to prevent sleep mode (Windows only)
def prevent_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

# Function to allow system sleep (Windows only)
def allow_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

# Prevent sleep at the start of the script
prevent_sleep()

# Function to standardize and clean address
def standardize_address(address):
    # Remove punctuation and special characters
    address = re.sub(r'[.,\-()/]', '', address)
    
    # Standardize common abbreviations
    address = re.sub(r'\bSTREET\b|\bST\b', 'ST', address, flags=re.IGNORECASE)
    address = re.sub(r'\bROAD\b|\bRD\b', 'RD', address, flags=re.IGNORECASE)
    address = re.sub(r'\bBOULEVARD\b|\bBLVD\b', 'BLVD', address, flags=re.IGNORECASE)
    address = re.sub(r'\bDRIVE\b|\bDR\b', 'DR', address, flags=re.IGNORECASE)
    address = re.sub(r'\bAVENUE\b|\bAVE\b', 'AVE', address, flags=re.IGNORECASE)
    address = re.sub(r'\bCOURT\b|\bCT\b', 'CT', address, flags=re.IGNORECASE)
    address = re.sub(r'\bLANE\b|\bLN\b', 'LN', address, flags=re.IGNORECASE)
    address = re.sub(r'\bCIRCLE\b|\bCIR\b', 'CIR', address, flags=re.IGNORECASE)
    address = re.sub(r'\bPARKWAY\b|\bPKWY\b', 'PKWY', address, flags=re.IGNORECASE)
    address = re.sub(r'\bSUITE\b|\bSTE\b', 'STE', address, flags=re.IGNORECASE)
    address = re.sub(r'\bHIGHWAY\b|\bHWY\b', 'HWY', address, flags=re.IGNORECASE)
    address = re.sub(r'\bPLACE\b|\bPL\b', 'PL', address, flags=re.IGNORECASE)

    # Directional standardizations
    address = re.sub(r'\bNORTH\b', 'N', address, flags=re.IGNORECASE)
    address = re.sub(r'\bSOUTH\b', 'S', address, flags=re.IGNORECASE)
    address = re.sub(r'\bEAST\b', 'E', address, flags=re.IGNORECASE)
    address = re.sub(r'\bWEST\b', 'W', address, flags=re.IGNORECASE)
    address = re.sub(r'\bNORTHEAST\b', 'NE', address, flags=re.IGNORECASE)
    address = re.sub(r'\bNORTHWEST\b', 'NW', address, flags=re.IGNORECASE)
    address = re.sub(r'\bSOUTHEAST\b', 'SE', address, flags=re.IGNORECASE)
    address = re.sub(r'\bSOUTHWEST\b', 'SW', address, flags=re.IGNORECASE)

    # Remove double spaces
    address = re.sub(r'\s+', ' ', address)
    
    return address.strip()

# Function to clean and prepare text (general cleaning for names)
def clean_text(text):
    return ' '.join(str(text).upper().split())

# Function to return match results
def return_match(best_match_row, score):
    return {
        'Match ID': best_match_row['systemId'],
        'Match Score': score,
        'Matched Name': best_match_row['companyName'],
        'Matched Address': best_match_row['companyAddress'],
        'Matched City': best_match_row['companyCity'],
        'Matched State': best_match_row['companyState'],
        'Matched Zip': best_match_row['companyZipCode']
    }

# Function to return no match
def no_match():
    return {
        'Match ID': '',
        'Match Score': 0,
        'Matched Name': '',
        'Matched Address': '',
        'Matched City': '',
        'Matched State': '',
        'Matched Zip': ''
    }

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
        # If there's an exact address match, return it with a 100% match score
        best_match_row = exact_matches.iloc[0]
        return return_match(best_match_row, 100)

    # Step 2: Fuzzy match on name but ensure same city or state matches
    best_match_tuple = process.extractOne(query_name, crm_df['combined_name'], scorer=fuzz.token_sort_ratio)

    if best_match_tuple and best_match_tuple[1] >= 65:
        best_match_name = best_match_tuple[0]
        best_match_row = crm_df[crm_df['combined_name'] == best_match_name].iloc[0]

        # Ensure the city or state is at least a partial match
        if query_city == clean_text(best_match_row['companyCity']) or query_state == clean_text(best_match_row['companyState']):
            return return_match(best_match_row, best_match_tuple[1])

    return no_match()

# Streamlit UI
st.title("Fuzzy Matching Tool")

# Connect to SQLite database and limit the columns pulled
conn = sqlite3.connect("crm_data.db")
crm_df = pd.read_sql("SELECT companyName, companyAddress, companyState, companyCity, companyZipCode, systemId FROM crm", conn)

# Standardize CRM data right after loading
crm_df['companyName'] = crm_df['companyName'].apply(clean_text)
crm_df['companyAddress'] = crm_df['companyAddress'].apply(standardize_address)
crm_df['companyCity'] = crm_df['companyCity'].apply(clean_text)
crm_df['companyState'] = crm_df['companyState'].apply(clean_text)
crm_df['combined_name'] = crm_df['companyName']  # For fuzzy matching on names

# Preview CRM data
st.write(f"CRM Data Loaded: {crm_df.shape[0]} rows")
st.dataframe(crm_df.head(500))  # Preview first 500 rows

# File uploader for the user’s file
uploaded_file = st.file_uploader("Upload your file for matching", type=["csv", "xlsx"])
if uploaded_file is not None:
    # Load the uploaded file (CSV or Excel)
    user_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Clean and standardize the uploaded file just like the CRM data
    user_df['companyName'] = user_df['companyName'].apply(clean_text)
    user_df['companyAddress'] = user_df['companyAddress'].apply(standardize_address)
    user_df['companyCity'] = user_df['companyCity'].apply(clean_text)
    user_df['companyState'] = user_df['companyState'].apply(clean_text)

    # Perform fuzzy matching in parallel with progress tracking
    results = []
    st.write("Fuzzy matching in progress, please wait...")
    progress_bar = st.progress(0)
    total_rows = len(user_df)

    with ThreadPoolExecutor(max_workers=4) as executor:
        for idx, result in enumerate(executor.map(lambda row: get_best_match(row, crm_df), [row for _, row in user_df.iterrows()])):
            results.append(result)
            progress_bar.progress((idx + 1) / total_rows)  # Update progress bar

    # Create results DataFrame
    for key in results[0].keys():
        user_df[key] = [result[key] for result in results]

    # Write the results to a new Excel file
    output_path = "matched_file.xlsx"
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        user_df.to_excel(writer, index=False, sheet_name='Original Data')

    # Streamlit download button
    st.success("Fuzzy matching completed! Download the results below:")
    st.download_button(label="Download Matched Results", data=open(output_path, "rb").read(), file_name=output_path)

# Allow the system to sleep after execution
allow_sleep()

conn.close()
