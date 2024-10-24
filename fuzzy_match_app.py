import pandas as pd
import streamlit as st
import sqlite3
import re
from rapidfuzz import process, fuzz  # Use rapidfuzz for faster processing
from concurrent.futures import ThreadPoolExecutor
import time
import os
import hashlib
import pickle

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

# Caching CRM data in session state to avoid reloading
@st.cache_data
def load_crm_data(db_path):
    # Connect to SQLite database
    conn = sqlite3.connect(db_path)
    crm_df = pd.read_sql("SELECT companyName, companyAddress, companyCity, companyState, companyZipCode, systemId FROM crm", conn)
    conn.close()
    return crm_df

# Generic terms to deprioritize in name matches
GENERIC_TERMS = ['REGIONAL', 'MEDICAL', 'CENTER', 'HEALTH', 'HOSPITAL']

# Function to deprioritize generic terms in the match score
def deprioritize_generic_terms(name):
    score_penalty = 0
    for term in GENERIC_TERMS:
        if term in name.upper():
            score_penalty += 5  # Deduct points for common terms to lower their match priority
    return score_penalty

# Function to standardize and clean address
def standardize_address(address):
    if not isinstance(address, str):
        return ""
    address = address.upper()  # Convert to uppercase
    address = re.sub(r'[.,\-()/]', '', address)  # Remove punctuation
    address = re.sub(r'\bSTREET\b', 'ST', address)
    address = re.sub(r'\bROAD\b', 'RD', address)
    address = re.sub(r'\bBOULEVARD\b', 'BLVD', address)
    address = re.sub(r'\bDRIVE\b', 'DR', address)
    address = re.sub(r'\bAVENUE\b', 'AVE', address)
    address = re.sub(r'\bCOURT\b', 'CT', address)
    address = re.sub(r'\bLANE\b', 'LN', address)
    address = re.sub(r'\bCIRCLE\b', 'CIR', address)
    address = re.sub(r'\bPARKWAY\b', 'PKWY', address)
    address = re.sub(r'\bSUITE\b', 'STE', address)
    address = re.sub(r'\bBUILDING\b', 'BLDG', address)
    address = re.sub(r'\bHIGHWAY\b', 'HWY', address)
    address = re.sub(r'\s+', ' ', address)  # Remove extra spaces
    return address.strip()

# Function to clean and prepare text (general cleaning)
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

    # Step 2: Fuzzy match on name but enforce city match and deprioritize common terms
    crm_df['adjusted_name'] = crm_df['companyName'].apply(clean_text)
    
    # Perform fuzzy matching on names
    best_match_tuple = process.extractOne(query_name, crm_df['adjusted_name'].tolist(), scorer=fuzz.token_sort_ratio)

    if best_match_tuple and best_match_tuple[1] >= 85:  # Stricter threshold for fuzzy name match
        best_match_name = best_match_tuple[0]
        best_match_row = crm_df[crm_df['adjusted_name'] == best_match_name].iloc[0]

        # Only accept the match if the city matches exactly
        if query_city == clean_text(best_match_row['companyCity']):
            # Adjust score for generic terms in the match
            score_penalty = deprioritize_generic_terms(best_match_name)
            adjusted_score = best_match_tuple[1] - score_penalty
            
            return {
                'Match ID': best_match_row['systemId'],
                'Match Score': max(adjusted_score, 0),  # Ensure score doesn't go below 0
                'Matched Name': best_match_row['companyName'],
                'Matched Address': best_match_row['companyAddress'],
                'Matched City': best_match_row['companyCity'],
                'Matched State': best_match_row['companyState'],
                'Matched Zip': best_match_row['companyZipCode']
            }

    # If no match found
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

# Allow file path customization for CRM database
db_path = "C:/Users/ScottPhillips/OneDrive - Affinity Group/Desktop/Applications/Fuzzy Matching/crm_data.db"

# Load and cache CRM data
crm_df = load_crm_data(db_path)
st.write(f"CRM Data Loaded: {crm_df.shape[0]} rows")
st.dataframe(crm_df.head(500))  # Preview first 500 rows

# File uploader for the user’s file
uploaded_file = st.file_uploader("Upload your file for matching", type=["csv", "xlsx"])
if uploaded_file is not None:
    # Load the uploaded file (CSV or Excel)
    user_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Clean and standardize user data
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
