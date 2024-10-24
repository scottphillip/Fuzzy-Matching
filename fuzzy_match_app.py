import pandas as pd
import streamlit as st
import sqlite3
import re
from rapidfuzz import process, fuzz  # Use rapidfuzz for faster processing
from concurrent.futures import ThreadPoolExecutor
import platform

# Function to prevent sleep mode (Windows only)
def prevent_sleep():
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

# Function to allow system sleep (Windows only)
def allow_sleep():
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

# Prevent sleep at the start of the script
prevent_sleep()

# Function to load CRM data from the static default path
@st.cache_data
def load_crm_data():
    db_path = "crm_data.db"  # Relative path to the file in your GitHub repository
    st.write(f"Trying to connect to database at: {db_path}")
    
    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(db_path)
        
        # Query to retrieve necessary columns
        crm_df = pd.read_sql("SELECT companyName, companyAddress, companyCity, companyState, companyZipCode, systemId FROM crm", conn)
        
        # Close the database connection
        conn.close()
        
        # Return the CRM dataframe
        return crm_df
    
    except sqlite3.OperationalError as e:
        # Display an error if the database couldn't be opened
        st.error(f"Error opening database: {e}")
        # Return an empty dataframe on failure
        return pd.DataFrame()

# Function to standardize and clean address
def standardize_address(address):
    address = address.upper()
    address = re.sub(r'\.', '', address)
    address = re.sub(r',', '', address)
    address = re.sub(r'-', '', address)
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
    address = re.sub(r'\bGROUND\b', 'GDS', address)
    address = re.sub(r'\bHIGHWAY\b', 'HWY', address)
    address = re.sub(r'\bPLACE\b', 'PL', address)

    # Directional standardizations
    address = re.sub(r'\bNORTH\b', 'N', address)
    address = re.sub(r'\bSOUTH\b', 'S', address)
    address = re.sub(r'\bEAST\b', 'E', address)
    address = re.sub(r'\bWEST\b', 'W', address)
    address = re.sub(r'\bNORTHEAST\b', 'NE', address)
    address = re.sub(r'\bNORTHWEST\b', 'NW', address)
    address = re.sub(r'\bSOUTHEAST\b', 'SE', address)
    address = re.sub(r'\bSOUTHWEST\b', 'SW', address)

    # Remove common suffixes (St, Rd, etc.)
    address = re.sub(r'\bST\b|\bRD\b|\bBLVD\b|\bDR\b|\bAVE\b|\bCT\b|\bLN\b|\bCIR\b|\bPKWY\b|\bHWY\b', '', address)

    # Remove extra spaces
    address = re.sub(r'\s+', ' ', address).strip()

    return address

# Function to clean general text
def clean_text(text):
    return ' '.join(str(text).upper().split())

# Function to match based on name and address
def get_best_match(row, crm_df):
    query_name = clean_text(row['companyName'])
    query_address = standardize_address(row['companyAddress'])  # Standardize address here
    query_city = clean_text(row['companyCity'])
    query_state = clean_text(row['companyState'])

    # Exact address match criteria: street and city must match after standardization
    for _, crm_row in crm_df.iterrows():
        crm_address = standardize_address(crm_row['companyAddress'])
        crm_city = clean_text(crm_row['companyCity'])
        crm_state = clean_text(crm_row['companyState'])
        
        if query_address == crm_address and query_city == crm_city:
            return {
                'Match ID': crm_row['systemId'],
                'Match Score': 100,  # Perfect match for address
                'Matched Name': crm_row['companyName'],
                'Matched Address': crm_row['companyAddress'],
                'Matched City': crm_row['companyCity'],
                'Matched State': crm_row['companyState'],
                'Matched Zip': crm_row['companyZipCode']
            }

    # If no exact address match, use fuzzy name matching but require city match
    best_match_tuple = process.extractOne(f"{query_name} {query_city}", crm_df['combined'], scorer=fuzz.token_sort_ratio)

    if best_match_tuple:
        best_match, score = best_match_tuple[0], best_match_tuple[1]
        best_match_row = crm_df.loc[crm_df['combined'] == best_match].iloc[0]

        # Ensure fuzzy match score >= 70 and city matches
        if score >= 70 and query_city == clean_text(best_match_row['companyCity']):
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

# Load CRM data
crm_df = load_crm_data()
crm_df['combined'] = crm_df['companyName'] + ' ' + crm_df['companyCity'] + ' ' + crm_df['companyState']
st.write(f"CRM Data Loaded: {crm_df.shape[0]} rows")
st.dataframe(crm_df.head(500))  # Preview first 500 rows

# File uploader for the user's file
uploaded_file = st.file_uploader("Upload your file for matching", type=["csv", "xlsx"])
if uploaded_file is not None:
    # Load the uploaded file (CSV or Excel)
    user_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

    # Clean and standardize the text data
    user_df['companyName'] = user_df['companyName'].apply(clean_text)
    user_df['companyAddress'] = user_df['companyAddress'].apply(standardize_address)  # Standardize addresses here
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
