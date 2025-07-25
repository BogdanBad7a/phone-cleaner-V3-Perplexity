import streamlit as st
import pandas as pd
from phone_extractor import extract_uae_phone_numbers

st.title('UAE Phone Number Extractor')

uploaded_file = st.file_uploader('Upload your Excel file', type=['xls', 'xlsx'])

if uploaded_file:
    st.write('Processing...')
    df = extract_uae_phone_numbers(uploaded_file)
    st.success(f'Extracted {len(df)} unique UAE phone numbers.')
    st.dataframe(df)
    csv = df.to_csv(index=False).encode()
    st.download_button(
        label='Download Cleaned Phone Numbers CSV',
        data=csv,
        file_name='uae_phone_numbers.csv',
        mime='text/csv')
