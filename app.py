import streamlit as st
import pandas as pd

# Load Excel files
@st.cache_data
def load_data():
    raw = pd.read_excel("raw-data.xlsx", sheet_name="APP")
    wholesale = pd.read_excel("price-matrix_EUROPE.xlsx", sheet_name="Price matrix wholesale")
    retail = pd.read_excel("price-matrix_EUROPE.xlsx", sheet_name="Price matrix retail")
    template = pd.read_excel("Masterdata-output-template.xlsx")
    return raw, wholesale, retail, template

raw_data, wholesale_prices, retail_prices, output_template = load_data()

# Intro
st.title("Muuto Configurator Tool")
st.markdown("""
Welcome to the Muuto configurator.

This tool helps you:
1. Select product families and textile combinations
2. Choose currency
3. Download a complete masterdata file with prices.

Let’s begin!
""")
