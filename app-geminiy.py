import streamlit as st
import pandas as pd
import io
import os # For checking file existence

# --- Configuration & Constants ---
# Define file paths (assuming they are in the same directory as the script)
# IMPORTANT: Adjust these paths if your files are in a different location/subdirectory.
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) # Gets the directory of the current script
RAW_DATA_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
PRICE_MATRIX_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx")
MASTERDATA_TEMPLATE_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")

RAW_DATA_APP_SHEET = "APP"
PRICE_MATRIX_WHOLESALE_SHEET = "Price matrix wholesale"
PRICE_MATRIX_RETAIL_SHEET = "Price matrix retail"
DEFAULT_NO_SELECTION = "--- Please Select ---"


# --- Helper Function to Construct Product Display Name ---
def construct_product_display_name(row):
    """
    Constructs a display name for a product based on its type, model, and direction.
    Filters out "N/A" values.
    """
    name_parts = []
    product_type = row.get('Product Type')
    product_model = row.get('Product Model')
    sofa_direction = row.get('Sofa Direction')

    if pd.notna(product_type) and str(product_type).strip().upper() != "N/A":
        name_parts.append(str(product_type))
    if pd.notna(product_model) and str(product_model).strip().upper() != "N/A":
        name_parts.append(str(product_model))

    if str(product_type).strip().lower() == "sofa chaise longue":
        if pd.notna(sofa_direction) and str(sofa_direction).strip().upper() != "N/A":
            name_parts.append(str(sofa_direction))
    return " - ".join(name_parts) if name_parts else "Unnamed Product"

# --- Helper function to add selected combination ---
def add_to_selected_combinations(description, item_no, article_no, family, product, textile_family, textile_color, base_color_desc):
    """Adds a fully specified item to the session state list, avoiding duplicates."""
    is_duplicate = any(combo['item_no'] == item_no for combo in st.session_state.selected_combinations)
    if is_duplicate:
        # This message might be redundant if button is disabled, but good for programmatic calls
        # st.warning(f"Item No: {item_no} ({description}) is already in the list.")
        pass
    else:
        st.session_state.selected_combinations.append({
            "description": description,
            "item_no": item_no,
            "article_no": article_no,
            "family": family,
            "product": product,
            "textile_family": textile_family,
            "textile_color": textile_color,
            "base_color": base_color_desc
        })
        st.success(f"Added: {description} (Item No: {item_no})")


# --- Main App Logic ---
st.set_page_config(layout="wide")
st.title("Product Configurator & Masterdata Generator")

# --- Initialize session state variables ---
if 'selected_combinations' not in st.session_state:
    st.session_state.selected_combinations = []
if 'raw_df' not in st.session_state:
    st.session_state.raw_df = None
if 'wholesale_prices_df' not in st.session_state:
    st.session_state.wholesale_prices_df = None
if 'retail_prices_df' not in st.session_state:
    st.session_state.retail_prices_df = None
if 'template_cols' not in st.session_state:
    st.session_state.template_cols = None
if 'selected_family_session' not in st.session_state: # For remembering family selection
    st.session_state.selected_family_session = None


# --- Load Data Directly ---
files_loaded_successfully = True

if st.session_state.raw_df is None: # Load only once
    if os.path.exists(RAW_DATA_PATH):
        try:
            st.session_state.raw_df = pd.read_excel(RAW_DATA_PATH, sheet_name=RAW_DATA_APP_SHEET)
            st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
            # st.sidebar.success("Raw Data loaded.") # Sidebar messages might be less relevant now
        except Exception as e:
            st.error(f"Error loading Raw Data from '{RAW_DATA_PATH}' (Sheet: '{RAW_DATA_APP_SHEET}'): {e}")
            files_loaded_successfully = False
    else:
        st.error(f"Raw Data file not found at: {RAW_DATA_PATH}")
        files_loaded_successfully = False

if st.session_state.wholesale_prices_df is None: # Load only once
    if os.path.exists(PRICE_MATRIX_PATH):
        try:
            st.session_state.wholesale_prices_df = pd.read_excel(PRICE_MATRIX_PATH, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
            st.session_state.retail_prices_df = pd.read_excel(PRICE_MATRIX_PATH, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
            # st.sidebar.success("Price Matrix loaded.")
        except Exception as e:
            st.error(f"Error loading Price Matrix from '{PRICE_MATRIX_PATH}' (Sheets: '{PRICE_MATRIX_WHOLESALE_SHEET}', '{PRICE_MATRIX_RETAIL_SHEET}'): {e}")
            files_loaded_successfully = False
    else:
        st.error(f"Price Matrix file not found at: {PRICE_MATRIX_PATH}")
        files_loaded_successfully = False

if st.session_state.template_cols is None: # Load only once
    if os.path.exists(MASTERDATA_TEMPLATE_PATH):
        try:
            template_df = pd.read_excel(MASTERDATA_TEMPLATE_PATH)
            st.session_state.template_cols = template_df.columns.tolist()
            if "Wholesale price" not in st.session_state.template_cols:
                st.session_state.template_cols.append("Wholesale price")
            if "Retail price" not in st.session_state.template_cols:
                st.session_state.template_cols.append("Retail price")
            # st.sidebar.success("Masterdata Template loaded.")
        except Exception as e:
            st.error(f"Error loading Masterdata Template from '{MASTERDATA_TEMPLATE_PATH}': {e}")
            files_loaded_successfully = False
    else:
        st.error(f"Masterdata Template file not found at: {MASTERDATA_TEMPLATE_PATH}")
        files_loaded_successfully = False

# --- Main Application Area ---
if files_loaded_successfully and \
   st.session_state.raw_df is not None and \
   st.session_state.wholesale_prices_df is not None and \
   st.session_state.retail_prices_df is not None and \
   st.session_state.template_cols is not None:

    st.header("1. Select Product Combinations")

    # --- Product Family Selection ---
    available_families = [DEFAULT_NO_SELECTION] + sorted(st.session_state.raw_df['Product Family'].dropna().unique())
    
    # Use a key for the selectbox to help manage its state if needed, and store selection in session state
    selected_family = st.selectbox(
        "Select Product Family:",
        options=available_families,
        index=available_families.index(st.session_state.selected_family_session) if st.session_state.selected_family_session in available_families else 0,
        key="family_selector_main"
    )
    st.session_state.selected_family_session = selected_family # Update session state with current selection

    if selected_family and selected_family != DEFAULT_NO_SELECTION:
        family_df_all_items = st.session_state.raw_df[st.session_state.raw_df['Product Family'] == selected_family].copy()
        
        if 'Product Display Name' not in family_df_all_items.columns:
            st.error("'Product Display Name' column is missing. Please check the 'raw-data.xlsx' file and its processing.")
        else:
            products_in_family = sorted(family_df_all_items['Product Display Name'].dropna().unique())

            if not products_in_family:
                st.info(f"No products found for the family: {selected_family}")

            for product_name_iter in products_in_family:
                # Expander for each product
                with st.expander(f"Product: {product_name_iter}", expanded=False):
                    # Get all unique items for this product within the family
                    # Sorting by Item No for consistent display order within the product expander
                    product_items_df = family_df_all_items[
                        family_df_all_items['Product Display Name'] == product_name_iter
                    ].drop_duplicates(subset=['Item No']).sort_values(by=['Item No'])
                    
                    if product_items_df.empty:
                        st.write("No unique item configurations found for this product.")
                        continue

                    for idx, item_row in product_items_df.iterrows():
                        item_no = item_row['Item No']
                        article_no = item_row['Article No'] # Needed for price lookup
                        uph_type = item_row.get('Upholstery Type', "N/A")
                        uph_color = str(item_row.get('Upholstery Color', "N/A"))
                        base_color_val = str(item_row.get('Base Color', "N/A")) if pd.notna(item_row.get('Base Color')) else "N/A"
                        swatch_url = item_row.get('Image URL swatch')

                        # Construct description for display and for the selected_combinations list
                        desc_parts = [
                            selected_family,
                            product_name_iter,
                            uph_type,
                            uph_color
                        ]
                        # Only add "Base: color" if base_color_val is not "N/A"
                        if base_color_val.upper() != "N/A":
                            desc_parts.append(f"Base: {base_color_val}")
                        # else: # If N/A, it's implicitly handled by not adding it, or could add "Base: N/A"
                        #    desc_parts.append("Base: N/A") # Uncomment if explicit "Base: N/A" is desired in description

                        full_description = " / ".join(map(str, desc_parts))

                        # Layout for each item: Swatch | Details & Add Button
                        item_cols = st.columns([1, 3, 1.5]) # Swatch, Info, Button
                        
                        with item_cols[0]:
                            if pd.notna(swatch_url) and isinstance(swatch_url, str) and swatch_url.strip() != "":
                                st.image(swatch_url, width=60, caption=uph_color if len(uph_color) < 15 else "") # Short caption
                            else:
                                st.markdown(f"<div style='width:60px; height:60px; border:1px solid #ddd; display:flex; align-items:center; justify-content:center; font-size:0.8em; text-align:center;'>No Swatch</div>", unsafe_allow_html=True)
                        
                        with item_cols[1]:
                            st.markdown(f"""
                                **Textile:** {uph_type} - **Color:** {uph_color}<br>
                                **Base Color:** {base_color_val}<br>
                                <small><i>Item No: {item_no} / Article: {article_no}</i></small>
                            """, unsafe_allow_html=True)
                        
                        with item_cols[2]:
                            is_selected = any(sel_combo['item_no'] == item_no for sel_combo in st.session_state.selected_combinations)
                            button_label = "Added ✔️" if is_selected else "Add to List"
                            
                            if st.button(button_label, key=f"add_{item_no}", disabled=is_selected):
                                add_to_selected_combinations(
                                    full_description, 
                                    item_no, 
                                    article_no,
                                    selected_family,
                                    product_name_iter,
                                    uph_type,
                                    uph_color,
                                    base_color_val # Pass the actual base color value (could be "N/A")
                                )
                                st.experimental_rerun() # To update button state and selected list
                        st.markdown("---") # Visual separator between items within an expander

    # --- Display Current Selections ---
    if st.session_state.selected_combinations:
        st.header("2. Review Selected Combinations")
        for i, combo in enumerate(st.session_state.selected_combinations):
            col1, col2 = st.columns([0.9, 0.1])
            col1.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")
            if col2.button(f"Remove", key=f"remove_{i}_{combo['item_no']}"):
                st.session_state.selected_combinations.pop(i)
                st.experimental_rerun()
        
        st.header("3. Select Currency and Generate File")
        # --- Currency Selection ---
        try:
            article_no_col_name_ws = st.session_state.wholesale_prices_df.columns[0]
            currency_options = [col for col in st.session_state.wholesale_prices_df.columns if str(col).lower() != str(article_no_col_name_ws).lower()]
            
            if not currency_options:
                 st.error("No currency columns found in Price Matrix. Check column names (first column should be Article No).")
                 selected_currency = None
            else:
                selected_currency = st.selectbox("Select Currency:", options=currency_options, key="currency_selector")

        except Exception as e:
            st.error(f"Could not determine currency options: {e}")
            selected_currency = None

        if st.button("Generate Masterdata File", key="generate_file") and selected_currency:
            output_data = []
            master_template_columns_ordered = st.session_state.template_cols.copy()

            for combo_selection in st.session_state.selected_combinations:
                item_no_to_find = combo_selection['item_no']
                article_no_to_find = combo_selection['article_no'] 
                
                item_data_row_series_df = st.session_state.raw_df[st.session_state.raw_df['Item No'] == item_no_to_find]
                
                if not item_data_row_series_df.empty:
                    item_data_row_series = item_data_row_series_df.iloc[0]
                    output_row_dict = {}

                    for col_template in master_template_columns_ordered:
                        if col_template == "Wholesale price" or col_template == "Retail price":
                            continue
                        if col_template in item_data_row_series.index:
                            output_row_dict[col_template] = item_data_row_series[col_template]
                        else:
                            output_row_dict[col_template] = None 

                    ws_article_col = st.session_state.wholesale_prices_df.columns[0]
                    ws_price_row_df = st.session_state.wholesale_prices_df[
                        st.session_state.wholesale_prices_df[ws_article_col].astype(str) == str(article_no_to_find)
                    ]
                    if not ws_price_row_df.empty and selected_currency in ws_price_row_df.columns:
                        price_val = ws_price_row_df.iloc[0][selected_currency]
                        output_row_dict["Wholesale price"] = price_val if pd.notna(price_val) else "N/A in Price Matrix"
                    else:
                        output_row_dict["Wholesale price"] = "Price Not Found"

                    rt_article_col = st.session_state.retail_prices_df.columns[0]
                    rt_price_row_df = st.session_state.retail_prices_df[
                        st.session_state.retail_prices_df[rt_article_col].astype(str) == str(article_no_to_find)
                    ]
                    if not rt_price_row_df.empty and selected_currency in rt_price_row_df.columns:
                        price_val = rt_price_row_df.iloc[0][selected_currency]
                        output_row_dict["Retail price"] = price_val if pd.notna(price_val) else "N/A in Price Matrix"
                    else:
                        output_row_dict["Retail price"] = "Price Not Found"
                    
                    output_data.append(output_row_dict)
                else:
                    st.warning(f"Data for Item No: {item_no_to_find} not found in raw_data. Skipping.")

            if output_data:
                output_df = pd.DataFrame(output_data, columns=master_template_columns_ordered)
                output_excel_buffer = io.BytesIO()
                with pd.ExcelWriter(output_excel_buffer, engine='xlsxwriter') as writer:
                    output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
                output_excel_buffer.seek(0)

                st.download_button(
                    label="📥 Download Masterdata Excel File",
                    data=output_excel_buffer,
                    file_name=f"masterdata_output_{selected_currency.replace(' ', '_').replace('.', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No data to generate. Please add combinations to the list.")
        elif not selected_currency and st.session_state.selected_combinations : 
             st.warning("Please select a currency to enable file generation.")

    elif st.session_state.raw_df is not None: # Files loaded, but no family selected yet
        st.info("Select a Product Family to see available items.")

else:
    st.error("One or more data files could not be loaded. Please ensure the files (raw-data.xlsx, price-matrix_EUROPE.xlsx, Masterdata-output-template.xlsx) exist in the same directory as the script and are correctly formatted.")
    st.info(f"Expected raw data at: {RAW_DATA_PATH}")
    st.info(f"Expected price matrix at: {PRICE_MATRIX_PATH}")
    st.info(f"Expected template at: {MASTERDATA_TEMPLATE_PATH}")


# --- Styling (Optional) ---
st.markdown("""
<style>
    /* Style for expander headers to make them more prominent */
    .st-expanderHeader {
        font-size: 1.1em;
        font-weight: bold;
        background-color: #f0f2f6; /* Light background for expander header */
        border-radius: 5px;
        padding: 8px !important; /* Override default padding if necessary */
    }
    div[data-testid="stImage"] img {
        object-fit: contain;
        max-height: 60px; /* Adjusted swatch size */
        border: 1px solid #eee;
        border-radius: 4px;
        padding: 2px;
        margin: auto;
        display: block;
    }
    div[data-testid="stImage"] figcaption {
        text-align: center;
        font-size: 0.8em;
        padding-top: 2px;
    }
    .stButton>button {
        width: auto;
        padding: 0.3em 0.6em; /* Smaller padding for buttons */
        font-size: 0.9em; /* Smaller font for buttons */
    }
    hr { /* Style for the st.markdown("---") separator */
        margin-top: 0.5rem;
        margin-bottom: 0.5rem;
        border-top: 1px solid #ddd;
    }
</style>
""", unsafe_allow_html=True)
