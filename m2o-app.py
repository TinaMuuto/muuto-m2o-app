import streamlit as st
import pandas as pd
import io
import os

# --- Page Configuration (MUST BE THE FIRST STREAMLIT COMMAND) ---
st.set_page_config(
    layout="wide",
    page_title="Muuto M2O",
    page_icon="favicon.png"  # Ensure this file exists or remove/replace
)

# --- Configuration & Constants ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_XLSX_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
PRICE_MATRIX_EUROPE_XLSX_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx")
PRICE_MATRIX_UK_XLSX_PATH = os.path.join(BASE_DIR, "price-matrix_UK-EI.xlsx") # Path for UK/EI prices
MASTERDATA_TEMPLATE_XLSX_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")

RAW_DATA_APP_SHEET = "APP"
PRICE_MATRIX_WHOLESALE_SHEET = "Price matrix wholesale"
PRICE_MATRIX_RETAIL_SHEET = "Price matrix retail"

DEFAULT_NO_SELECTION = "--- Please Select ---"

# --- Helper Function to Construct Product Display Name ---
def construct_product_display_name(row):
    name_parts = []
    product_type = row.get('Product Type')
    product_model = row.get('Product Model')
    sofa_direction = row.get('Sofa Direction')
    if pd.notna(product_type) and str(product_type).strip().upper() != "N/A": name_parts.append(str(product_type))
    if pd.notna(product_model) and str(product_model).strip().upper() != "N/A": name_parts.append(str(product_model))
    if str(product_type).strip().lower() == "sofa chaise longue":
        if pd.notna(sofa_direction) and str(sofa_direction).strip().upper() != "N/A": name_parts.append(str(sofa_direction))
    return " - ".join(name_parts) if name_parts else "Unnamed Product"

# --- Main App Logic ---

# --- Logo and Title Section ---
top_col1, top_col_spacer, top_col2 = st.columns([5.5, 0.5, 1])

with top_col1:
    # Updated Title
    st.title("Welcome to the Muuto M2O master data generator")

with top_col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)
    else:
        st.error(f"Muuto Logo not found. Expected at: {LOGO_PATH}. Please ensure 'muuto_logo.png' is in the script's directory.")

# --- App Introduction (Updated) ---
st.markdown("""
Select your preferred M2O sofa combinations — and instantly generate all the data you need.

Here’s how it works:
1.  **Select your currency**
    Start by selecting your preferred currency. This ensures that your tailored data pack includes recommended retail prices aligned with your market.
2.  **Explore your options and choose your sofacombinations**
    Browse the curated made-to-order variants across sofa families, configurations, textiles and (where relevant) base colors. Simply tick off the combinations you'd like to include — it’s guided, visual and easy to use.
3.  **Review your list**
    Once you’ve made your selections, scroll down to review your full list. You can remove items directly from here if needed.
4.  **Download and add to your assortment**
    Click ‘Generate’ to instantly download a file with all the essential details: product names, item numbers, pricing, packshots, textile info, textile swatches and more. Use the data pack to upload your new M2O variants to your webshop – and expand beyond Ready-to-Ship.
""")

# --- Initialize session state variables ---
if 'raw_df_original' not in st.session_state: st.session_state.raw_df_original = None
if 'raw_df' not in st.session_state: st.session_state.raw_df = None # This will hold the filtered data
if 'wholesale_prices_df' not in st.session_state: st.session_state.wholesale_prices_df = None
if 'retail_prices_df' not in st.session_state: st.session_state.retail_prices_df = None
if 'template_cols' not in st.session_state: st.session_state.template_cols = None
if 'selected_family_session' not in st.session_state: st.session_state.selected_family_session = None
if 'matrix_selected_generic_items' not in st.session_state: st.session_state.matrix_selected_generic_items = {}
if 'user_chosen_base_colors_for_items' not in st.session_state: st.session_state.user_chosen_base_colors_for_items = {}
if 'final_items_for_download' not in st.session_state: st.session_state.final_items_for_download = []
if 'selected_currency_session' not in st.session_state: st.session_state.selected_currency_session = None

# --- Load Data Directly from XLSX files ---
files_loaded_successfully = True

@st.cache_data # Use Streamlit's caching for data loading
def load_data():
    """
    Loads all necessary data files and performs initial processing.
    Returns:
        tuple: Contains loaded dataframes (raw_df_original, wholesale_prices_df, retail_prices_df, template_cols)
               and a list of data_load_errors.
    """
    raw_df_original = None
    wholesale_prices_df = None
    retail_prices_df = None
    template_cols = None
    data_load_errors = []

    # Load Raw Data
    if os.path.exists(RAW_DATA_XLSX_PATH):
        try:
            raw_df_original = pd.read_excel(RAW_DATA_XLSX_PATH, sheet_name=RAW_DATA_APP_SHEET)
            required_cols = ['Product Type', 'Product Model', 'Sofa Direction', 'Base Color', 'Product Family', 'Item No', 'Article No', 'Image URL swatch', 'Upholstery Type', 'Upholstery Color', 'Market']
            missing = [col for col in required_cols if col not in raw_df_original.columns]
            if missing:
                data_load_errors.append(f"Required columns are missing in '{os.path.basename(RAW_DATA_XLSX_PATH)}': {missing}.")
            else:
                # Apply transformations
                raw_df_original['Product Display Name'] = raw_df_original.apply(construct_product_display_name, axis=1)
                raw_df_original['Base Color Cleaned'] = raw_df_original['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)
                raw_df_original['Upholstery Type'] = raw_df_original['Upholstery Type'].astype(str).str.strip()
        except Exception as e:
            data_load_errors.append(f"Error loading Raw Data from '{RAW_DATA_XLSX_PATH}': {e}")
    else:
        data_load_errors.append(f"Raw Data file not found at: {RAW_DATA_XLSX_PATH}")

    # Load Price Matrices
    price_matrices_to_load = {
        'europe': PRICE_MATRIX_EUROPE_XLSX_PATH,
        'uk_ei': PRICE_MATRIX_UK_XLSX_PATH  # Added UK/EI price matrix
    }
    ws_dfs = {}
    rt_dfs = {}

    for market_key, path in price_matrices_to_load.items():
        if os.path.exists(path):
            try:
                ws_dfs[market_key] = pd.read_excel(path, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
                rt_dfs[market_key] = pd.read_excel(path, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
            except Exception as e:
                data_load_errors.append(f"Error loading price data from {os.path.basename(path)}: {e}")
        else:
            # Only error if the EUROPE one is missing, as UK/EI is an addition and might not always be present
            if market_key == 'europe':
                data_load_errors.append(f"Required Price Matrix file not found: {os.path.basename(path)}")
            elif market_key == 'uk_ei' and os.path.exists(PRICE_MATRIX_UK_XLSX_PATH): # only warn if it was expected
                 st.warning(f"Optional Price Matrix file for UK/EI not found at: {os.path.basename(path)}, proceeding without it.")


    # Merge price dataframes if they exist
    if ws_dfs:
        # Start with Europe if available, otherwise take UK/EI if it's the only one
        base_ws_df = ws_dfs.get('europe')
        if base_ws_df is None and 'uk_ei' in ws_dfs:
             base_ws_df = ws_dfs.get('uk_ei')
        elif base_ws_df is not None and 'uk_ei' in ws_dfs:
            article_no_col = base_ws_df.columns[0] # Assume first col is Article No
            uk_ws_df = ws_dfs['uk_ei']
            # Ensure the key column has the same name for merging
            uk_ws_df.rename(columns={uk_ws_df.columns[0]: article_no_col}, inplace=True)
            # Exclude common columns from the second df before merge, except for the key
            uk_cols_to_use = [article_no_col] + [col for col in uk_ws_df.columns if col not in base_ws_df.columns or col == article_no_col]
            base_ws_df = pd.merge(base_ws_df, uk_ws_df[uk_cols_to_use].drop_duplicates(subset=[article_no_col]), on=article_no_col, how='outer')
        wholesale_prices_df = base_ws_df

    if rt_dfs:
        base_rt_df = rt_dfs.get('europe')
        if base_rt_df is None and 'uk_ei' in rt_dfs:
            base_rt_df = rt_dfs.get('uk_ei')
        elif base_rt_df is not None and 'uk_ei' in rt_dfs:
            article_no_col = base_rt_df.columns[0]
            uk_rt_df = rt_dfs['uk_ei']
            uk_rt_df.rename(columns={uk_rt_df.columns[0]: article_no_col}, inplace=True)
            uk_cols_to_use = [article_no_col] + [col for col in uk_rt_df.columns if col not in base_rt_df.columns or col == article_no_col]
            base_rt_df = pd.merge(base_rt_df, uk_rt_df[uk_cols_to_use].drop_duplicates(subset=[article_no_col]), on=article_no_col, how='outer')
        retail_prices_df = base_rt_df


    # Load Template
    if os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH):
        try:
            template_cols = pd.read_excel(MASTERDATA_TEMPLATE_XLSX_PATH).columns.tolist()
            # Ensure price columns are in the template, add if missing
            if "Wholesale price" not in template_cols: template_cols.append("Wholesale price")
            if "Retail price" not in template_cols: template_cols.append("Retail price")
        except Exception as e:
            data_load_errors.append(f"Error loading Template from '{MASTERDATA_TEMPLATE_XLSX_PATH}': {e}")
    else:
        data_load_errors.append(f"Template file not found at: {MASTERDATA_TEMPLATE_XLSX_PATH}")

    return raw_df_original, wholesale_prices_df, retail_prices_df, template_cols, data_load_errors

# Load data and manage success flag
raw_df_original_loaded, wholesale_prices_df_loaded, retail_prices_df_loaded, template_cols_loaded, data_load_errors_list = load_data()

if data_load_errors_list:
    for error in data_load_errors_list:
        st.error(error)
    files_loaded_successfully = False # Set to false if any error occurred
else:
    st.session_state.raw_df_original = raw_df_original_loaded
    st.session_state.wholesale_prices_df = wholesale_prices_df_loaded
    st.session_state.retail_prices_df = retail_prices_df_loaded
    st.session_state.template_cols = template_cols_loaded
    files_loaded_successfully = True


# --- Main Application Area ---
if files_loaded_successfully:
    
    # --- Step 1 (New): Select Currency ---
    st.header("Step 1: Select your currency")
    
    def on_currency_change():
        # Clear selections when currency changes, as available products might change
        st.session_state.selected_family_session = DEFAULT_NO_SELECTION # Reset family
        st.session_state.matrix_selected_generic_items = {}
        st.session_state.user_chosen_base_colors_for_items = {}
        st.session_state.final_items_for_download = []


    try:
        if st.session_state.wholesale_prices_df is not None and not st.session_state.wholesale_prices_df.empty:
            article_no_col_name_ws = st.session_state.wholesale_prices_df.columns[0] # Assumed first column
            currency_options = [DEFAULT_NO_SELECTION] + sorted([
                col for col in st.session_state.wholesale_prices_df.columns 
                if str(col).lower() != str(article_no_col_name_ws).lower() and str(col).strip() != ""
            ])
        else:
            currency_options = [DEFAULT_NO_SELECTION]
            st.error("Wholesale price matrix is empty or could not be loaded. Currency selection is not possible.")

        current_currency_idx = 0
        if st.session_state.selected_currency_session and st.session_state.selected_currency_session in currency_options:
            current_currency_idx = currency_options.index(st.session_state.selected_currency_session)
        else: # If current selection is invalid, reset to default
             st.session_state.selected_currency_session = None


        selected_currency_choice = st.selectbox(
            "Select Currency:",
            options=currency_options,
            index=current_currency_idx,
            key="currency_selector_main_final",
            on_change=on_currency_change # Callback to reset things if currency changes
        )
        
        if selected_currency_choice != DEFAULT_NO_SELECTION:
            st.session_state.selected_currency_session = selected_currency_choice
        else:
            st.session_state.selected_currency_session = None # Explicitly set to None if default is chosen
    
    except Exception as e:
        st.error(f"Error with currency selection: {e}")
        st.session_state.selected_currency_session = None

    # Filter raw_df based on currency selection for product display
    if st.session_state.selected_currency_session and st.session_state.raw_df_original is not None:
        selected_curr_upper = st.session_state.selected_currency_session.upper()
        if selected_curr_upper in ['GBP', 'EI', 'IE']: # EI and IE for Ireland
            st.session_state.raw_df = st.session_state.raw_df_original[
                st.session_state.raw_df_original['Market'].astype(str).str.upper() == 'UK'
            ].copy()
        else: # For all other currencies, filter out 'UK'
            st.session_state.raw_df = st.session_state.raw_df_original[
                st.session_state.raw_df_original['Market'].astype(str).str.upper() != 'UK'
            ].copy()
    elif st.session_state.raw_df_original is not None: # If no currency selected, but raw data is loaded
        st.session_state.raw_df = None # No products to show yet
    else: # If raw_df_original itself is None
        st.session_state.raw_df = None


    # Proceed only if currency is selected AND raw_df has been processed
    if st.session_state.selected_currency_session and st.session_state.raw_df is not None:

        st.markdown("---")
        # --- Step 2: Explore your options and choose your sofacombinations ---
        st.header("Step 2: Explore your options and choose your sofacombinations")

        df_for_display = st.session_state.raw_df # Use the filtered df

        available_families_in_view = [DEFAULT_NO_SELECTION] + sorted(df_for_display['Product Family'].dropna().unique()) if 'Product Family' in df_for_display.columns else [DEFAULT_NO_SELECTION]
        
        # Ensure selected_family_session is valid for the current df_for_display
        if st.session_state.selected_family_session not in available_families_in_view:
            st.session_state.selected_family_session = DEFAULT_NO_SELECTION

        selected_family_idx = 0
        if st.session_state.selected_family_session in available_families_in_view: # Check if it's a valid option
            selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)
        else: # If not valid (e.g. after currency change), reset to default
            st.session_state.selected_family_session = DEFAULT_NO_SELECTION


        selected_family = st.selectbox("Select Product Family:", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
        st.session_state.selected_family_session = selected_family # Update session state

        def handle_matrix_cb_toggle(prod_name, uph_type, uph_color, checkbox_key_matrix):
            is_checked = st.session_state[checkbox_key_matrix]
            # Ensure selected_family from the widget is used for the key
            current_selected_family_for_key = st.session_state.selected_family_session 
            generic_item_key = f"{current_selected_family_for_key}_{prod_name}_{uph_type}_{uph_color}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")

            if is_checked:
                # Use the currently displayed (potentially filtered) raw_df for matching
                matching_items = st.session_state.raw_df[  
                    (st.session_state.raw_df['Product Family'] == current_selected_family_for_key) &
                    (st.session_state.raw_df['Product Display Name'] == prod_name) &
                    (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == uph_type) &
                    (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == uph_color)
                ]
                if not matching_items.empty:
                    unique_base_colors = matching_items['Base Color Cleaned'].dropna().unique().tolist()
                    first_item_match = matching_items.iloc[0]

                    item_data = {
                        'key': generic_item_key, 'family': current_selected_family_for_key, 'product': prod_name,
                        'upholstery_type': uph_type, 'upholstery_color': uph_color,
                        'requires_base_choice': len(unique_base_colors) > 1,
                        'available_bases': unique_base_colors if len(unique_base_colors) > 1 else [],
                        'item_no_if_single_base': first_item_match['Item No'] if len(unique_base_colors) <= 1 else None,
                        'article_no_if_single_base': first_item_match['Article No'] if len(unique_base_colors) <= 1 else None,
                        'resolved_base_if_single': unique_base_colors[0] if len(unique_base_colors) == 1 else (pd.NA if not unique_base_colors and len(unique_base_colors) == 0 else None) # Handle no base vs single base
                    }
                    st.session_state.matrix_selected_generic_items[generic_item_key] = item_data
                    st.toast(f"Selected: {prod_name} / {uph_type} / {uph_color}", icon="➕")
            else: # If unchecked
                if generic_item_key in st.session_state.matrix_selected_generic_items:
                    del st.session_state.matrix_selected_generic_items[generic_item_key]
                    # Also remove any base color choices associated with this generic item
                    if generic_item_key in st.session_state.user_chosen_base_colors_for_items:
                        del st.session_state.user_chosen_base_colors_for_items[generic_item_key]
                    st.toast(f"Deselected: {prod_name} / {uph_type} / {uph_color}", icon="➖")

        def handle_base_color_multiselect_change(item_key_for_base_select):
            multiselect_widget_key = f"ms_base_{item_key_for_base_select}"
            st.session_state.user_chosen_base_colors_for_items[item_key_for_base_select] = st.session_state[multiselect_widget_key]

        # Display matrix if a family is selected
        if selected_family and selected_family != DEFAULT_NO_SELECTION and 'Product Family' in df_for_display.columns:
            family_df = df_for_display[df_for_display['Product Family'] == selected_family]
            if not family_df.empty and 'Upholstery Type' in family_df.columns:
                products_in_family = sorted(family_df['Product Display Name'].dropna().unique())
                upholstery_types_in_family = sorted(family_df['Upholstery Type'].dropna().unique())

                if not products_in_family: st.info(f"No products found in the family: {selected_family} for the selected market/currency.")
                elif not upholstery_types_in_family: st.info(f"No upholstery types found for the product family: {selected_family} for the selected market/currency.")
                else:
                    # Matrix Display Logic 
                    header_upholstery_types = ["Product"]
                    header_swatches = [" "] # Placeholder for the first column (Product Name)
                    header_color_numbers = [" "] # Placeholder
                    data_column_map = [] # To map matrix columns back to uph_type and uph_color

                    for uph_type_clean in upholstery_types_in_family:
                        # Get unique colors and their swatches for this upholstery type WITHIN the current family_df
                        colors_for_type_df = family_df[family_df['Upholstery Type'] == uph_type_clean][['Upholstery Color', 'Image URL swatch']].drop_duplicates().sort_values(by='Upholstery Color')
                        if not colors_for_type_df.empty:
                            # For the Upholstery Type header, it spans all its colors
                            header_upholstery_types.extend([uph_type_clean] + [""] * (len(colors_for_type_df) -1) )
                            for _, color_row in colors_for_type_df.iterrows():
                                color_val = str(color_row['Upholstery Color'])
                                swatch_val = color_row['Image URL swatch']
                                header_swatches.append(swatch_val if pd.notna(swatch_val) else None)
                                header_color_numbers.append(color_val)
                                data_column_map.append({'uph_type': uph_type_clean, 'uph_color': color_val, 'swatch': swatch_val})
                    
                    num_data_columns = len(data_column_map)
                    if num_data_columns == 0:
                        st.info(f"No upholstery/color combinations to display for the family: {selected_family}")
                    else:
                        # --- Render Matrix Headers ---
                        # Upholstery Type Header
                        cols_uph_type_header = st.columns([2.5] + [1] * num_data_columns)
                        current_uph_type_header_display = None # To span header correctly
                        for i, col_widget in enumerate(cols_uph_type_header):
                            if i == 0: # First column is for Product Name, so empty caption
                                with col_widget: st.caption("")
                            else:
                                map_entry = data_column_map[i-1] # data_column_map is 0-indexed for data cols
                                if map_entry['uph_type'] != current_uph_type_header_display:
                                    with col_widget: st.caption(f"<div class='upholstery-header'>{map_entry['uph_type']}</div>", unsafe_allow_html=True)
                                    current_uph_type_header_display = map_entry['uph_type']
                                # else, it's spanned by the previous type header, so do nothing for this cell

                        # Swatch Header
                        cols_swatch_header = st.columns([2.5] + [1] * num_data_columns)
                        for i, col_widget in enumerate(cols_swatch_header):
                            if i == 0:
                                with col_widget: st.markdown("<div class='zoom-instruction'><br>Click swatch in header to zoom</div>", unsafe_allow_html=True)
                            else:
                                sw_url = data_column_map[i-1]['swatch']
                                with col_widget:
                                    if sw_url and pd.notna(sw_url): st.image(sw_url, width=30)
                                    else: st.markdown("<div class='swatch-placeholder'></div>", unsafe_allow_html=True)
                        
                        # Color Number Header
                        cols_color_num_header = st.columns([2.5] + [1] * num_data_columns)
                        for i, col_widget in enumerate(cols_color_num_header):
                            if i == 0:
                                with col_widget: st.caption("")
                            else:
                                with col_widget: st.caption(f"<small>{data_column_map[i-1]['uph_color']}</small>", unsafe_allow_html=True)

                        st.markdown("---") # Divider after headers

                        # --- Render Matrix Rows (Products and Checkboxes) ---
                        for prod_name in products_in_family:
                            cols_product_row = st.columns([2.5] + [1] * num_data_columns, vertical_alignment="center")
                            # Product Name Cell
                            with cols_product_row[0]:
                                st.markdown(f"<div class='product-name-cell'>{prod_name}</div>", unsafe_allow_html=True)

                            # Checkbox Cells
                            for i, col_widget in enumerate(cols_product_row[1:]): # Start from the second col in cols_product_row
                                current_col_uph_type_filter = data_column_map[i]['uph_type']
                                current_col_uph_color_filter = data_column_map[i]['uph_color']

                                # Check if this specific product/upholstery/color combination exists in the family_df
                                item_exists_df = family_df[
                                    (family_df['Product Display Name'] == prod_name) &
                                    (family_df['Upholstery Type'] == current_col_uph_type_filter) &
                                    (family_df['Upholstery Color'].astype(str).fillna("N/A") == current_col_uph_color_filter)
                                ]
                                
                                cell_container = col_widget.container() # Use a container for each cell

                                if not item_exists_df.empty:
                                    # Construct a unique key for the checkbox and for the generic item
                                    cb_key_str = f"cb_{selected_family}_{prod_name}_{current_col_uph_type_filter}_{current_col_uph_color_filter}".replace(" ","_").replace("/","_").replace("(","").replace(")","")
                                    generic_item_key_for_check = f"{selected_family}_{prod_name}_{current_col_uph_type_filter}_{current_col_uph_color_filter}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")
                                    is_gen_selected = generic_item_key_for_check in st.session_state.matrix_selected_generic_items

                                    cell_container.checkbox(" ", value=is_gen_selected, key=cb_key_str,
                                                            on_change=handle_matrix_cb_toggle,
                                                            args=(prod_name, current_col_uph_type_filter, current_col_uph_color_filter, cb_key_str),
                                                            label_visibility="collapsed")
                                else:
                                    # If item combination doesn't exist, cell is empty (no checkbox)
                                    pass 
            else: # family_df is empty or no Upholstery Type
                if selected_family and selected_family != DEFAULT_NO_SELECTION : st.info(f"No data found for product family: {selected_family} with the current currency/market selection.")
        
        # --- Base Color Specification (Sub-step of Step 2) ---
        items_needing_base_choice_now = [
            item_data for key, item_data in st.session_state.matrix_selected_generic_items.items() if item_data.get('requires_base_choice')
        ]
        if items_needing_base_choice_now:
            st.subheader("Specify base colors for selected items") 
            for generic_item in items_needing_base_choice_now:
                item_key = generic_item['key']
                multiselect_key = f"ms_base_{item_key}" # Unique key for multiselect widget
                st.markdown(f"**{generic_item['product']}** ({generic_item['upholstery_type']} - {generic_item['upholstery_color']})")

                # Get current selections for this item, default to empty list if not yet chosen
                current_selection_for_this_item = st.session_state.user_chosen_base_colors_for_items.get(item_key, [])
                
                # Filter available_bases to ensure they are valid options
                valid_bases = [base for base in generic_item['available_bases'] if pd.notna(base)]


                st.multiselect(
                    f"Available base colors. You can select multiple:",
                    options=valid_bases,
                    default=current_selection_for_this_item,
                    key=multiselect_key,
                    on_change=handle_base_color_multiselect_change,
                    args=(item_key,) # Pass item_key to callback
                )
                st.markdown("---")

        # --- Step 3: Review Selections ---
        st.header("Step 3: Review your list")
        _current_final_items = [] # Temporary list to build final items for review
        for key, gen_item_data in st.session_state.matrix_selected_generic_items.items():
            if not gen_item_data['requires_base_choice']: # Single base or N/A base
                if gen_item_data.get('item_no_if_single_base') is not None:
                    desc_base_part = ""
                    if pd.notna(gen_item_data['resolved_base_if_single']) and str(gen_item_data['resolved_base_if_single']).strip().upper() != "N/A":
                        desc_base_part = f" / Base: {gen_item_data['resolved_base_if_single']}"

                    _current_final_items.append({
                        "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']}{desc_base_part}",
                        "item_no": gen_item_data['item_no_if_single_base'],
                        "article_no": gen_item_data['article_no_if_single_base'],
                        "key_in_matrix": key # Link back to the generic item
                    })
            else: # Requires base choice
                selected_bases_for_this = st.session_state.user_chosen_base_colors_for_items.get(key, [])
                for bc in selected_bases_for_this:
                    # Find the specific item in the currently displayed raw_df
                    specific_item_df = st.session_state.raw_df[
                        (st.session_state.raw_df['Product Family'] == gen_item_data['family']) &
                        (st.session_state.raw_df['Product Display Name'] == gen_item_data['product']) &
                        (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == gen_item_data['upholstery_type']) &
                        (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == gen_item_data['upholstery_color']) &
                        (st.session_state.raw_df['Base Color Cleaned'].fillna("N/A") == bc) # Match chosen base
                    ]
                    if not specific_item_df.empty:
                        actual_item = specific_item_df.iloc[0]
                        _current_final_items.append({
                            "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']} / Base: {bc}",
                            "item_no": actual_item['Item No'],
                            "article_no": actual_item['Article No'],
                            "key_in_matrix": key, # Link back to generic
                            "chosen_base": bc    # Store the chosen base for this specific final item
                        })

        # Deduplicate final items before storing in session state (based on item_no and chosen_base if applicable)
        temp_final_list_review = []
        seen_item_keys_review = set() # To track unique final items
        for item_rev in _current_final_items:
            # Create a unique key for the final item (Item No + chosen base if it exists)
            unique_final_item_key = f"{item_rev['item_no']}_{item_rev.get('chosen_base', 'NO_BASE_CHOSEN')}"
            if unique_final_item_key not in seen_item_keys_review:
                temp_final_list_review.append(item_rev)
                seen_item_keys_review.add(unique_final_item_key)
        st.session_state.final_items_for_download = temp_final_list_review


        if st.session_state.final_items_for_download:
            st.markdown("**Current Selections for Download:**")
            for i, combo in enumerate(st.session_state.final_items_for_download):
                col1_rev, col2_rev = st.columns([0.9, 0.1]) # Description and Remove button
                col1_rev.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")
                
                # Unique key for the remove button
                remove_button_key = f"final_review_remove_{i}_{combo['item_no']}_{combo.get('chosen_base','nobase')}"

                if col2_rev.button(f"Remove", key=remove_button_key):
                    original_matrix_key = combo['key_in_matrix'] # Key of the generic item
                    
                    if original_matrix_key in st.session_state.matrix_selected_generic_items:
                        # If it was an item requiring base choice and had a specific base selected
                        if st.session_state.matrix_selected_generic_items[original_matrix_key].get('requires_base_choice') and 'chosen_base' in combo:
                            chosen_base_to_remove = combo['chosen_base']
                            if original_matrix_key in st.session_state.user_chosen_base_colors_for_items:
                                if chosen_base_to_remove in st.session_state.user_chosen_base_colors_for_items[original_matrix_key]:
                                    st.session_state.user_chosen_base_colors_for_items[original_matrix_key].remove(chosen_base_to_remove)
                                    # If no bases are left selected for this generic item, deselect the generic item itself
                                    if not st.session_state.user_chosen_base_colors_for_items[original_matrix_key]:
                                        del st.session_state.user_chosen_base_colors_for_items[original_matrix_key] # Clean up empty list
                                        # Check if the generic item should be removed entirely (no other bases selected)
                                        # This logic might need refinement if we want to keep the generic item selected even if all its bases are removed from review
                                        # For now, if all bases are removed, the generic item is also removed.
                                        del st.session_state.matrix_selected_generic_items[original_matrix_key]

                        else: # Item did not require base choice, or was a single-base item
                            del st.session_state.matrix_selected_generic_items[original_matrix_key]
                            # Clean up user_chosen_base_colors_for_items if it exists for this key (though unlikely for non-multi-base)
                            if original_matrix_key in st.session_state.user_chosen_base_colors_for_items:
                                del st.session_state.user_chosen_base_colors_for_items[original_matrix_key]
                    
                    # Remove from the final_items_for_download list directly (this list is rebuilt anyway on rerun)
                    # The st.rerun() will handle rebuilding this list correctly.
                    st.toast(f"Removed: {combo['description']}", icon="🗑️")
                    st.rerun() # Rerun to update the review list and potentially the matrix checkboxes
            st.markdown("---")
        else: # No items in final_items_for_download
            st.info("Your list is empty. Please select products in Step 2 to continue.")


        # --- Step 4: Generate Master Data File ---
        st.header("Step 4: Download and add to your assortment")

        def prepare_excel_for_download_final():
            if not st.session_state.final_items_for_download: return None
            current_selected_currency_for_dl = st.session_state.selected_currency_session # Already selected in Step 1
            if not current_selected_currency_for_dl: return None # Should not happen if we reach here

            output_data = []
            ws_price_col_name_dynamic = f"Wholesale price ({current_selected_currency_for_dl})"
            rt_price_col_name_dynamic = f"Retail price ({current_selected_currency_for_dl})"

            # Determine final output columns, replacing generic price cols with dynamic ones
            final_cols = []
            seen_output_cols = set() # To handle potential duplicates if template has generic and specific
            for col_template in st.session_state.template_cols:
                col_template_lower = col_template.lower()
                if col_template_lower == "wholesale price": target_col = ws_price_col_name_dynamic
                elif col_template_lower == "retail price": target_col = rt_price_col_name_dynamic
                else: target_col = col_template
                
                if target_col not in seen_output_cols:
                    final_cols.append(target_col)
                    seen_output_cols.add(target_col)
            master_template_columns_final_output = final_cols


            for combo_selection in st.session_state.final_items_for_download:
                item_no_to_find = combo_selection['item_no']
                article_no_to_find = combo_selection['article_no']
                
                # Fetch data from the ORIGINAL (unfiltered by market) raw_df for the output file
                # This ensures all data fields are present, regardless of initial market filtering for display
                item_data_row_series_df = st.session_state.raw_df_original[st.session_state.raw_df_original['Item No'] == item_no_to_find]
                
                if not item_data_row_series_df.empty:
                    item_data_row_series = item_data_row_series_df.iloc[0]
                    output_row_dict = {}
                    for col_template_final_name in master_template_columns_final_output:
                        # If it's a dynamic price column, skip for now (will be populated from price matrix)
                        if col_template_final_name == ws_price_col_name_dynamic or col_template_final_name == rt_price_col_name_dynamic:
                            continue
                        # Otherwise, try to get data from the raw_df_original series
                        if col_template_final_name in item_data_row_series.index:
                            output_row_dict[col_template_final_name] = item_data_row_series[col_template_final_name]
                        else:
                            # If a column from template is not in raw_df, fill with None or placeholder
                            output_row_dict[col_template_final_name] = None 
                    
                    # Populate Wholesale Price
                    if st.session_state.wholesale_prices_df is not None and not st.session_state.wholesale_prices_df.empty:
                        # Ensure Article No is string for matching
                        ws_price_row_df = st.session_state.wholesale_prices_df[st.session_state.wholesale_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                        if not ws_price_row_df.empty and current_selected_currency_for_dl in ws_price_row_df.columns:
                            price_val = ws_price_row_df.iloc[0][current_selected_currency_for_dl]
                            output_row_dict[ws_price_col_name_dynamic] = price_val if pd.notna(price_val) else "N/A"
                        else:
                            output_row_dict[ws_price_col_name_dynamic] = "Price Not Found"
                    else:
                        output_row_dict[ws_price_col_name_dynamic] = "Wholesale Matrix Empty/Error"
                    
                    # Populate Retail Price
                    if st.session_state.retail_prices_df is not None and not st.session_state.retail_prices_df.empty:
                        rt_price_row_df = st.session_state.retail_prices_df[st.session_state.retail_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                        if not rt_price_row_df.empty and current_selected_currency_for_dl in rt_price_row_df.columns:
                            price_val = rt_price_row_df.iloc[0][current_selected_currency_for_dl]
                            output_row_dict[rt_price_col_name_dynamic] = price_val if pd.notna(price_val) else "N/A"
                        else:
                            output_row_dict[rt_price_col_name_dynamic] = "Price Not Found"
                    else:
                        output_row_dict[rt_price_col_name_dynamic] = "Retail Matrix Empty/Error"
                    
                    output_data.append(output_row_dict)

            if not output_data: return None

            output_df = pd.DataFrame(output_data, columns=master_template_columns_final_output) # Ensure column order
            output_excel_buffer = io.BytesIO()
            with pd.ExcelWriter(output_excel_buffer, engine='xlsxwriter') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
            return output_excel_buffer.getvalue()

        can_download_now = bool(st.session_state.final_items_for_download and st.session_state.selected_currency_session)

        if can_download_now:
            file_bytes = prepare_excel_for_download_final()
            if file_bytes:
                st.download_button(
                    label="Generate", # Changed from "Generate and Download Master Data File"
                    data=file_bytes,
                    file_name=f"masterdata_output_{st.session_state.selected_currency_session.replace(' ', '_').replace('.', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="final_download_action_button_v9", # Incremented key
                    help="Click to generate and download your customized master data file."
                )
        else:
            st.button("Generate", key="generate_file_disabled_button_v7", disabled=True, help="Please ensure a currency is chosen (Step 1) and items are selected (Step 2 & 3).")

    elif not st.session_state.selected_currency_session and st.session_state.raw_df_original is not None : # If currency not selected but data loaded
        st.info("Please select a currency in Step 1 to see available products and continue.")
    # If raw_df_original is None, the main error message at the end will cover it.

else: # files_loaded_successfully is False
    # Errors should have been displayed during the load_data phase.
    # This is a fallback generic message.
    st.error("One or more essential data files could not be loaded correctly, or required columns are missing. The application cannot continue. Please check the file paths, formats, and column names in your .xlsx files and ensure they are in the same directory as the script.")


# --- Styling (Original CSS from user's initial code) ---
st.markdown("""
<style>
    /* Apply background color to the main app container and body */
    .stApp, body {
        background-color: #EFEEEB !important;
    }
    .main .block-container {
        background-color: #EFEEEB !important;
        padding-top: 2rem; 
    }

    h1, h2, h3 { 
        text-transform: none !important; 
    }
    h1 { color: #333; } 
    h2 { 
        color: #1E40AF;
        padding-bottom: 5px;
        margin-top: 30px;
        margin-bottom: 15px;
    }
     h3 { 
        color: #1E40AF;
        font-size: 1.25em;
        padding-bottom: 3px;
        margin-top: 20px;
        margin-bottom: 10px;
    }

    /* Styling for the matrix-like headers */
    div[data-testid="stCaptionContainer"] > div > p { 
        font-weight: bold;
        font-size: 0.8em !important; 
        color: #31333F !important; 
        text-align: center;
        white-space: normal;
        overflow-wrap:break-word;
        line-height: 1.2; 
        padding: 2px;
    }
    .upholstery-header { 
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: clip !important;
        display: block;
        max-width: 100%;
        line-height: 1.2;
        color: #31333F !important; 
        text-transform: capitalize !important; 
        font-weight: bold !important;
        font-size: 0.8em !important;
    }
    div[data-testid="stCaptionContainer"] small { /* Color numbers */
        color: #31333F !important; 
        font-weight: normal !important; 
        font-size: 0.75em !important;
    }

    div[data-testid="stCaptionContainer"] img { 
        max-height: 25px !important;
        width: 25px !important;
        object-fit: cover !important;
        margin-right:2px;
    }
    .swatch-placeholder {
        width:25px !important;
        height:25px !important;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.6em;
        color: #ccc;
        border: 1px dashed #ddd;
        background-color: #f9f9f9;
    }
    .zoom-instruction {
        font-size: 0.6em;
        color: #555;
        text-align: left;
        padding-top: 10px;
    }

    /* --- Logo Styling --- */
    div[data-testid="stImage"],
    div[data-testid="stImage"] img {
        border-radius: 0 !important;
        overflow: visible !important;
    }

    /* --- Matrix Row and Cell Content Alignment --- */
    .product-name-cell {
        display: flex;
        align-items: center; 
        height: auto; 
        min-height: 30px; 
        line-height: 1.3; 
        max-height: calc(1.3em * 2 + 4px); 
        overflow-y: hidden; 
        color: #31333F !important; 
        font-weight: normal !important; 
        font-size: 0.8em !important; 
        padding-right: 5px; 
        word-break: break-word; 
        box-sizing: border-box;
    }

    /* Container for checkbox or unavailable cell content within each matrix data cell */
    div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] { 
        height: 30px !important; 
        min-height: 30px !important;
        display: flex !important;
        align-items: center !important; 
        justify-content: center !important; 
        padding: 0 !important; 
        margin: 0 !important;
        box-sizing: border-box;
    }
    /* This ensures the stMarkdownContainer (when used for unavailable cells, now removed) also behaves for centering */
    /* Keeping it in case grey boxes are re-introduced, but it won't affect empty cells */
    div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] > div[data-testid="stMarkdown"] > div[data-testid="stMarkdownContainer"] {
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 100%; 
        height: 100%; 
        box-sizing: border-box;
    }


    /* --- Checkbox Styling --- */
    div.stCheckbox { /* The main wrapper for st.checkbox widget */
         margin: 0 !important;
         padding: 0 !important; 
         display: flex !important;
         align-items: center !important;
         justify-content: center !important;
         width: 20px !important; 
         height: 20px !important; 
         box-sizing: border-box !important;
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] { /* The label that wraps the visual parts */
        width: 20px !important; 
        height: 20px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        padding: 0 !important;
        margin: 0 !important;
        box-sizing: border-box !important;
    }
    /* Visual box of the checkbox - UNCHECKED STATE */
    /* Targets the first span child of the label, which is the visual box */
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child { 
        background-color: #FFFFFF !important; 
        border: 1px solid #5B4A14 !important; 
        box-shadow: none !important;
        width: 20px !important; 
        height: 20px !important;
        border-radius: 0.25rem !important;
        margin: 0 !important;
        padding: 0 !important;
        box-sizing: border-box !important;
        display: flex !important; /* To center the SVG checkmark */
        align-items: center !important;
        justify-content: center !important;
    }
    /* Checkmark SVG - UNCHECKED STATE (effectively invisible) */
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child svg {
        fill: #FFFFFF !important; 
        width: 12px !important; /* Adjust size of SVG if needed */
        height: 12px !important;
    }

    /* Visual box of the checkbox - CHECKED STATE */
    /* Uses :has() to style the span when the input sibling is checked */
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"]:has(input[type="checkbox"][aria-checked="true"]) > span:first-child {
        background-color: #5B4A14 !important; 
        border-color: #5B4A14 !important; 
    }
    /* Checkmark SVG - CHECKED STATE (white) */
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"]:has(input[type="checkbox"][aria-checked="true"]) > span:first-child svg {
        fill: #FFFFFF !important; 
    }

    /* --- Unavailable Matrix Cell Styling - REMOVED as per request --- */


    hr {
        margin-top: 0.5rem !important;
        margin-bottom: 0.5rem !important;
        border-top: 1px solid #dee2e6;
    }
    section[data-testid="stSidebar"] hr {
        margin-top: 0.1rem !important;
        margin-bottom: 0.1rem !important;
    }


    /* --- Button Styling (General and Download Button) --- */
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"],
    div[data-testid="stButton"] button[data-testid^="stBaseButton"] {
        border: 1px solid #5B4A14 !important;
        background-color: #FFFFFF !important;
        color: #5B4A14 !important;
        padding: 0.375rem 0.75rem !important;
        font-size: 1rem !important;
        line-height: 1.5 !important;
        border-radius: 0.25rem !important;
        transition: color 0.15s ease-in-out, background-color 0.15s ease-in-out, border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out !important;
        font-weight: 500 !important;
        text-transform: none !important; 
    }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"] p,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"] p {
        color: inherit !important; 
        text-transform: none !important;
        margin: 0 !important; 
    }

    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:hover {
        background-color: #5B4A14 !important;
        color: #FFFFFF !important;
        border-color: #5B4A14 !important;
    }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover p,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:hover p {
        color: #FFFFFF !important; 
    }

    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:active,
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:focus,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:active,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:focus {
        background-color: #4A3D10 !important; 
        color: #FFFFFF !important;
        border-color: #4A3D10 !important;
        box-shadow: 0 0 0 0.2rem rgba(91, 74, 20, 0.4) !important;
        outline: none !important;
    }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:active p,
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:focus p,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:active p,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:focus p {
        color: #FFFFFF !important; 
    }


    small {
        font-size:0.9em;
        display:block;
        line-height:1.1;
    }
    /* --- Multiselect Tags Styling --- */
    /* This targets the selected tag itself within the stMultiSelect widget */
    /* Adding .st-ei and .st-eh to the selector for higher specificity against Streamlit's defaults */
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"].st-ei,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"].st-eh,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] /* Fallback */
     {
        background-color: transparent !important; 
        background-image: none !important; 
        border: 1px solid #000000 !important; /* Black border */
        border-radius: 0.25rem !important; 
        padding: 0.2em 0.4em !important; /* Adjusted padding */
        line-height: 1.2 !important; 
    }
    /* Text inside selected tag */
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[title] {
        color: #000000 !important; /* Black text */
        font-size: 0.85em !important;
        line-height: inherit !important; 
        margin-right: 4px !important; 
        vertical-align: middle !important; 
    }
    /* Close 'x' icon container span */
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[aria-hidden="true"] {
        display: inline-flex !important; 
        align-items: center !important;
    }
    /* Close 'x' icon SVG in selected tag */
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[aria-hidden="true"] svg {
        fill: #000000 !important; /* Black 'x' icon */
        width: 1em !important; 
        height: 1em !important;
        vertical-align: middle !important; 
    }


    /* White background and black text for input fields and dropdowns */
    div[data-testid="stTextInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:first-child,
    div[data-testid="stMultiSelect"] div[data-baseweb="input"],
    div[data-testid="stMultiSelect"] > div > div[data-baseweb="select"] > div:first-child {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #CCCCCC !important;
    }
    /* Text color for dropdown list items */
    div[data-baseweb="popover"] ul li {
        color: #000000 !important;
        background-color: #FFFFFF !important;
    }
    div[data-baseweb="popover"] ul li:hover {
        background-color: #f0f0f0 !important;
    }
     /* Text color for selected value in dropdown when not expanded */
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:first-child > div > div,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div:first-child > div > div {
         color: #000000 !important;
    }


    /* Active/focused border color for inputs and dropdowns */
    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stSelectbox"] div[data-baseweb="select"][aria-expanded="true"] > div:first-child,
    div[data-testid="stMultiSelect"] div[data-baseweb="input"]:focus-within,
    div[data-testid="stMultiSelect"] div[aria-expanded="true"] {
        border-color: #5B4A14 !important;
        box-shadow: 0 0 0 1px #5B4A14 !important;
    }
</style>
""", unsafe_allow_html=True)
