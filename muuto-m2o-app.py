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
PRICE_MATRIX_GBP_IE_XLSX_PATH = os.path.join(BASE_DIR, "price-matrix_GBP-IE.xlsx") 
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
    st.title("Muuto made-to-order master data tool") 

with top_col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)
    else:
        st.error(f"Muuto Logo not found. Expected at: {LOGO_PATH}.")


# --- App Introduction ---
st.markdown("""
Welcome to Muuto's Made-to-Order (MTO) Product Configurator!

This tool simplifies selecting MTO products and generating the data you need for your systems. Here's how it works:

* **Step 1: Select currency:**
    * Choose your preferred currency for pricing. This will determine which products are available.
* **Step 2: Select product family & combinations:**
    * Choose a product family to view its available products and upholstery options.
    * Select your desired product, upholstery, and color combinations directly in the matrix.
    * Use the "Select All" checkbox at the top of an upholstery color column to select/deselect all available products in that column.
    * **Step 2a: Specify base colors:** For items requiring base colors, they will be grouped by product family. For each family, you can select a specific base color to apply to all applicable products within that family, or choose base colors individually per product using the dropdowns.
* **Step 3: Review selections:**
    * Review the final list of configured products. You can remove items from this list if needed.
* **Step 4: Generate master data file:**
    * After making your selections, generate and download an Excel file containing all master data for your selected items.
""")

# --- Initialize session state variables ---
if 'raw_df' not in st.session_state: st.session_state.raw_df = None
if 'wholesale_prices_df' not in st.session_state: st.session_state.wholesale_prices_df = None
if 'retail_prices_df' not in st.session_state: st.session_state.retail_prices_df = None
if 'wholesale_prices_gbp_ie_df' not in st.session_state: st.session_state.wholesale_prices_gbp_ie_df = None
if 'retail_prices_gbp_ie_df' not in st.session_state: st.session_state.retail_prices_gbp_ie_df = None
if 'filtered_raw_df' not in st.session_state: st.session_state.filtered_raw_df = None
if 'template_cols' not in st.session_state: st.session_state.template_cols = None
if 'selected_family_session' not in st.session_state: st.session_state.selected_family_session = None
if 'matrix_selected_generic_items' not in st.session_state: st.session_state.matrix_selected_generic_items = {}
if 'user_chosen_base_colors_for_items' not in st.session_state: st.session_state.user_chosen_base_colors_for_items = {}
if 'final_items_for_download' not in st.session_state: st.session_state.final_items_for_download = []
if 'selected_currency_session' not in st.session_state: st.session_state.selected_currency_session = None


# --- Load Data Directly from XLSX files ---
files_loaded_successfully = True

if st.session_state.raw_df is None:
    if os.path.exists(RAW_DATA_XLSX_PATH):
        try:
            st.session_state.raw_df = pd.read_excel(RAW_DATA_XLSX_PATH, sheet_name=RAW_DATA_APP_SHEET)
            required_cols = ['Product Type', 'Product Model', 'Sofa Direction', 'Base Color', 'Product Family', 'Item No', 'Article No', 'Image URL swatch', 'Upholstery Type', 'Upholstery Color', 'Market']
            missing = [col for col in required_cols if col not in st.session_state.raw_df.columns]
            if missing:
                st.error(f"Required columns missing in '{os.path.basename(RAW_DATA_XLSX_PATH)}': {', '.join(missing)}.")
                files_loaded_successfully = False
            else:
                st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
                st.session_state.raw_df['Base Color Cleaned'] = st.session_state.raw_df['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)
                st.session_state.raw_df['Upholstery Type'] = st.session_state.raw_df['Upholstery Type'].astype(str).str.strip()
                st.session_state.raw_df['Market'] = st.session_state.raw_df['Market'].astype(str).str.upper()
        except Exception as e: st.error(f"Error loading Raw Data: {e}"); files_loaded_successfully = False
    else: st.error(f"Raw Data file not found: {RAW_DATA_XLSX_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.wholesale_prices_df is None:
    if os.path.exists(PRICE_MATRIX_EUROPE_XLSX_PATH):
        try:
            st.session_state.wholesale_prices_df = pd.read_excel(PRICE_MATRIX_EUROPE_XLSX_PATH, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
            st.session_state.retail_prices_df = pd.read_excel(PRICE_MATRIX_EUROPE_XLSX_PATH, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
        except Exception as e: st.error(f"Error loading EUROPE Prices: {e}"); files_loaded_successfully = False
    else: st.error(f"Price Matrix EUROPE file not found: {PRICE_MATRIX_EUROPE_XLSX_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.wholesale_prices_gbp_ie_df is None:
    if os.path.exists(PRICE_MATRIX_GBP_IE_XLSX_PATH):
        try:
            st.session_state.wholesale_prices_gbp_ie_df = pd.read_excel(PRICE_MATRIX_GBP_IE_XLSX_PATH, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
            st.session_state.retail_prices_gbp_ie_df = pd.read_excel(PRICE_MATRIX_GBP_IE_XLSX_PATH, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
        except Exception as e: st.error(f"Error loading GBP/IE Prices: {e}"); files_loaded_successfully = False
    else: st.error(f"Price Matrix GBP/IE file not found: {PRICE_MATRIX_GBP_IE_XLSX_PATH}"); files_loaded_successfully = False


if files_loaded_successfully and st.session_state.template_cols is None:
    if os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH):
        try:
            st.session_state.template_cols = pd.read_excel(MASTERDATA_TEMPLATE_XLSX_PATH).columns.tolist()
            if "Wholesale price" not in st.session_state.template_cols: st.session_state.template_cols.append("Wholesale price")
            if "Retail price" not in st.session_state.template_cols: st.session_state.template_cols.append("Retail price")
        except Exception as e: st.error(f"Error loading Template: {e}"); files_loaded_successfully = False
    else: st.error(f"Template file not found: {MASTERDATA_TEMPLATE_XLSX_PATH}"); files_loaded_successfully = False

# --- Main Application Area ---
if files_loaded_successfully:

    # --- Step 1: Select Currency ---
    st.header("Step 1: Select currency")
    
    europe_currencies = []
    gbp_ie_currencies = []
    EXPECTED_EUROPE_CURRENCIES = ['DACH - EURO', 'DKK', 'EURO', 'NOK', 'PLN', 'SEK', 'AUD']
    EXPECTED_GBP_IE_CURRENCIES = ['GBP', 'IE - EUR'] 

    try:
        if st.session_state.wholesale_prices_df is not None and not st.session_state.wholesale_prices_df.empty:
            article_no_col_name_ws_eu = st.session_state.wholesale_prices_df.columns[0]
            europe_currencies = [col for col in st.session_state.wholesale_prices_df.columns if col in EXPECTED_EUROPE_CURRENCIES and str(col).lower() != str(article_no_col_name_ws_eu).lower()]
        
        if st.session_state.wholesale_prices_gbp_ie_df is not None and not st.session_state.wholesale_prices_gbp_ie_df.empty:
            article_no_col_name_ws_gbp = st.session_state.wholesale_prices_gbp_ie_df.columns[0]
            gbp_ie_currencies = [col for col in st.session_state.wholesale_prices_gbp_ie_df.columns if col in EXPECTED_GBP_IE_CURRENCIES and str(col).lower() != str(article_no_col_name_ws_gbp).lower()]

        currency_options = [DEFAULT_NO_SELECTION] + sorted(list(set(europe_currencies + gbp_ie_currencies)))
        
        current_currency_idx = 0
        if st.session_state.selected_currency_session and st.session_state.selected_currency_session in currency_options:
            current_currency_idx = currency_options.index(st.session_state.selected_currency_session)
        
        prev_selected_currency = st.session_state.selected_currency_session
        selected_currency_choice = st.selectbox("Select Currency:", options=currency_options, index=current_currency_idx, key="currency_selector_main_key")

        if selected_currency_choice and selected_currency_choice != DEFAULT_NO_SELECTION:
            st.session_state.selected_currency_session = selected_currency_choice
        else:
            st.session_state.selected_currency_session = None

        if st.session_state.selected_currency_session != prev_selected_currency:
            st.session_state.matrix_selected_generic_items = {}
            st.session_state.user_chosen_base_colors_for_items = {}
            st.session_state.final_items_for_download = []
            st.session_state.selected_family_session = DEFAULT_NO_SELECTION
            if prev_selected_currency is not None : st.toast(f"Currency changed. Product selections reset.", icon="⚠️")


        if st.session_state.selected_currency_session and st.session_state.raw_df is not None:
            current_currency = st.session_state.selected_currency_session
            temp_df = st.session_state.raw_df.copy()
            if current_currency in EXPECTED_GBP_IE_CURRENCIES:
                st.session_state.filtered_raw_df = temp_df[temp_df['Market'] != 'EU']
            elif current_currency in EXPECTED_EUROPE_CURRENCIES:
                st.session_state.filtered_raw_df = temp_df[temp_df['Market'] != 'UK']
            else:
                 st.session_state.filtered_raw_df = pd.DataFrame(columns=st.session_state.raw_df.columns)
        elif st.session_state.raw_df is not None:
            st.session_state.filtered_raw_df = pd.DataFrame(columns=st.session_state.raw_df.columns)
        else: 
            st.session_state.filtered_raw_df = pd.DataFrame()
    except Exception as e:
        st.error(f"Error processing currency selection/filtering: {e}")
        st.session_state.selected_currency_session = None
        st.session_state.filtered_raw_df = pd.DataFrame()

    # --- Step 2: Select product combinations ---
    st.header("Step 2: Select product combinations (product / upholstery / color)")

    if not st.session_state.selected_currency_session:
        st.info("Please select a currency in Step 1 to see available products.")
    elif st.session_state.filtered_raw_df is None or st.session_state.filtered_raw_df.empty:
        st.info(f"No products available for {st.session_state.selected_currency_session} based on market rules.")
    else:
        df_for_display = st.session_state.filtered_raw_df

        available_families_in_view = [DEFAULT_NO_SELECTION] + sorted(df_for_display['Product Family'].dropna().unique()) if 'Product Family' in df_for_display.columns else [DEFAULT_NO_SELECTION]
        
        if st.session_state.selected_family_session not in available_families_in_view:
            st.session_state.selected_family_session = DEFAULT_NO_SELECTION

        selected_family_idx = 0
        if st.session_state.selected_family_session in available_families_in_view:
            selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)

        selected_family = st.selectbox("Select Product Family:", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
        st.session_state.selected_family_session = selected_family

        # --- Callback for individual checkbox toggle ---
        def handle_matrix_cb_toggle(prod_name, uph_type, uph_color, checkbox_key_matrix):
            is_checked = st.session_state[checkbox_key_matrix]
            current_selected_family_for_key = st.session_state.selected_family_session 
            generic_item_key = f"{current_selected_family_for_key}_{prod_name}_{uph_type}_{uph_color}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")

            if is_checked:
                matching_items = st.session_state.filtered_raw_df[
                    (st.session_state.filtered_raw_df['Product Family'] == current_selected_family_for_key) &
                    (st.session_state.filtered_raw_df['Product Display Name'] == prod_name) &
                    (st.session_state.filtered_raw_df['Upholstery Type'].fillna("N/A") == uph_type) &
                    (st.session_state.filtered_raw_df['Upholstery Color'].astype(str).fillna("N/A") == uph_color)
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
                        'resolved_base_if_single': unique_base_colors[0] if len(unique_base_colors) == 1 else (pd.NA if not unique_base_colors else None)
                    }
                    st.session_state.matrix_selected_generic_items[generic_item_key] = item_data
            else: 
                if generic_item_key in st.session_state.matrix_selected_generic_items:
                    del st.session_state.matrix_selected_generic_items[generic_item_key]
                    if generic_item_key in st.session_state.user_chosen_base_colors_for_items:
                        del st.session_state.user_chosen_base_colors_for_items[generic_item_key]

        # --- Callback for "Select All" column checkbox ---
        def handle_select_all_column_toggle(uph_type_col, uph_color_col, products_in_col, select_all_key):
            is_all_selected_for_column_now = st.session_state[select_all_key]
            current_selected_family_for_key = st.session_state.selected_family_session

            for prod_name in products_in_col:
                item_exists_df_col = st.session_state.filtered_raw_df[
                    (st.session_state.filtered_raw_df['Product Family'] == current_selected_family_for_key) &
                    (st.session_state.filtered_raw_df['Product Display Name'] == prod_name) &
                    (st.session_state.filtered_raw_df['Upholstery Type'].fillna("N/A") == uph_type_col) &
                    (st.session_state.filtered_raw_df['Upholstery Color'].astype(str).fillna("N/A") == uph_color_col)
                ]
                if not item_exists_df_col.empty:
                    generic_item_key_col = f"{current_selected_family_for_key}_{prod_name}_{uph_type_col}_{uph_color_col}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")
                    
                    if is_all_selected_for_column_now: 
                        if generic_item_key_col not in st.session_state.matrix_selected_generic_items:
                            unique_base_colors_col = item_exists_df_col['Base Color Cleaned'].dropna().unique().tolist()
                            first_item_match_col = item_exists_df_col.iloc[0]
                            item_data_col = {
                                'key': generic_item_key_col, 'family': current_selected_family_for_key, 'product': prod_name,
                                'upholstery_type': uph_type_col, 'upholstery_color': uph_color_col,
                                'requires_base_choice': len(unique_base_colors_col) > 1,
                                'available_bases': unique_base_colors_col if len(unique_base_colors_col) > 1 else [],
                                'item_no_if_single_base': first_item_match_col['Item No'] if len(unique_base_colors_col) <= 1 else None,
                                'article_no_if_single_base': first_item_match_col['Article No'] if len(unique_base_colors_col) <= 1 else None,
                                'resolved_base_if_single': unique_base_colors_col[0] if len(unique_base_colors_col) == 1 else (pd.NA if not unique_base_colors_col else None)
                            }
                            st.session_state.matrix_selected_generic_items[generic_item_key_col] = item_data_col
                    else: 
                        if generic_item_key_col in st.session_state.matrix_selected_generic_items:
                            del st.session_state.matrix_selected_generic_items[generic_item_key_col]
                            if generic_item_key_col in st.session_state.user_chosen_base_colors_for_items:
                                del st.session_state.user_chosen_base_colors_for_items[generic_item_key_col]
            
            action = "selected" if is_all_selected_for_column_now else "deselected"
            st.toast(f"All available items in column '{uph_type_col} - {uph_color_col}' {action}.", icon="✅" if is_all_selected_for_column_now else "❌")

        # --- Callback for individual item's base color multiselect ---
        def handle_base_color_multiselect_change(item_key_for_base_select):
            multiselect_widget_key = f"ms_base_{item_key_for_base_select}"
            st.session_state.user_chosen_base_colors_for_items[item_key_for_base_select] = st.session_state[multiselect_widget_key]

        # --- Callback for family-level "Select All [Base Color X] for this family" ---
        def handle_family_base_color_select_all_toggle(family_name_cb, base_color_cb, items_in_family_cb, checkbox_key_cb):
            is_checked = st.session_state[checkbox_key_cb]
            action_count = 0
            for item_data_cb in items_in_family_cb:
                item_key_cb = item_data_cb['key']
                # Ensure this item *can* have this base color
                if base_color_cb in item_data_cb['available_bases']:
                    current_bases_for_item = st.session_state.user_chosen_base_colors_for_items.get(item_key_cb, [])
                    if is_checked: # Add this base color
                        if base_color_cb not in current_bases_for_item:
                            st.session_state.user_chosen_base_colors_for_items[item_key_cb] = current_bases_for_item + [base_color_cb]
                            action_count += 1
                    else: # Remove this base color
                        if base_color_cb in current_bases_for_item:
                            new_bases = [b for b in current_bases_for_item if b != base_color_cb]
                            st.session_state.user_chosen_base_colors_for_items[item_key_cb] = new_bases
                            action_count += 1
            
            if action_count > 0:
                action_desc = "applied to" if is_checked else "removed from"
                st.toast(f"Base color '{base_color_cb}' {action_desc} {action_count} applicable product(s) in {family_name_cb}.", icon="✅" if is_checked else "❌")


        if selected_family and selected_family != DEFAULT_NO_SELECTION and 'Product Family' in df_for_display.columns:
            family_df = df_for_display[df_for_display['Product Family'] == selected_family]
            if not family_df.empty and 'Upholstery Type' in family_df.columns:
                products_in_family = sorted(family_df['Product Display Name'].dropna().unique())
                upholstery_types_in_family = sorted(family_df['Upholstery Type'].dropna().unique())

                if not products_in_family: st.info(f"No products in {selected_family} for current currency/market.")
                elif not upholstery_types_in_family: st.info(f"No upholstery types for {selected_family} for current currency/market.")
                else:
                    header_upholstery_types, header_swatches, header_color_numbers, data_column_map = [], [], [], []
                    header_upholstery_types.append("Product") 
                    header_swatches.append(" ") 
                    header_color_numbers.append(" ") 

                    for uph_type_clean in upholstery_types_in_family:
                        colors_for_type_df = family_df[family_df['Upholstery Type'] == uph_type_clean][['Upholstery Color', 'Image URL swatch']].drop_duplicates().sort_values(by='Upholstery Color')
                        if not colors_for_type_df.empty:
                            header_upholstery_types.extend([uph_type_clean] + [""] * (len(colors_for_type_df) -1) ) 
                            for _, color_row in colors_for_type_df.iterrows():
                                color_val, swatch_val = str(color_row['Upholstery Color']), color_row['Image URL swatch']
                                header_swatches.append(swatch_val if pd.notna(swatch_val) else None)
                                header_color_numbers.append(color_val)
                                data_column_map.append({'uph_type': uph_type_clean, 'uph_color': color_val, 'swatch': swatch_val})
                    
                    num_data_columns = len(data_column_map)
                    if num_data_columns > 0:
                        cols_uph_type_header = st.columns([2.5] + [1] * num_data_columns)
                        current_uph_type_header_display = None
                        for i, col_widget in enumerate(cols_uph_type_header):
                            if i > 0:
                                map_entry = data_column_map[i-1] 
                                if map_entry['uph_type'] != current_uph_type_header_display: 
                                    with col_widget: st.caption(f"<div class='upholstery-header'>{map_entry['uph_type']}</div>", unsafe_allow_html=True)
                                    current_uph_type_header_display = map_entry['uph_type']

                        cols_swatch_header = st.columns([2.5] + [1] * num_data_columns)
                        cols_swatch_header[0].markdown("<div class='zoom-instruction'><br>Click swatch to zoom</div>", unsafe_allow_html=True)
                        for i, col_widget in enumerate(cols_swatch_header[1:]): 
                            sw_url = data_column_map[i]['swatch'] 
                            with col_widget:
                                if sw_url and pd.notna(sw_url): st.image(sw_url, width=30)
                                else: st.markdown("<div class='swatch-placeholder'></div>", unsafe_allow_html=True)

                        cols_color_num_header = st.columns([2.5] + [1] * num_data_columns)
                        for i, col_widget in enumerate(cols_color_num_header):
                            if i > 0: 
                                with col_widget: st.caption(f"<small>{data_column_map[i-1]['uph_color']}</small>", unsafe_allow_html=True)
                        
                        # --- "Select All" Checkbox Row for Upholstery Columns ---
                        cols_select_all_header = st.columns([2.5] + [1] * num_data_columns, vertical_alignment="center") 
                        cols_select_all_header[0].markdown("<div class='select-all-label'>Select All:</div>", unsafe_allow_html=True) 
                        for i, col_widget_sa in enumerate(cols_select_all_header[1:]):
                            current_col_map_entry = data_column_map[i]
                            uph_type_for_col_sa = current_col_map_entry['uph_type']
                            uph_color_for_col_sa = current_col_map_entry['uph_color']
                            
                            all_in_col_selected = True
                            num_selectable_in_col = 0
                            for prod_name_sa in products_in_family:
                                item_exists_df_sa = family_df[
                                    (family_df['Product Display Name'] == prod_name_sa) &
                                    (family_df['Upholstery Type'] == uph_type_for_col_sa) &
                                    (family_df['Upholstery Color'].astype(str).fillna("N/A") == uph_color_for_col_sa)
                                ]
                                if not item_exists_df_sa.empty:
                                    num_selectable_in_col += 1
                                    generic_item_key_sa = f"{selected_family}_{prod_name_sa}_{uph_type_for_col_sa}_{uph_color_for_col_sa}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")
                                    if generic_item_key_sa not in st.session_state.matrix_selected_generic_items:
                                        all_in_col_selected = False; break
                            if num_selectable_in_col == 0 : all_in_col_selected = False

                            select_all_key = f"select_all_cb_{selected_family}_{uph_type_for_col_sa}_{uph_color_for_col_sa}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")
                            
                            with col_widget_sa:
                                if num_selectable_in_col > 0: 
                                    st.checkbox(" ", value=all_in_col_selected, key=select_all_key, 
                                                on_change=handle_select_all_column_toggle, 
                                                args=(uph_type_for_col_sa, uph_color_for_col_sa, products_in_family, select_all_key),
                                                label_visibility="collapsed",
                                                help=f"Select/Deselect all for {uph_type_for_col_sa} - {uph_color_for_col_sa}")
                                else: st.markdown("<div class='checkbox-placeholder'></div>", unsafe_allow_html=True)

                        st.markdown("---") 

                        for prod_name in products_in_family:
                            cols_product_row = st.columns([2.5] + [1] * num_data_columns, vertical_alignment="center")
                            cols_product_row[0].markdown(f"<div class='product-name-cell'>{prod_name}</div>", unsafe_allow_html=True)

                            for i, col_widget in enumerate(cols_product_row[1:]): 
                                current_col_uph_type_filter = data_column_map[i]['uph_type']
                                current_col_uph_color_filter = data_column_map[i]['uph_color']
                                item_exists_df = family_df[
                                    (family_df['Product Display Name'] == prod_name) &
                                    (family_df['Upholstery Type'] == current_col_uph_type_filter) &
                                    (family_df['Upholstery Color'].astype(str).fillna("N/A") == current_col_uph_color_filter)
                                ]
                                cell_container = col_widget.container() 
                                if not item_exists_df.empty:
                                    cb_key_str = f"cb_{selected_family}_{prod_name}_{current_col_uph_type_filter}_{current_col_uph_color_filter}".replace(" ","_").replace("/","_").replace("(","").replace(")","")
                                    generic_item_key_for_check = f"{selected_family}_{prod_name}_{current_col_uph_type_filter}_{current_col_uph_color_filter}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")
                                    is_gen_selected = generic_item_key_for_check in st.session_state.matrix_selected_generic_items
                                    cell_container.checkbox(" ", value=is_gen_selected, key=cb_key_str, 
                                                            on_change=handle_matrix_cb_toggle, 
                                                            args=(prod_name, current_col_uph_type_filter, current_col_uph_color_filter, cb_key_str), 
                                                            label_visibility="collapsed")
            else: 
                 if selected_family and selected_family != DEFAULT_NO_SELECTION : st.info(f"No data for {selected_family} with current currency/market.")
        else: 
            if selected_family and selected_family != DEFAULT_NO_SELECTION : st.info(f"Select product family.")


    # --- Step 2a: Specify Base Colors (Grouped by Family) ---
    items_needing_base_choice_now = [item_data for item_data in st.session_state.matrix_selected_generic_items.values() if item_data.get('requires_base_choice')]
    
    if items_needing_base_choice_now:
        st.subheader("Step 2a: Specify base colors")

        # Group items by product family
        items_by_family_for_base_step = {}
        for item_data in items_needing_base_choice_now:
            family_name = item_data['family']
            if family_name not in items_by_family_for_base_step:
                items_by_family_for_base_step[family_name] = []
            items_by_family_for_base_step[family_name].append(item_data)

        if not items_by_family_for_base_step:
            st.info("No selected items currently require base color specification.")
        else:
            for family_name_for_base, items_in_this_family for_base in items_by_family_for_base_step.items():
                st.markdown(f"#### {family_name_for_base}")

                # Determine all unique base colors available for *this family group*
                unique_bases_for_family_group = set()
                for item_in_fam in items_in_this_family for_base:
                    unique_bases_for_family_group.update(item_in_fam['available_bases'])
                
                sorted_unique_bases_for_family_group = sorted(list(unique_bases_for_family_group))

                if not sorted_unique_bases_for_family_group:
                    st.caption("No common base colors available or no items need base selection in this family.")
                else:
                    st.markdown("<small>Apply specific base color to all applicable products in this family:</small>", unsafe_allow_html=True)
                    # Create checkboxes for each unique base color in this family group
                    # Use 2 columns for these checkboxes for better layout if many base colors
                    num_base_cols = 2 if len(sorted_unique_bases_for_family_group) > 3 else 1 
                    base_color_cols = st.columns(num_base_cols)
                    col_idx = 0
                    for base_color_option in sorted_unique_bases_for_family_group:
                        with base_color_cols[col_idx % num_base_cols]:
                            family_base_cb_key = f"fam_base_all_{family_name_for_base}_{base_color_option}".replace(" ","_")
                            
                            # Determine if this base_color_option is selected for ALL applicable items in this family
                            is_this_base_selected_for_all_applicable_in_fam = True
                            num_applicable_for_this_base = 0
                            for item_in_fam_check in items_in_this_family for_base:
                                if base_color_option in item_in_fam_check['available_bases']:
                                    num_applicable_for_this_base +=1
                                    chosen_bases_for_item = st.session_state.user_chosen_base_colors_for_items.get(item_in_fam_check['key'], [])
                                    if base_color_option not in chosen_bases_for_item:
                                        is_this_base_selected_for_all_applicable_in_fam = False
                                        break
                            if num_applicable_for_this_base == 0: # If no items can even have this base color
                                is_this_base_selected_for_all_applicable_in_fam = False


                            if num_applicable_for_this_base > 0: # Only show checkbox if it applies to at least one item
                                st.checkbox(f"{base_color_option}", 
                                            value=is_this_base_selected_for_all_applicable_in_fam, 
                                            key=family_base_cb_key,
                                            on_change=handle_family_base_color_select_all_toggle,
                                            args=(family_name_for_base, base_color_option, items_in_this_family for_base, family_base_cb_key),
                                            help=f"Apply/Remove '{base_color_option}' for all applicable products in {family_name_for_base}.")
                        col_idx +=1
                    st.markdown("---") # Separator after family-level base selectors

                # List individual products within this family for base color selection
                for generic_item in items_in_this_family for_base:
                    item_key = generic_item['key']
                    multiselect_key = f"ms_base_{item_key}"
                    
                    st.markdown(f"**{generic_item['product']}** ({generic_item['upholstery_type']} - {generic_item['upholstery_color']})")
                    
                    st.multiselect(
                        label=f"Available base colors for this item:", 
                        options=generic_item['available_bases'],
                        default=st.session_state.user_chosen_base_colors_for_items.get(item_key, []),
                        key=multiselect_key,
                        on_change=handle_base_color_multiselect_change, # Simple callback
                        args=(item_key,)
                    )
                    st.markdown("---") # Separator between items
                st.markdown("---") # Separator between families
    
    # --- Step 3: Review Selections ---
    st.header("Step 3: Review selections")
    _current_final_items = [] 
    
    if st.session_state.filtered_raw_df is not None and not st.session_state.filtered_raw_df.empty:
        for key, gen_item_data in st.session_state.matrix_selected_generic_items.items():
            if not gen_item_data['requires_base_choice']: 
                if gen_item_data.get('item_no_if_single_base') is not None: 
                    _current_final_items.append({"description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']}" + (f" / Base: {gen_item_data['resolved_base_if_single']}" if pd.notna(gen_item_data['resolved_base_if_single']) else ""), "item_no": gen_item_data['item_no_if_single_base'], "article_no": gen_item_data['article_no_if_single_base'], "key_in_matrix": key})
            else: 
                selected_bases_for_this = st.session_state.user_chosen_base_colors_for_items.get(key, [])
                for bc in selected_bases_for_this:
                    specific_item_df = st.session_state.filtered_raw_df[
                        (st.session_state.filtered_raw_df['Product Family'] == gen_item_data['family']) &
                        (st.session_state.filtered_raw_df['Product Display Name'] == gen_item_data['product']) &
                        (st.session_state.filtered_raw_df['Upholstery Type'].fillna("N/A") == gen_item_data['upholstery_type']) &
                        (st.session_state.filtered_raw_df['Upholstery Color'].astype(str).fillna("N/A") == gen_item_data['upholstery_color']) &
                        (st.session_state.filtered_raw_df['Base Color Cleaned'].fillna("N/A") == bc)]
                    if not specific_item_df.empty:
                        actual_item = specific_item_df.iloc[0] 
                        _current_final_items.append({"description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']} / Base: {bc}", "item_no": actual_item['Item No'], "article_no": actual_item['Article No'], "key_in_matrix": key, "chosen_base": bc})
    
    temp_final_list_review, seen_item_keys_for_review = [], set() 
    for item_rev in _current_final_items:
        unique_config_key = f"{item_rev['item_no']}_{item_rev.get('chosen_base', 'NO_BASE_APPLICABLE')}"
        if unique_config_key not in seen_item_keys_for_review:
            temp_final_list_review.append(item_rev)
            seen_item_keys_for_review.add(unique_config_key)
    st.session_state.final_items_for_download = temp_final_list_review


    if st.session_state.final_items_for_download:
        st.markdown("**Current Selections for Download:**")
        for i in range(len(st.session_state.final_items_for_download) -1, -1, -1):
            combo = st.session_state.final_items_for_download[i]
            col1_rev, col2_rev = st.columns([0.9, 0.1])
            col1_rev.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")
            remove_button_key = f"final_review_remove_{i}_{combo['item_no']}_{combo.get('chosen_base','nobase')}"
            if col2_rev.button(f"Remove", key=remove_button_key):
                original_matrix_key = combo['key_in_matrix'] 
                if original_matrix_key in st.session_state.matrix_selected_generic_items:
                    generic_item_details = st.session_state.matrix_selected_generic_items[original_matrix_key]
                    if generic_item_details.get('requires_base_choice') and 'chosen_base' in combo:
                        chosen_base_to_remove = combo['chosen_base']
                        if original_matrix_key in st.session_state.user_chosen_base_colors_for_items:
                            if chosen_base_to_remove in st.session_state.user_chosen_base_colors_for_items[original_matrix_key]:
                                st.session_state.user_chosen_base_colors_for_items[original_matrix_key].remove(chosen_base_to_remove)
                                if not st.session_state.user_chosen_base_colors_for_items[original_matrix_key]:
                                    del st.session_state.user_chosen_base_colors_for_items[original_matrix_key] 
                                    # Only delete from matrix_selected_generic_items if NO bases are selected for an item that requires base choice
                                    # And if all its base choices were removed via this button.
                                    # This logic might need refinement if a generic item should persist even with no bases selected yet.
                                    # For now, if all chosen bases are removed, and it requires a choice, it's effectively deselected.
                                    if not st.session_state.user_chosen_base_colors_for_items.get(original_matrix_key): # Check if list is now empty or key gone
                                         del st.session_state.matrix_selected_generic_items[original_matrix_key]
                    else: # Item did not require base choice, or it's a single base item
                        del st.session_state.matrix_selected_generic_items[original_matrix_key]
                        # Clean up if it was mistakenly in user_chosen_base_colors_for_items
                        if original_matrix_key in st.session_state.user_chosen_base_colors_for_items:
                             del st.session_state.user_chosen_base_colors_for_items[original_matrix_key]
                st.session_state.final_items_for_download.pop(i)
                st.toast(f"Removed: {combo['description']}", icon="🗑️")
                st.rerun() 
        st.markdown("---")
    else:
        st.info("No items selected for download yet.")


    # --- Step 4: Generate Master Data File ---
    st.header("Step 4: Generate master data file")

    def prepare_excel_for_download_final():
        if not st.session_state.final_items_for_download: st.warning("No items selected."); return None
        current_selected_currency_for_dl = st.session_state.selected_currency_session
        if not current_selected_currency_for_dl: st.warning("Select currency first."); return None

        if current_selected_currency_for_dl in EXPECTED_GBP_IE_CURRENCIES:
            ws_prices, rt_prices = st.session_state.wholesale_prices_gbp_ie_df, st.session_state.retail_prices_gbp_ie_df
            if ws_prices is None or rt_prices is None: st.error(f"GBP/IE price matrix not loaded."); return None
        elif current_selected_currency_for_dl in EXPECTED_EUROPE_CURRENCIES:
            ws_prices, rt_prices = st.session_state.wholesale_prices_df, st.session_state.retail_prices_df
            if ws_prices is None or rt_prices is None: st.error(f"Europe price matrix not loaded."); return None
        else: st.error(f"Currency '{current_selected_currency_for_dl}' not configured."); return None
        
        output_data = []
        ws_price_col_dyn = f"Wholesale price ({current_selected_currency_for_dl})"
        rt_price_col_dyn = f"Retail price ({current_selected_currency_for_dl})"
        
        final_output_cols, seen_cols = [], set()
        for col_temp in st.session_state.template_cols:
            target_col = ws_price_col_dyn if col_temp.lower() == "wholesale price" else (rt_price_col_dyn if col_temp.lower() == "retail price" else col_temp)
            if target_col not in seen_cols: final_output_cols.append(target_col); seen_cols.add(target_col)
        if ws_price_col_dyn not in final_output_cols: final_output_cols.append(ws_price_col_dyn)
        if rt_price_col_dyn not in final_output_cols: final_output_cols.append(rt_price_col_dyn)
        
        if st.session_state.raw_df is None: st.error("Raw data unavailable."); return None

        for combo in st.session_state.final_items_for_download:
            item_no, article_no = combo['item_no'], combo['article_no']
            item_data_df = st.session_state.raw_df[st.session_state.raw_df['Item No'] == item_no]
            if not item_data_df.empty:
                item_series = item_data_df.iloc[0]
                row_dict = {col: item_series.get(col) for col in final_output_cols if col not in [ws_price_col_dyn, rt_price_col_dyn]}
                
                # Wholesale Price
                if not ws_prices.empty:
                    article_col_ws = ws_prices.columns[0]
                    price_df = ws_prices[ws_prices[article_col_ws].astype(str) == str(article_no)]
                    row_dict[ws_price_col_dyn] = price_df.iloc[0][current_selected_currency_for_dl] if not price_df.empty and current_selected_currency_for_dl in price_df.columns and pd.notna(price_df.iloc[0][current_selected_currency_for_dl]) else "Price Not Found"
                else: row_dict[ws_price_col_dyn] = "Wholesale Matrix Empty"
                
                # Retail Price
                if not rt_prices.empty:
                    article_col_rt = rt_prices.columns[0]
                    price_df = rt_prices[rt_prices[article_col_rt].astype(str) == str(article_no)]
                    row_dict[rt_price_col_dyn] = price_df.iloc[0][current_selected_currency_for_dl] if not price_df.empty and current_selected_currency_for_dl in price_df.columns and pd.notna(price_df.iloc[0][current_selected_currency_for_dl]) else "Price Not Found"
                else: row_dict[rt_price_col_dyn] = "Retail Matrix Empty"
                output_data.append(row_dict)
            else: st.warning(f"Item No {item_no} not found. Skipping.")

        if not output_data: st.info("No data to output."); return None
        output_df = pd.DataFrame(output_data, columns=final_output_cols)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
        return buffer.getvalue()

    can_download_now = bool(st.session_state.final_items_for_download and st.session_state.selected_currency_session)
    if can_download_now:
        file_bytes = prepare_excel_for_download_final()
        if file_bytes: 
            st.download_button(label="Generate and Download Master Data File", data=file_bytes, file_name=f"masterdata_output_{st.session_state.selected_currency_session.replace(' ', '_').replace('.', '')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="final_download_button_v10", help="Click to download.")
    else:
        help_msg = "Select currency (Step 1) and add items (Step 2 & 3)."
        if not st.session_state.selected_currency_session: help_msg = "Select currency first."
        elif not st.session_state.final_items_for_download: help_msg = "Select items first."
        st.button("Generate Master Data File", key="generate_disabled_button_v8", disabled=True, help=help_msg)

else: 
    st.error("Application cannot start. Critical data files missing or corrupt. Check paths and file integrity.")


# --- Styling ---
st.markdown("""
<style>
    /* Apply background color to the main app container and body */
    .stApp, body { background-color: #EFEEEB !important; }
    .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }
    h1, h2, h3 { text-transform: none !important; }
    h1 { color: #333; } 
    h2 { color: #1E40AF; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    h3 { color: #1E40AF; font-size: 1.25em; padding-bottom: 3px; margin-top: 20px; margin-bottom: 10px; }
    h4 { color: #102A63; font-size: 1.1em; margin-top: 15px; margin-bottom: 5px; } /* Styling for new h4 family headers */

    /* Styling for the matrix-like headers */
    div[data-testid="stCaptionContainer"] > div > p { font-weight: bold; font-size: 0.8em !important; color: #31333F !important; text-align: center; white-space: normal; overflow-wrap:break-word; line-height: 1.2; padding: 2px; }
    .upholstery-header { white-space: normal !important; overflow: visible !important; text-overflow: clip !important; display: block; max-width: 100%; line-height: 1.2; color: #31333F !important; text-transform: capitalize !important; font-weight: bold !important; font-size: 0.8em !important; }
    div[data-testid="stCaptionContainer"] small { color: #31333F !important; font-weight: normal !important; font-size: 0.75em !important; }
    div[data-testid="stCaptionContainer"] img { max-height: 25px !important; width: 25px !important; object-fit: cover !important; margin-right:2px; }
    .swatch-placeholder { width:25px !important; height:25px !important; display: flex; align-items: center; justify-content: center; font-size: 0.6em; color: #ccc; border: 1px dashed #ddd; background-color: #f9f9f9; }
    .zoom-instruction { font-size: 0.6em; color: #555; text-align: left; padding-top: 10px; }
    
    .select-all-label { 
        font-size: 0.75em; 
        color: #31333F; 
        text-align: right; 
        padding-right: 5px; 
        font-weight:bold;
        display: flex; /* Added for vertical alignment */
        align-items: center; /* Vertically align text with checkbox */
        height: 100%; /* Ensure it takes full cell height */
    }
    .checkbox-placeholder { width: 20px; height: 20px; margin: auto; }


    /* Logo Styling */
    div[data-testid="stImage"], div[data-testid="stImage"] img { border-radius: 0 !important; overflow: visible !important; }

    /* Matrix Row and Cell Content Alignment */
    .product-name-cell { display: flex; align-items: center; height: auto; min-height: 30px; line-height: 1.3; max-height: calc(1.3em * 2 + 4px); overflow-y: hidden; color: #31333F !important; font-weight: normal !important; font-size: 0.8em !important; padding-right: 5px; word-break: break-word; box-sizing: border-box; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] { height: 30px !important; min-height: 30px !important; display: flex !important; align-items: center !important; justify-content: center !important; padding: 0 !important; margin: 0 !important; box-sizing: border-box; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] > div[data-testid="stMarkdown"] > div[data-testid="stMarkdownContainer"] { display: flex !important; align-items: center !important; justify-content: center !important; width: 100%; height: 100%; box-sizing: border-box; }

    /* Checkbox Styling */
    div.stCheckbox { margin: 0 !important; padding: 0 !important; display: flex !important; align-items: center !important; justify-content: center !important; width: 20px !important; height: 20px !important; box-sizing: border-box !important; }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] { width: 20px !important; height: 20px !important; display: flex !important; align-items: center !important; justify-content: center !important; padding: 0 !important; margin: 0 !important; box-sizing: border-box !important; }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child { background-color: #FFFFFF !important; border: 1px solid #5B4A14 !important; box-shadow: none !important; width: 20px !important; height: 20px !important; border-radius: 0.25rem !important; margin: 0 !important; padding: 0 !important; box-sizing: border-box !important; display: flex !important; align-items: center !important; justify-content: center !important; }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child svg { fill: #FFFFFF !important; width: 12px !important; height: 12px !important; }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"]:has(input[type="checkbox"][aria-checked="true"]) > span:first-child { background-color: #5B4A14 !important; border-color: #5B4A14 !important; }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"]:has(input[type="checkbox"][aria-checked="true"]) > span:first-child svg { fill: #FFFFFF !important; }

    hr { margin-top: 0.5rem !important; margin-bottom: 0.5rem !important; border-top: 1px solid #dee2e6; }
    section[data-testid="stSidebar"] hr { margin-top: 0.1rem !important; margin-bottom: 0.1rem !important; }

    /* Button Styling */
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"], div[data-testid="stButton"] button[data-testid^="stBaseButton"] { border: 1px solid #5B4A14 !important; background-color: #FFFFFF !important; color: #5B4A14 !important; padding: 0.375rem 0.75rem !important; font-size: 1rem !important; line-height: 1.5 !important; border-radius: 0.25rem !important; transition: color 0.15s ease-in-out, background-color 0.15s ease-in-out, border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out !important; font-weight: 500 !important; text-transform: none !important; }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"] p, div[data-testid="stButton"] button[data-testid^="stBaseButton"] p { color: inherit !important; text-transform: none !important; margin: 0 !important; }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover, div[data-testid="stButton"] button[data-testid^="stBaseButton"]:hover { background-color: #5B4A14 !important; color: #FFFFFF !important; border-color: #5B4A14 !important; }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover p, div[data-testid="stButton"] button[data-testid^="stBaseButton"]:hover p { color: #FFFFFF !important; }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:active, div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:focus, div[data-testid="stButton"] button[data-testid^="stBaseButton"]:active, div[data-testid="stButton"] button[data-testid^="stBaseButton"]:focus { background-color: #4A3D10 !important; color: #FFFFFF !important; border-color: #4A3D10 !important; box-shadow: 0 0 0 0.2rem rgba(91, 74, 20, 0.4) !important; outline: none !important; }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:active p, div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:focus p, div[data-testid="stButton"] button[data-testid^="stBaseButton"]:active p, div[data-testid="stButton"] button[data-testid^="stBaseButton"]:focus p { color: #FFFFFF !important; }

    small { font-size:0.9em; display:block; line-height:1.1; }
    /* Multiselect Tags Styling */
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"].st-ei, div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"].st-eh, div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] { background-color: transparent !important; background-image: none !important; border: 1px solid #000000 !important; border-radius: 0.25rem !important; padding: 0.2em 0.4em !important; line-height: 1.2 !important; }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[title] { color: #000000 !important; font-size: 0.85em !important; line-height: inherit !important; margin-right: 4px !important; vertical-align: middle !important; }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[aria-hidden="true"] { display: inline-flex !important; align-items: center !important; }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[aria-hidden="true"] svg { fill: #000000 !important; width: 1em !important; height: 1em !important; vertical-align: middle !important; }

    /* Input fields and dropdowns styling */
    div[data-testid="stTextInput"] input, div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:first-child, div[data-testid="stMultiSelect"] div[data-baseweb="input"], div[data-testid="stMultiSelect"] > div > div[data-baseweb="select"] > div:first-child { background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #CCCCCC !important; }
    div[data-baseweb="popover"] ul li { color: #000000 !important; background-color: #FFFFFF !important; }
    div[data-baseweb="popover"] ul li:hover { background-color: #f0f0f0 !important; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:first-child > div > div, div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div:first-child > div > div { color: #000000 !important; }
    div[data-testid="stTextInput"] input:focus, div[data-testid="stSelectbox"] div[data-baseweb="select"][aria-expanded="true"] > div:first-child, div[data-testid="stMultiSelect"] div[data-baseweb="input"]:focus-within, div[data-testid="stMultiSelect"] div[aria-expanded="true"] { border-color: #5B4A14 !important; box-shadow: 0 0 0 1px #5B4A14 !important; }

    /* Styling for ALL Info/Warning/Alert Boxes */
    div[data-testid="stAlert"] { background-color: #f0f2f6 !important; border: 1px solid #D1D5DB !important; border-radius: 0.25rem !important; }
    div[data-testid="stAlert"] > div:first-child { background-color: transparent !important; }
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"], div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"] p { color: #31333F !important; }
    div[data-testid="stAlert"] svg { fill: #4B5563 !important; }
</style>
""", unsafe_allow_html=True)
