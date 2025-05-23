import streamlit as st
import pandas as pd
import io
import os
import re # For regex in cleaning

# --- Page Configuration (MUST BE THE FIRST STREAMLIT COMMAND) ---
st.set_page_config(
    layout="wide",
    page_title="Muuto M2O",
    page_icon="favicon.png"
)

# --- Configuration & Constants ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_XLSX_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
PRICE_MATRIX_EUROPE_XLSX_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx")
PRICE_MATRIX_UK_XLSX_PATH = os.path.join(BASE_DIR, "price-matrix_UK-EI.xlsx")
MASTERDATA_TEMPLATE_XLSX_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")

RAW_DATA_APP_SHEET = "APP"
PRICE_MATRIX_WHOLESALE_SHEET = "Price matrix wholesale"
PRICE_MATRIX_RETAIL_SHEET = "Price matrix retail"
DEFAULT_NO_SELECTION = "--- Please Select ---"
# Fallback key name if no price files are successfully loaded to determine it.
# User specified "Article No." for price files and "Article No" for raw-data.
# The code will try to use the actual first column header from price files.
# For raw-data, it will specifically look for "Article No".
PRICE_FILE_ARTICLE_NO_KEY_DEFAULT = "Article No." 
RAW_DATA_ARTICLE_NO_COL = "Article No"
RAW_DATA_ITEM_NO_COL = "Item No"
RAW_DATA_MARKET_COL = "Market"


# Define currency groups as per user specification
CURRENCIES_GROUP_1_EU = ['DACH - EURO', 'DKK', 'EURO', 'NOK', 'PLN', 'SEK', 'AUD']
CURRENCIES_GROUP_2_UK = ['GBP', 'IE - EUR']


# --- Helper Function to Construct Product Display Name ---
def construct_product_display_name(row):
    name_parts = []
    product_type, product_model, sofa_direction = row.get('Product Type'), row.get('Product Model'), row.get('Sofa Direction')
    if pd.notna(product_type) and str(product_type).strip().upper() != "N/A": name_parts.append(str(product_type))
    if pd.notna(product_model) and str(product_model).strip().upper() != "N/A": name_parts.append(str(product_model))
    if str(product_type).strip().lower() == "sofa chaise longue" and pd.notna(sofa_direction) and str(sofa_direction).strip().upper() != "N/A":
        name_parts.append(str(sofa_direction))
    return " - ".join(name_parts) if name_parts else "Unnamed Product"

# --- Helper Function to Clean Key Columns (Article No, Item No) ---
def clean_key_series(series):
    if series is None: return None
    s_cleaned = series.astype(str).str.strip().str.upper()
    s_cleaned = s_cleaned.str.replace(r'\.0$', '', regex=True) # "12345.0" -> "12345"
    s_cleaned = s_cleaned.replace(['', 'NAN', '<NA>', 'NONE'], None) # Replace common NA strings with None
    return s_cleaned

# --- Main App Logic ---
top_col1, _, top_col2 = st.columns([5.5, 0.5, 1])
with top_col1: st.title("Welcome to the Muuto M2O master data generator")
with top_col2:
    if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=120)
    else: st.error(f"Logo not found: {LOGO_PATH}.")

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

# --- Initialize session state ---
for key, default_val in [
    ('raw_df_original', None), ('raw_df', None), 
    ('europe_wholesale_df', None), ('europe_retail_df', None),
    ('uk_wholesale_df', None), ('uk_retail_df', None),
    ('template_cols', None), ('selected_family_session', None),
    ('matrix_selected_generic_items', {}), ('user_chosen_base_colors_for_items', {}),
    ('final_items_for_download', []), ('selected_currency_session', None),
    ('price_file_article_no_key', PRICE_FILE_ARTICLE_NO_KEY_DEFAULT) # Stores the determined key for price files
]:
    if key not in st.session_state: st.session_state[key] = default_val

@st.cache_data
def load_data():
    raw_df_original_data, europe_ws, europe_rt, uk_ws, uk_rt, template_cols_data = None, None, None, None, None, None
    data_load_errors_list = []
    # Determined from the first column of the first successfully loaded price file.
    # User specified "Article No." for price files.
    price_file_key_col_name = PRICE_FILE_ARTICLE_NO_KEY_DEFAULT 
    key_name_is_set = False

    # Load Raw Data
    if os.path.exists(RAW_DATA_XLSX_PATH):
        try:
            raw_df_original_data = pd.read_excel(RAW_DATA_XLSX_PATH, sheet_name=RAW_DATA_APP_SHEET)
            required_cols = ['Product Type', 'Product Model', RAW_DATA_MARKET_COL, RAW_DATA_ITEM_NO_COL, RAW_DATA_ARTICLE_NO_COL] # Check for essential raw data cols
            missing = [col for col in required_cols if col not in raw_df_original_data.columns]
            if missing: data_load_errors_list.append(f"Raw data missing columns: {missing}.")
            else:
                raw_df_original_data['Product Display Name'] = raw_df_original_data.apply(construct_product_display_name, axis=1)
                raw_df_original_data['Base Color Cleaned'] = raw_df_original_data['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)
                raw_df_original_data['Upholstery Type'] = raw_df_original_data['Upholstery Type'].astype(str).str.strip()
                # Clean Item No and Article No from raw_data
                raw_df_original_data[RAW_DATA_ARTICLE_NO_COL] = clean_key_series(raw_df_original_data[RAW_DATA_ARTICLE_NO_COL])
                raw_df_original_data[RAW_DATA_ITEM_NO_COL] = clean_key_series(raw_df_original_data[RAW_DATA_ITEM_NO_COL])
        except Exception as e: data_load_errors_list.append(f"Error loading Raw Data: {e}")
    else: data_load_errors_list.append(f"Raw Data file not found: {RAW_DATA_XLSX_PATH}")

    # Helper to load, clean, and process a single price sheet
    def _load_and_process_price_sheet(file_path, sheet_name, key_name_ref_list, key_is_set_ref_list):
        if not os.path.exists(file_path): return None
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if df.empty: return None

            current_first_col_name_stripped = df.columns[0].strip()
            
            if not key_is_set_ref_list[0]: # If this is the first price sheet defining the key
                key_name_ref_list[0] = current_first_col_name_stripped
                key_is_set_ref_list[0] = True
            
            # Rename current df's first col to the consistently determined & stripped key name
            if df.columns[0] != key_name_ref_list[0]:
                 df.rename(columns={df.columns[0]: key_name_ref_list[0]}, inplace=True)
            
            df[key_name_ref_list[0]] = clean_key_series(df[key_name_ref_list[0]])
            df.drop_duplicates(subset=[key_name_ref_list[0]], keep='first', inplace=True)
            df.columns = [col.strip() for col in df.columns] # Strip all column names
            return df
        except Exception as e:
            data_load_errors_list.append(f"Error loading {sheet_name} from {os.path.basename(file_path)}: {e}")
            return None

    key_name_ref = [price_file_key_col_name] # Pass by reference
    key_name_was_set_ref = [key_name_is_set]

    europe_ws = _load_and_process_price_sheet(PRICE_MATRIX_EUROPE_XLSX_PATH, PRICE_MATRIX_WHOLESALE_SHEET, key_name_ref, key_name_was_set_ref)
    europe_rt = _load_and_process_price_sheet(PRICE_MATRIX_EUROPE_XLSX_PATH, PRICE_MATRIX_RETAIL_SHEET, key_name_ref, key_name_was_set_ref)
    
    uk_ws = _load_and_process_price_sheet(PRICE_MATRIX_UK_XLSX_PATH, PRICE_MATRIX_WHOLESALE_SHEET, key_name_ref, key_name_was_set_ref)
    uk_rt = _load_and_process_price_sheet(PRICE_MATRIX_UK_XLSX_PATH, PRICE_MATRIX_RETAIL_SHEET, key_name_ref, key_name_was_set_ref)

    price_file_key_col_name = key_name_ref[0] # Final determined and stripped key name for price files

    if not europe_ws and not europe_rt and not uk_ws and not uk_rt and not data_load_errors_list :
        data_load_errors_list.append("No price data could be loaded from any price matrix file.")


    # Load Template
    if os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH):
        try:
            template_cols_data = pd.read_excel(MASTERDATA_TEMPLATE_XLSX_PATH).columns.tolist()
            if "Wholesale price" not in template_cols_data: template_cols_data.append("Wholesale price")
            if "Retail price" not in template_cols_data: template_cols_data.append("Retail price")
        except Exception as e: data_load_errors_list.append(f"Error loading Template: {e}")
    else: data_load_errors_list.append(f"Template file not found: {MASTERDATA_TEMPLATE_XLSX_PATH}")

    return raw_df_original_data, europe_ws, europe_rt, uk_ws, uk_rt, template_cols_data, data_load_errors_list, price_file_key_col_name

# --- Load data and set session state ---
raw_df_orig, eu_ws, eu_rt, uk_ws, uk_rt, tpl_cols, errors, price_key = load_data()
files_loaded_successfully = not errors 

if files_loaded_successfully:
    st.session_state.raw_df_original = raw_df_orig
    st.session_state.europe_wholesale_df = eu_ws
    st.session_state.europe_retail_df = eu_rt
    st.session_state.uk_wholesale_df = uk_ws
    st.session_state.uk_retail_df = uk_rt
    st.session_state.template_cols = tpl_cols
    st.session_state.price_file_article_no_key = price_key
else:
    for error in errors: st.error(error)

# --- Main Application Area ---
if files_loaded_successfully:
    st.header("Step 1: Select your currency")
    def on_currency_change():
        st.session_state.selected_family_session = DEFAULT_NO_SELECTION
        st.session_state.matrix_selected_generic_items.clear()
        st.session_state.user_chosen_base_colors_for_items.clear()
        st.session_state.final_items_for_download.clear()

    # Define combined currency list for the dropdown based on user spec
    # These are the display names in the dropdown. The actual column lookup will use these.
    combined_currency_options = [DEFAULT_NO_SELECTION] + CURRENCIES_GROUP_1_EU + CURRENCIES_GROUP_2_UK
    
    current_idx = combined_currency_options.index(st.session_state.selected_currency_session) if st.session_state.selected_currency_session in combined_currency_options else 0
    selected_choice = st.selectbox("Select Currency:", combined_currency_options, index=current_idx, key="currency_selector", on_change=on_currency_change)
    st.session_state.selected_currency_session = selected_choice if selected_choice != DEFAULT_NO_SELECTION else None

    # Filter raw_df for display based on currency and market logic
    if st.session_state.selected_currency_session and st.session_state.raw_df_original is not None:
        selected_curr = st.session_state.selected_currency_session
        if selected_curr in CURRENCIES_GROUP_1_EU:
            st.session_state.raw_df = st.session_state.raw_df_original[st.session_state.raw_df_original[RAW_DATA_MARKET_COL].astype(str).str.upper() == 'EU'].copy()
        elif selected_curr in CURRENCIES_GROUP_2_UK:
            st.session_state.raw_df = st.session_state.raw_df_original[st.session_state.raw_df_original[RAW_DATA_MARKET_COL].astype(str).str.upper() == 'UK'].copy()
        else: # Should not happen if dropdown is restricted to the two groups
            st.session_state.raw_df = None 
            st.warning("Selected currency does not map to a defined market group.")
    else: st.session_state.raw_df = None
    
    # --- Steps 2, 3, 4 (UI and core logic largely as before, ensure it uses cleaned data) ---
    # (The rest of the Streamlit UI code for steps 2, 3, 4 follows, with adjustments to price lookup in Step 4)
    if st.session_state.selected_currency_session and st.session_state.raw_df is not None and not st.session_state.raw_df.empty:
        st.markdown("---"); st.header("Step 2: Explore your options and choose your sofacombinations")
        df_for_display = st.session_state.raw_df
        available_families_in_view = [DEFAULT_NO_SELECTION] + sorted(df_for_display['Product Family'].dropna().unique()) if 'Product Family' in df_for_display.columns else [DEFAULT_NO_SELECTION]
        if st.session_state.selected_family_session not in available_families_in_view: st.session_state.selected_family_session = DEFAULT_NO_SELECTION
        selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)
        selected_family = st.selectbox("Select Product Family:", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
        st.session_state.selected_family_session = selected_family

        def handle_matrix_cb_toggle(prod_name, uph_type, uph_color, checkbox_key_matrix):
            is_checked = st.session_state[checkbox_key_matrix]
            current_selected_family_for_key = st.session_state.selected_family_session 
            generic_item_key = f"{current_selected_family_for_key}_{prod_name}_{uph_type}_{uph_color}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")
            if is_checked:
                matching_items = st.session_state.raw_df[ (st.session_state.raw_df['Product Family'] == current_selected_family_for_key) & (st.session_state.raw_df['Product Display Name'] == prod_name) & (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == uph_type) & (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == uph_color) ]
                if not matching_items.empty:
                    unique_base_colors = matching_items['Base Color Cleaned'].dropna().unique().tolist()
                    first_item_match = matching_items.iloc[0]
                    item_data = { 'key': generic_item_key, 'family': current_selected_family_for_key, 'product': prod_name, 'upholstery_type': uph_type, 'upholstery_color': uph_color, 'requires_base_choice': len(unique_base_colors) > 1, 'available_bases': unique_base_colors if len(unique_base_colors) > 1 else [], 'item_no_if_single_base': first_item_match[RAW_DATA_ITEM_NO_COL], 'article_no_if_single_base': first_item_match[RAW_DATA_ARTICLE_NO_COL], 'resolved_base_if_single': unique_base_colors[0] if len(unique_base_colors) == 1 else (pd.NA if not unique_base_colors and len(unique_base_colors) == 0 else None) }
                    st.session_state.matrix_selected_generic_items[generic_item_key] = item_data
                    st.toast(f"Selected: {prod_name} / {uph_type} / {uph_color}", icon="➕")
            else:
                if generic_item_key in st.session_state.matrix_selected_generic_items:
                    del st.session_state.matrix_selected_generic_items[generic_item_key]
                    if generic_item_key in st.session_state.user_chosen_base_colors_for_items: del st.session_state.user_chosen_base_colors_for_items[generic_item_key]
                    st.toast(f"Deselected: {prod_name} / {uph_type} / {uph_color}", icon="➖")
        def handle_base_color_multiselect_change(item_key_for_base_select): st.session_state.user_chosen_base_colors_for_items[item_key_for_base_select] = st.session_state[f"ms_base_{item_key_for_base_select}"]

        if selected_family and selected_family != DEFAULT_NO_SELECTION and 'Product Family' in df_for_display.columns:
            family_df = df_for_display[df_for_display['Product Family'] == selected_family]
            if not family_df.empty and 'Upholstery Type' in family_df.columns:
                products_in_family, upholstery_types_in_family = sorted(family_df['Product Display Name'].dropna().unique()), sorted(family_df['Upholstery Type'].dropna().unique())
                if not products_in_family: st.info(f"No products in: {selected_family} for current market.")
                elif not upholstery_types_in_family: st.info(f"No upholstery types for: {selected_family} for current market.")
                else: # Matrix rendering ...
                    header_upholstery_types, header_swatches, header_color_numbers, data_column_map = ["Product"], [" "], [" "], []
                    for uph_type_clean in upholstery_types_in_family:
                        colors_for_type_df = family_df[family_df['Upholstery Type'] == uph_type_clean][['Upholstery Color', 'Image URL swatch']].drop_duplicates().sort_values(by='Upholstery Color')
                        if not colors_for_type_df.empty:
                            header_upholstery_types.extend([uph_type_clean] + [""] * (len(colors_for_type_df) -1) )
                            for _, color_row in colors_for_type_df.iterrows():
                                header_swatches.append(color_row['Image URL swatch'] if pd.notna(color_row['Image URL swatch']) else None)
                                header_color_numbers.append(str(color_row['Upholstery Color']))
                                data_column_map.append({'uph_type': uph_type_clean, 'uph_color': str(color_row['Upholstery Color']), 'swatch': color_row['Image URL swatch']})
                    num_data_columns = len(data_column_map)
                    if num_data_columns == 0: st.info(f"No upholstery/color combinations for: {selected_family}")
                    else: 
                        cols_uph_type_header = st.columns([2.5] + [1] * num_data_columns)
                        current_uph_type_header_display = None
                        for i, col_widget in enumerate(cols_uph_type_header):
                            if i == 0: col_widget.caption("")
                            else:
                                map_entry = data_column_map[i-1]
                                if map_entry['uph_type'] != current_uph_type_header_display: col_widget.caption(f"<div class='upholstery-header'>{map_entry['uph_type']}</div>", unsafe_allow_html=True); current_uph_type_header_display = map_entry['uph_type']
                        cols_swatch_header = st.columns([2.5] + [1] * num_data_columns)
                        for i, col_widget in enumerate(cols_swatch_header):
                            if i == 0: col_widget.markdown("<div class='zoom-instruction'><br>Click swatch to zoom</div>", unsafe_allow_html=True)
                            else:
                                sw_url = data_column_map[i-1]['swatch']
                                if sw_url and pd.notna(sw_url): col_widget.image(sw_url, width=30)
                                else: col_widget.markdown("<div class='swatch-placeholder'></div>", unsafe_allow_html=True)
                        cols_color_num_header = st.columns([2.5] + [1] * num_data_columns)
                        for i, col_widget in enumerate(cols_color_num_header):
                            if i == 0: col_widget.caption("")
                            else: col_widget.caption(f"<small>{data_column_map[i-1]['uph_color']}</small>", unsafe_allow_html=True)
                        st.markdown("---")
                        for prod_name in products_in_family:
                            cols_product_row = st.columns([2.5] + [1] * num_data_columns, vertical_alignment="center")
                            cols_product_row[0].markdown(f"<div class='product-name-cell'>{prod_name}</div>", unsafe_allow_html=True)
                            for i, col_widget in enumerate(cols_product_row[1:]):
                                current_col_uph_type, current_col_uph_color = data_column_map[i]['uph_type'], data_column_map[i]['uph_color']
                                item_exists_df = family_df[(family_df['Product Display Name'] == prod_name) & (family_df['Upholstery Type'] == current_col_uph_type) & (family_df['Upholstery Color'].astype(str).fillna("N/A") == current_col_uph_color)]
                                cell_container = col_widget.container()
                                if not item_exists_df.empty:
                                    cb_key_str = f"cb_{selected_family}_{prod_name}_{current_col_uph_type}_{current_col_uph_color}".replace(" ","_").replace("/","_").replace("(","").replace(")","")
                                    generic_item_key_for_check = f"{selected_family}_{prod_name}_{current_col_uph_type}_{current_col_uph_color}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")
                                    is_gen_selected = generic_item_key_for_check in st.session_state.matrix_selected_generic_items
                                    cell_container.checkbox(" ", value=is_gen_selected, key=cb_key_str, on_change=handle_matrix_cb_toggle, args=(prod_name, current_col_uph_type, current_col_uph_color, cb_key_str), label_visibility="collapsed")
            else: 
                if selected_family and selected_family != DEFAULT_NO_SELECTION : st.info(f"No data for: {selected_family} with current selections.")
        
        items_needing_base_choice_now = [item_data for item_data in st.session_state.matrix_selected_generic_items.values() if item_data.get('requires_base_choice')]
        if items_needing_base_choice_now: # Base color selection
            st.subheader("Specify base colors for selected items") 
            for generic_item in items_needing_base_choice_now:
                item_key, multiselect_key = generic_item['key'], f"ms_base_{generic_item['key']}"
                st.markdown(f"**{generic_item['product']}** ({generic_item['upholstery_type']} - {generic_item['upholstery_color']})")
                current_selection = st.session_state.user_chosen_base_colors_for_items.get(item_key, [])
                valid_bases = [base for base in generic_item['available_bases'] if pd.notna(base)]
                st.multiselect("Available base colors:", options=valid_bases, default=current_selection, key=multiselect_key, on_change=handle_base_color_multiselect_change, args=(item_key,))
                st.markdown("---")

        st.header("Step 3: Review your list") # Review selections
        _current_final_items = []
        for key, gen_item_data in st.session_state.matrix_selected_generic_items.items():
            if not gen_item_data['requires_base_choice']:
                if gen_item_data.get('item_no_if_single_base') is not None:
                    desc_base = f" / Base: {gen_item_data['resolved_base_if_single']}" if pd.notna(gen_item_data['resolved_base_if_single']) and str(gen_item_data['resolved_base_if_single']).strip().upper() != "N/A" else ""
                    _current_final_items.append({"description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']}{desc_base}", "item_no": gen_item_data['item_no_if_single_base'], "article_no": gen_item_data['article_no_if_single_base'], "key_in_matrix": key})
            else:
                for bc in st.session_state.user_chosen_base_colors_for_items.get(key, []):
                    specific_item_df = st.session_state.raw_df[(st.session_state.raw_df['Product Family'] == gen_item_data['family']) & (st.session_state.raw_df['Product Display Name'] == gen_item_data['product']) & (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == gen_item_data['upholstery_type']) & (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == gen_item_data['upholstery_color']) & (st.session_state.raw_df['Base Color Cleaned'].fillna("N/A") == bc)]
                    if not specific_item_df.empty:
                        actual_item = specific_item_df.iloc[0]
                        _current_final_items.append({"description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']} / Base: {bc}", "item_no": actual_item[RAW_DATA_ITEM_NO_COL], "article_no": actual_item[RAW_DATA_ARTICLE_NO_COL], "key_in_matrix": key, "chosen_base": bc})
        temp_final_list_review, seen_item_keys_review = [], set()
        for item_rev in _current_final_items:
            unique_final_item_key = f"{item_rev['item_no']}_{item_rev.get('chosen_base', 'NO_BASE')}"
            if unique_final_item_key not in seen_item_keys_review: temp_final_list_review.append(item_rev); seen_item_keys_review.add(unique_final_item_key)
        st.session_state.final_items_for_download = temp_final_list_review
        if st.session_state.final_items_for_download: # Review list display and remove
            st.markdown("**Current Selections for Download:**")
            for i, combo in enumerate(st.session_state.final_items_for_download):
                col1_rev, col2_rev = st.columns([0.9, 0.1])
                col1_rev.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")
                if col2_rev.button(f"Remove", key=f"final_review_remove_{i}_{combo['item_no']}_{combo.get('chosen_base','nobase')}"):
                    original_matrix_key = combo['key_in_matrix']
                    if original_matrix_key in st.session_state.matrix_selected_generic_items:
                        if st.session_state.matrix_selected_generic_items[original_matrix_key].get('requires_base_choice') and 'chosen_base' in combo:
                            chosen_base_to_remove = combo['chosen_base']
                            if original_matrix_key in st.session_state.user_chosen_base_colors_for_items and chosen_base_to_remove in st.session_state.user_chosen_base_colors_for_items[original_matrix_key]:
                                st.session_state.user_chosen_base_colors_for_items[original_matrix_key].remove(chosen_base_to_remove)
                                if not st.session_state.user_chosen_base_colors_for_items[original_matrix_key]: del st.session_state.user_chosen_base_colors_for_items[original_matrix_key]; del st.session_state.matrix_selected_generic_items[original_matrix_key]
                        else:
                            del st.session_state.matrix_selected_generic_items[original_matrix_key]
                            if original_matrix_key in st.session_state.user_chosen_base_colors_for_items: del st.session_state.user_chosen_base_colors_for_items[original_matrix_key]
                    st.toast(f"Removed: {combo['description']}", icon="🗑️"); st.rerun()
            st.markdown("---")
        else: st.info("Your list is empty. Please select products in Step 2.")

        st.header("Step 4: Download and add to your assortment")
        def prepare_excel_for_download_final():
            if not st.session_state.final_items_for_download or not st.session_state.selected_currency_session: return None
            
            current_selected_currency = st.session_state.selected_currency_session # Already stripped from dropdown
            # User wants specific naming convention for output columns
            ws_price_col = f"{current_selected_currency} - Wholesale"
            rt_price_col = f"{current_selected_currency} - retail" # Note lowercase 'retail' as per user spec
            
            final_cols_temp = []
            for col_template_name in st.session_state.template_cols:
                if col_template_name.lower() == "wholesale price": final_cols_temp.append(ws_price_col)
                elif col_template_name.lower() == "retail price": final_cols_temp.append(rt_price_col)
                else: final_cols_temp.append(col_template_name)
            final_cols = list(dict.fromkeys(final_cols_temp))
            
            output_data = []
            # Use the determined key name for price files from session state
            price_file_key = st.session_state.price_file_article_no_key 
            
            # Determine which price DFs to use based on currency group
            active_ws_df, active_rt_df = None, None
            if current_selected_currency in CURRENCIES_GROUP_1_EU:
                active_ws_df = st.session_state.europe_wholesale_df
                active_rt_df = st.session_state.europe_retail_df
            elif current_selected_currency in CURRENCIES_GROUP_2_UK:
                active_ws_df = st.session_state.uk_wholesale_df
                active_rt_df = st.session_state.uk_retail_df
            
            if active_ws_df is None and active_rt_df is None:
                st.error(f"Price data not available for the selected currency group of '{current_selected_currency}'.")
                return None # Cannot generate file

            for combo in st.session_state.final_items_for_download:
                item_data_row_df = st.session_state.raw_df_original[st.session_state.raw_df_original[RAW_DATA_ITEM_NO_COL] == combo['item_no']]
                if not item_data_row_df.empty:
                    item_data_row = item_data_row_df.iloc[0].copy()
                    output_row = {col: item_data_row.get(col) for col in final_cols if col not in [ws_price_col, rt_price_col]}
                    
                    # Get the cleaned Article No from raw_data (already cleaned during load)
                    # User specifically wants "Article No" column from raw-data for this lookup.
                    lookup_article_no_from_raw = item_data_row.get(RAW_DATA_ARTICLE_NO_COL) 
                    
                    if lookup_article_no_from_raw is None or pd.isna(lookup_article_no_from_raw):
                        output_row[ws_price_col], output_row[rt_price_col] = "ArticleNo Missing in Raw", "ArticleNo Missing in Raw"
                        output_data.append(output_row); continue

                    # Wholesale Price
                    if active_ws_df is not None and not active_ws_df.empty:
                        ws_price_df_match = active_ws_df[active_ws_df[price_file_key] == lookup_article_no_from_raw]
                        if not ws_price_df_match.empty and current_selected_currency in ws_price_df_match.columns:
                            price = ws_price_df_match.iloc[0][current_selected_currency]
                            output_row[ws_price_col] = price if pd.notna(price) else "N/A"
                        else: output_row[ws_price_col] = "Price Not Found"
                    else: output_row[ws_price_col] = "WS Price File Error"
                    # Retail Price
                    if active_rt_df is not None and not active_rt_df.empty:
                        rt_price_df_match = active_rt_df[active_rt_df[price_file_key] == lookup_article_no_from_raw]
                        if not rt_price_df_match.empty and current_selected_currency in rt_price_df_match.columns:
                            price = rt_price_df_match.iloc[0][current_selected_currency]
                            output_row[rt_price_col] = price if pd.notna(price) else "N/A"
                        else: output_row[rt_price_col] = "Price Not Found"
                    else: output_row[rt_price_col] = "RT Price File Error"
                    output_data.append(output_row)
            
            if not output_data: return None
            output_df = pd.DataFrame(output_data, columns=final_cols)
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer: output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
            return excel_buffer.getvalue()

        if st.session_state.final_items_for_download and st.session_state.selected_currency_session:
            file_bytes = prepare_excel_for_download_final()
            if file_bytes:
                st.download_button(label="Generate", data=file_bytes, file_name=f"masterdata_output_{st.session_state.selected_currency_session.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_button_final_v5")
        else:
            st.button("Generate", disabled=True, key="generate_disabled_final_v5", help="Select currency and items first.")

    elif not st.session_state.selected_currency_session and st.session_state.raw_df_original is not None:
        st.info("Please select a currency in Step 1 to proceed.")
    elif st.session_state.raw_df is not None and st.session_state.raw_df.empty and st.session_state.selected_currency_session:
        st.warning(f"No products found for market associated with currency '{st.session_state.selected_currency_session}'.")


elif not files_loaded_successfully and errors:
    st.error("Application initialization failed due to data loading issues. Please review error messages above and check your Excel files.")
elif not files_loaded_successfully:
    st.error("An unexpected issue occurred during data loading. Please check your files and script configuration.")

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
    div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] > div[data-testid="stMarkdown"] > div[data-testid="stMarkdownContainer"] {
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 100%; 
        height: 100%; 
        box-sizing: border-box;
    }


    /* --- Checkbox Styling --- */
    div.stCheckbox { 
         margin: 0 !important;
         padding: 0 !important; 
         display: flex !important;
         align-items: center !important;
         justify-content: center !important;
         width: 20px !important; 
         height: 20px !important; 
         box-sizing: border-box !important;
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] { 
        width: 20px !important; 
        height: 20px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        padding: 0 !important;
        margin: 0 !important;
        box-sizing: border-box !important;
    }
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
        display: flex !important; 
        align-items: center !important;
        justify-content: center !important;
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child svg {
        fill: #FFFFFF !important; 
        width: 12px !important; 
        height: 12px !important;
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"]:has(input[type="checkbox"][aria-checked="true"]) > span:first-child {
        background-color: #5B4A14 !important; 
        border-color: #5B4A14 !important; 
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"]:has(input[type="checkbox"][aria-checked="true"]) > span:first-child svg {
        fill: #FFFFFF !important; 
    }

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
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"].st-ei,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"].st-eh,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"]
     {
        background-color: transparent !important; 
        background-image: none !important; 
        border: 1px solid #000000 !important; 
        border-radius: 0.25rem !important; 
        padding: 0.2em 0.4em !important; 
        line-height: 1.2 !important; 
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[title] {
        color: #000000 !important; 
        font-size: 0.85em !important;
        line-height: inherit !important; 
        margin-right: 4px !important; 
        vertical-align: middle !important; 
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[aria-hidden="true"] {
        display: inline-flex !important; 
        align-items: center !important;
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span[data-baseweb="tag"][aria-selected="true"] > span[aria-hidden="true"] svg {
        fill: #000000 !important; 
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
