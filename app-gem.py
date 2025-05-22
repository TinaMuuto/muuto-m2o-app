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
PRICE_MATRIX_XLSX_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx")
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
    st.title("Muuto made-to-order master data tool") # Sentence case

with top_col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)
    else:
        st.error(f"Muuto Logo not found. Expected at: {LOGO_PATH}. Please ensure 'muuto_logo.png' is in the script's directory.")


# --- App Introduction ---
st.markdown("""
Welcome to Muuto's Made-to-Order (MTO) Product Configurator!

This tool simplifies selecting MTO products and generating the data you need for your systems. Here's how it works:

* **Step 1: Select product family & combinations:**
    * Choose a product family to view its available products and upholstery options.
    * Select your desired product, upholstery, and color combinations directly in the matrix.
    * **Step 1a: Specify base colors:** For items where multiple base colors are available, you can select one or more options.
* **Step 2: Review selections:**
    * Review the final list of configured products. You can remove items from this list if needed.
* **Step 3: Select currency:**
    * Choose your preferred currency for pricing.
* **Step 4: Generate master data file:**
    * After making your selections and choosing a currency, generate and download an Excel file containing all master data for your selected items.
""")

# --- Initialize session state variables ---
if 'raw_df' not in st.session_state: st.session_state.raw_df = None
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

if st.session_state.raw_df is None:
    if os.path.exists(RAW_DATA_XLSX_PATH):
        try:
            st.session_state.raw_df = pd.read_excel(RAW_DATA_XLSX_PATH, sheet_name=RAW_DATA_APP_SHEET)
            required_cols = ['Product Type', 'Product Model', 'Sofa Direction', 'Base Color', 'Product Family', 'Item No', 'Article No', 'Image URL swatch', 'Upholstery Type', 'Upholstery Color', 'Market']
            missing = [col for col in required_cols if col not in st.session_state.raw_df.columns]
            if missing:
                st.error(f"Required columns are missing in '{os.path.basename(RAW_DATA_XLSX_PATH)}': {missing}.")
                files_loaded_successfully = False
            else:
                st.session_state.raw_df = st.session_state.raw_df[st.session_state.raw_df['Market'].astype(str).str.upper() != 'UK']
                st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
                st.session_state.raw_df['Base Color Cleaned'] = st.session_state.raw_df['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)
                st.session_state.raw_df['Upholstery Type'] = st.session_state.raw_df['Upholstery Type'].astype(str).str.strip()
        except Exception as e: st.error(f"Error loading Raw Data: {e}"); files_loaded_successfully = False
    else: st.error(f"Raw Data file not found at: {RAW_DATA_XLSX_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.wholesale_prices_df is None:
    if os.path.exists(PRICE_MATRIX_XLSX_PATH):
        try: st.session_state.wholesale_prices_df = pd.read_excel(PRICE_MATRIX_XLSX_PATH, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
        except Exception as e: st.error(f"Error loading Wholesale Prices: {e}"); files_loaded_successfully = False
    else: st.error(f"Price Matrix file not found: {PRICE_MATRIX_XLSX_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.retail_prices_df is None:
    if os.path.exists(PRICE_MATRIX_XLSX_PATH):
        try: st.session_state.retail_prices_df = pd.read_excel(PRICE_MATRIX_XLSX_PATH, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
        except Exception as e: st.error(f"Error loading Retail Prices: {e}"); files_loaded_successfully = False
    else: st.error(f"Price Matrix file not found: {PRICE_MATRIX_XLSX_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.template_cols is None:
    if os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH):
        try:
            st.session_state.template_cols = pd.read_excel(MASTERDATA_TEMPLATE_XLSX_PATH).columns.tolist()
            if "Wholesale price" not in st.session_state.template_cols:
                st.session_state.template_cols.append("Wholesale price")
            if "Retail price" not in st.session_state.template_cols:
                st.session_state.template_cols.append("Retail price")
        except Exception as e: st.error(f"Error loading Template: {e}"); files_loaded_successfully = False
    else: st.error(f"Template file not found: {MASTERDATA_TEMPLATE_XLSX_PATH}"); files_loaded_successfully = False

# --- Main Application Area ---
if files_loaded_successfully:

    st.header("Step 1: Select product combinations (product / upholstery / color)") # Sentence case

    df_for_display = st.session_state.raw_df.copy()

    available_families_in_view = [DEFAULT_NO_SELECTION] + sorted(df_for_display['Product Family'].dropna().unique()) if 'Product Family' in df_for_display.columns else [DEFAULT_NO_SELECTION]
    if st.session_state.selected_family_session not in available_families_in_view:
        st.session_state.selected_family_session = DEFAULT_NO_SELECTION

    selected_family_idx = 0
    if st.session_state.selected_family_session in available_families_in_view:
        selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)

    selected_family = st.selectbox("Select Product Family:", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
    st.session_state.selected_family_session = selected_family

    def handle_matrix_cb_toggle(prod_name, uph_type, uph_color, checkbox_key_matrix):
        is_checked = st.session_state[checkbox_key_matrix]
        generic_item_key = f"{selected_family}_{prod_name}_{uph_type}_{uph_color}".replace(" ", "_").replace("/","_").replace("(","").replace(")","")

        if is_checked:
            matching_items = st.session_state.raw_df[
                (st.session_state.raw_df['Product Family'] == selected_family) &
                (st.session_state.raw_df['Product Display Name'] == prod_name) &
                (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == uph_type) &
                (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == uph_color)
            ]
            if not matching_items.empty:
                unique_base_colors = matching_items['Base Color Cleaned'].dropna().unique().tolist()
                first_item_match = matching_items.iloc[0]

                item_data = {
                    'key': generic_item_key, 'family': selected_family, 'product': prod_name,
                    'upholstery_type': uph_type, 'upholstery_color': uph_color,
                    'requires_base_choice': len(unique_base_colors) > 1,
                    'available_bases': unique_base_colors if len(unique_base_colors) > 1 else [],
                    'item_no_if_single_base': first_item_match['Item No'] if len(unique_base_colors) <= 1 else None,
                    'article_no_if_single_base': first_item_match['Article No'] if len(unique_base_colors) <= 1 else None,
                    'resolved_base_if_single': unique_base_colors[0] if len(unique_base_colors) == 1 else (pd.NA if not unique_base_colors else None)
                }
                st.session_state.matrix_selected_generic_items[generic_item_key] = item_data
                st.toast(f"Selected: {prod_name} / {uph_type} / {uph_color}", icon="➕")
        else:
            if generic_item_key in st.session_state.matrix_selected_generic_items:
                del st.session_state.matrix_selected_generic_items[generic_item_key]
                if generic_item_key in st.session_state.user_chosen_base_colors_for_items:
                    del st.session_state.user_chosen_base_colors_for_items[generic_item_key]
                st.toast(f"Deselected: {prod_name} / {uph_type} / {uph_color}", icon="➖")

    def handle_base_color_multiselect_change(item_key_for_base_select):
        multiselect_widget_key = f"ms_base_{item_key_for_base_select}"
        st.session_state.user_chosen_base_colors_for_items[item_key_for_base_select] = st.session_state[multiselect_widget_key]


    if selected_family and selected_family != DEFAULT_NO_SELECTION and 'Product Family' in df_for_display.columns:
        family_df = df_for_display[df_for_display['Product Family'] == selected_family]
        if not family_df.empty and 'Upholstery Type' in family_df.columns:
            products_in_family = sorted(family_df['Product Display Name'].dropna().unique())
            upholstery_types_in_family = sorted(family_df['Upholstery Type'].dropna().unique())

            if not products_in_family: st.info(f"No products in the family: {selected_family}")
            elif not upholstery_types_in_family: st.info(f"No upholstery types for the product family: {selected_family}")
            else:
                header_upholstery_types = ["Product"]
                header_swatches = [" "]
                header_color_numbers = [" "]
                data_column_map = []

                for uph_type_clean in upholstery_types_in_family:
                    colors_for_type_df = family_df[family_df['Upholstery Type'] == uph_type_clean][['Upholstery Color', 'Image URL swatch']].drop_duplicates().sort_values(by='Upholstery Color')
                    if not colors_for_type_df.empty:
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
                    cols_uph_type_header = st.columns([2.5] + [1] * num_data_columns)
                    current_uph_type_header_display = None
                    for i, col_widget in enumerate(cols_uph_type_header):
                        if i == 0:
                            with col_widget:
                                st.caption("")
                        else:
                            map_entry = data_column_map[i-1]
                            if map_entry['uph_type'] != current_uph_type_header_display:
                                with col_widget:
                                    st.caption(f"<div class='upholstery-header'>{map_entry['uph_type']}</div>", unsafe_allow_html=True)
                                current_uph_type_header_display = map_entry['uph_type']

                    cols_swatch_header = st.columns([2.5] + [1] * num_data_columns)
                    for i, col_widget in enumerate(cols_swatch_header):
                        if i == 0:
                            with col_widget:
                                st.markdown("<div class='zoom-instruction'><br>Click swatch in header to zoom</div>", unsafe_allow_html=True)
                        else:
                            sw_url = data_column_map[i-1]['swatch']
                            with col_widget:
                                if sw_url and pd.notna(sw_url):
                                    st.image(sw_url, width=30)
                                else:
                                    st.markdown("<div class='swatch-placeholder'></div>", unsafe_allow_html=True)

                    cols_color_num_header = st.columns([2.5] + [1] * num_data_columns)
                    for i, col_widget in enumerate(cols_color_num_header):
                        if i == 0:
                            with col_widget:
                                st.caption("")
                        else:
                            with col_widget:
                                st.caption(f"<small>{data_column_map[i-1]['uph_color']}</small>", unsafe_allow_html=True)

                    st.markdown("---") # This HR separates headers from product rows

                    for prod_name in products_in_family:
                        # Use vertical_alignment for the columns in this row
                        cols_product_row = st.columns([2.5] + [1] * num_data_columns, vertical_alignment="center")
                        with cols_product_row[0]:
                            # Wrap product name in a div for consistent cell styling if needed, or rely on column alignment
                            st.markdown(f"<div class='product-name-cell'>**{prod_name}**</div>", unsafe_allow_html=True)


                        for i, col_widget in enumerate(cols_product_row[1:]):
                            current_col_uph_type_filter = data_column_map[i]['uph_type']
                            current_col_uph_color_filter = data_column_map[i]['uph_color']

                            item_exists_df = family_df[
                                (family_df['Product Display Name'] == prod_name) &
                                (family_df['Upholstery Type'] == current_col_uph_type_filter) &
                                (family_df['Upholstery Color'].astype(str).fillna("N/A") == current_col_uph_color_filter)
                            ]

                            # The cell_container is the col_widget.container()
                            # The content inside this (checkbox or grey box) will be centered by CSS
                            # on div[data-testid="stHorizontalBlock"] > div > div[data-testid="stVerticalBlock"]
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
                                unique_key_for_grey_box = f"greybox_{selected_family}_{prod_name}_{current_col_uph_type_filter}_{current_col_uph_color_filter}_{i}"
                                cell_container.markdown(f"<div class='unavailable-matrix-cell' id='{unique_key_for_grey_box}'></div>", unsafe_allow_html=True)
        else:
            if selected_family and selected_family != DEFAULT_NO_SELECTION : st.info(f"No data found for product family: {selected_family}")

    # --- Step 1a: Specify Base Colors ---
    items_needing_base_choice_now = [
        item_data for key, item_data in st.session_state.matrix_selected_generic_items.items() if item_data.get('requires_base_choice')
    ]
    if items_needing_base_choice_now:
        st.subheader("Step 1a: Specify base colors") # Text change
        for generic_item in items_needing_base_choice_now:
            item_key = generic_item['key']
            multiselect_key = f"ms_base_{item_key}"
            st.markdown(f"**{generic_item['product']}** ({generic_item['upholstery_type']} - {generic_item['upholstery_color']})")

            current_selection_for_this_item = st.session_state.user_chosen_base_colors_for_items.get(item_key, [])

            st.multiselect(
                f"Available base colors. You can select multiple:",
                options=generic_item['available_bases'],
                default=current_selection_for_this_item,
                key=multiselect_key,
                on_change=handle_base_color_multiselect_change,
                args=(item_key,)
            )
            st.markdown("---")

    # --- Step 2: Review Selections ---
    st.header("Step 2: Review selections") # Sentence case
    _current_final_items = []
    for key, gen_item_data in st.session_state.matrix_selected_generic_items.items():
        if not gen_item_data['requires_base_choice']:
            if gen_item_data.get('item_no_if_single_base') is not None:
                _current_final_items.append({
                    "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']}" + (f" / Base: {gen_item_data['resolved_base_if_single']}" if pd.notna(gen_item_data['resolved_base_if_single']) else ""),
                    "item_no": gen_item_data['item_no_if_single_base'],
                    "article_no": gen_item_data['article_no_if_single_base'],
                    "key_in_matrix": key
                })
        else:
            selected_bases_for_this = st.session_state.user_chosen_base_colors_for_items.get(key, [])
            for bc in selected_bases_for_this:
                specific_item_df = st.session_state.raw_df[
                    (st.session_state.raw_df['Product Family'] == gen_item_data['family']) &
                    (st.session_state.raw_df['Product Display Name'] == gen_item_data['product']) &
                    (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == gen_item_data['upholstery_type']) &
                    (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == gen_item_data['upholstery_color']) &
                    (st.session_state.raw_df['Base Color Cleaned'].fillna("N/A") == bc)
                ]
                if not specific_item_df.empty:
                    actual_item = specific_item_df.iloc[0]
                    _current_final_items.append({
                        "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']} / Base: {bc}",
                        "item_no": actual_item['Item No'],
                        "article_no": actual_item['Article No'],
                        "key_in_matrix": key,
                        "chosen_base": bc
                    })

    temp_final_list_review = []
    seen_item_nos_review = set()
    for item_rev in _current_final_items:
        unique_key_rev = f"{item_rev['item_no']}_{item_rev.get('chosen_base', 'single_base')}"
        if unique_key_rev not in seen_item_nos_review:
            temp_final_list_review.append(item_rev)
            seen_item_nos_review.add(unique_key_rev)
    st.session_state.final_items_for_download = temp_final_list_review

    if st.session_state.final_items_for_download:
        st.markdown("**Current Selections for Download:**")
        for i, combo in enumerate(st.session_state.final_items_for_download):
            col1_rev, col2_rev = st.columns([0.9, 0.1])
            col1_rev.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")
            if col2_rev.button(f"Remove", key=f"final_review_remove_{i}_{combo['item_no']}_{combo.get('chosen_base','nobase')}"):
                original_matrix_key = combo['key_in_matrix']
                if original_matrix_key in st.session_state.matrix_selected_generic_items:
                    if st.session_state.matrix_selected_generic_items[original_matrix_key].get('requires_base_choice') and 'chosen_base' in combo:
                        if original_matrix_key in st.session_state.user_chosen_base_colors_for_items:
                            if combo['chosen_base'] in st.session_state.user_chosen_base_colors_for_items[original_matrix_key]:
                                st.session_state.user_chosen_base_colors_for_items[original_matrix_key].remove(combo['chosen_base'])
                                if not st.session_state.user_chosen_base_colors_for_items[original_matrix_key]:
                                    if not any(st.session_state.user_chosen_base_colors_for_items.get(original_matrix_key, [])):
                                         del st.session_state.matrix_selected_generic_items[original_matrix_key]
                    else:
                        if original_matrix_key in st.session_state.matrix_selected_generic_items:
                            del st.session_state.matrix_selected_generic_items[original_matrix_key]
                st.session_state.final_items_for_download.pop(i)
                st.toast(f"Removed: {combo['description']}", icon="🗑️")
                st.rerun()
        st.markdown("---")

    # --- Step 3: Select Currency ---
    st.header("Step 3: Select currency") # Sentence case
    selected_currency = None
    try:
        if not st.session_state.wholesale_prices_df.empty:
            article_no_col_name_ws = st.session_state.wholesale_prices_df.columns[0]
            currency_options = [DEFAULT_NO_SELECTION] + [col for col in st.session_state.wholesale_prices_df.columns if str(col).lower() != str(article_no_col_name_ws).lower()]
        else:
            currency_options = [DEFAULT_NO_SELECTION]

        current_currency_idx = 0
        if st.session_state.selected_currency_session and st.session_state.selected_currency_session in currency_options:
            current_currency_idx = currency_options.index(st.session_state.selected_currency_session)

        selected_currency_choice = st.selectbox("Select Currency:", options=currency_options, index=current_currency_idx, key="currency_selector_main_final")

        if selected_currency_choice and selected_currency_choice != DEFAULT_NO_SELECTION:
            st.session_state.selected_currency_session = selected_currency_choice
            selected_currency = selected_currency_choice
        else:
            st.session_state.selected_currency_session = None
            selected_currency = None

        if not currency_options or len(currency_options) <=1 :
            if not st.session_state.wholesale_prices_df.empty:
                st.error("No currency columns found in Price Matrix.")
    except Exception as e:
        st.error(f"Error with currency selection: {e}")
        selected_currency = None


    # --- Step 4: Generate Master Data File ---
    st.header("Step 4: Generate master data file") # Sentence case

    def prepare_excel_for_download_final():
        if not st.session_state.final_items_for_download:
            return None
        current_selected_currency_for_dl = st.session_state.selected_currency_session
        if not current_selected_currency_for_dl:
            return None

        output_data = []
        ws_price_col_name_dynamic = f"Wholesale price ({current_selected_currency_for_dl})"
        rt_price_col_name_dynamic = f"Retail price ({current_selected_currency_for_dl})"

        final_cols = []
        seen_output_cols = set()
        for col_template in st.session_state.template_cols:
            col_template_lower = col_template.lower()
            if col_template_lower == "wholesale price":
                target_col = ws_price_col_name_dynamic
            elif col_template_lower == "retail price":
                target_col = rt_price_col_name_dynamic
            else:
                target_col = col_template
            if target_col not in seen_output_cols:
                final_cols.append(target_col)
                seen_output_cols.add(target_col)
        master_template_columns_final_output = final_cols

        for combo_selection in st.session_state.final_items_for_download:
            item_no_to_find = combo_selection['item_no']; article_no_to_find = combo_selection['article_no']
            item_data_row_series_df = st.session_state.raw_df[st.session_state.raw_df['Item No'] == item_no_to_find]
            if not item_data_row_series_df.empty:
                item_data_row_series = item_data_row_series_df.iloc[0]; output_row_dict = {}
                for col_template in master_template_columns_final_output:
                    if col_template == ws_price_col_name_dynamic or col_template == rt_price_col_name_dynamic: continue
                    if col_template in item_data_row_series.index: output_row_dict[col_template] = item_data_row_series[col_template]
                    else: output_row_dict[col_template] = None
                if not st.session_state.wholesale_prices_df.empty:
                    ws_price_row_df = st.session_state.wholesale_prices_df[st.session_state.wholesale_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                    if not ws_price_row_df.empty and current_selected_currency_for_dl in ws_price_row_df.columns: output_row_dict[ws_price_col_name_dynamic] = ws_price_row_df.iloc[0][current_selected_currency_for_dl] if pd.notna(ws_price_row_df.iloc[0][current_selected_currency_for_dl]) else "N/A"
                    else: output_row_dict[ws_price_col_name_dynamic] = "Price Not Found"
                else: output_row_dict[ws_price_col_name_dynamic] = "Wholesale Matrix Empty"
                if not st.session_state.retail_prices_df.empty:
                    rt_price_row_df = st.session_state.retail_prices_df[st.session_state.retail_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                    if not rt_price_row_df.empty and current_selected_currency_for_dl in rt_price_row_df.columns: output_row_dict[rt_price_col_name_dynamic] = rt_price_row_df.iloc[0][current_selected_currency_for_dl] if pd.notna(rt_price_row_df.iloc[0][current_selected_currency_for_dl]) else "N/A"
                    else: output_row_dict[rt_price_col_name_dynamic] = "Price Not Found"
                else: output_row_dict[rt_price_col_name_dynamic] = "Retail Matrix Empty"
                output_data.append(output_row_dict)

        if not output_data:
            return None

        output_df = pd.DataFrame(output_data, columns=master_template_columns_final_output)
        output_excel_buffer = io.BytesIO()
        with pd.ExcelWriter(output_excel_buffer, engine='xlsxwriter') as writer: output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
        return output_excel_buffer.getvalue()

    can_download_now = bool(st.session_state.final_items_for_download and st.session_state.selected_currency_session)

    if can_download_now:
        file_bytes = prepare_excel_for_download_final()
        if file_bytes:
            st.download_button(
                label="Generate and Download Master Data File",
                data=file_bytes,
                file_name=f"masterdata_output_{st.session_state.selected_currency_session.replace(' ', '_').replace('.', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="final_download_action_button_v7",
                help="Click to generate and download your customized master data file."
            )
    else:
        st.button("Generate Master Data File", key="generate_file_disabled_button_v5", disabled=True, help="Please ensure items are selected (Step 1), reviewed (Step 2), and a currency is chosen (Step 3).")

else:
    st.error("One or more data files could not be loaded correctly, or required columns are missing. Please check file paths, formats, and column names in your .xlsx files.")


# --- Styling (Optional) ---
st.markdown("""
<style>
    /* Apply background color to the main app container and body */
    .stApp, body {
        background-color: #EFEEEB !important;
    }
    .main .block-container {
        background-color: #EFEEEB !important;
        padding-top: 2rem; /* Add some padding at the top of the main content */
    }

    h1, h2, h3 { /* General Header Styling */
        text-transform: none !important; /* Override any theme-based transforms */
    }
    h1 { color: #333; } /* App Title */
    h2 { /* Step Headers */
        color: #1E40AF;
        padding-bottom: 5px;
        margin-top: 30px;
        margin-bottom: 15px;
    }
     h3 { /* Sub-step Headers like 1a */
        color: #1E40AF;
        font-size: 1.25em;
        padding-bottom: 3px;
        margin-top: 20px;
        margin-bottom: 10px;
    }

    /* Styling for the matrix-like headers */
    div[data-testid="stCaptionContainer"] > div > p { /* Container for Upholstery Type and Color Number */
        font-weight: bold;
        font-size: 0.8em !important; /* Slightly larger for readability */
        color: #31333F !important; /* Standard dark text color */
        text-align: center;
        white-space: normal;
        overflow-wrap:break-word;
        line-height: 1.2; /* Adjusted line height */
        padding: 2px;
    }
    .upholstery-header { /* Specifically for Upholstery Type text */
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: clip !important;
        display: block;
        max-width: 100%;
        line-height: 1.2;
        color: #31333F !important; /* Standard dark text color */
        text-transform: capitalize !important; /* First letter of each word capitalized */
        font-weight: bold !important;
        font-size: 0.8em !important;
    }
    div[data-testid="stCaptionContainer"] small { /* Color numbers */
        color: #31333F !important; /* Standard dark text color */
        font-weight: normal !important; /* Make it normal weight to differentiate from Upholstery Type */
        font-size: 0.75em !important;
    }

    div[data-testid="stCaptionContainer"] img { /* Swatch in header */
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
    /* Product name cell styling */
    .product-name-cell {
        display: flex;
        align-items: center; /* Vertically center product name */
        height: 30px; /* Match the height of checkbox/grey box container */
        padding-left: 5px; /* Optional: add some padding */
    }
    .product-name-cell strong { /* Target the bold text within product name */
        color: #31333F !important; /* Standard dark text color */
    }


    /* Container for checkbox or grey box within each matrix data cell */
    div[data-testid="stHorizontalBlock"] > div > div[data-testid="stVerticalBlock"] {
        height: 30px !important; /* Fixed height for the entire cell content area */
        min-height: 30px !important;
        display: flex !important;
        align-items: center !important; /* Vertically center child (checkbox or grey box) */
        justify-content: center !important; /* Horizontally center child */
        padding: 0 !important; /* Remove padding if it misaligns */
        margin: 0 !important; /* Remove margin */

    }

    /* --- Checkbox Styling --- */
    div.stCheckbox {
         margin: 0 !important;
         display: flex !important;
         align-items: center !important;
         justify-content: center !important;
         width: 100%; /* Ensure it takes full width of its small column cell */
         height: 100%; /* Ensure it takes full height of its stVerticalBlock parent */
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] {
        width: 20px !important; /* Control size of the label containing the checkbox parts */
        height: 20px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child { /* The visual square */
        background-color: #5B4A14 !important;
        border-color: #5B4A14 !important;
        box-shadow: none !important;
        width: 20px !important; /* Ensure the square is 20x20 */
        height: 20px !important;
        border-radius: 0.25rem !important;
    }
    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child svg { /* The checkmark */
        fill: white !important;
    }


    /* --- Unavailable Matrix Cell (Grey Box) Styling --- */
    .unavailable-matrix-cell {
        width: 20px !important;
        height: 20px !important;
        min-width: 20px !important;
        min-height: 20px !important;
        background-color: #e9ecef !important;
        border: 1px solid #ced4da !important;
        border-radius: 0.25rem !important;
        /* Centered by its parent stVerticalBlock */
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
    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"],
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"],
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"] {
        border: 1px solid #5B4A14 !important;
        background-color: #FFFFFF !important;
        color: #5B4A14 !important;
        padding: 0.375rem 0.75rem !important;
        font-size: 1rem !important;
        line-height: 1.5 !important;
        border-radius: 0.25rem !important;
        transition: color 0.15s ease-in-out, background-color 0.15s ease-in-out, border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out !important;
        font-weight: 500 !important;
        text-transform: none !important; /* Ensure button text is not all caps */
    }
    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"] p, /* Target p tag inside button */
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"] p,
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"] p {
        color: inherit !important; /* Inherit color from button */
        text-transform: none !important; /* Ensure button text is not all caps */
    }

    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"]:hover,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"]:hover,
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"]:hover {
        background-color: #5B4A14 !important;
        color: #FFFFFF !important;
        border-color: #5B4A14 !important;
    }
    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"]:hover p,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"]:hover p,
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"]:hover p {
        color: #FFFFFF !important; /* Text color on hover */
    }


    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"]:active,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"]:focus,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"]:active,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"]:focus,
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"]:active,
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"]:focus {
        background-color: #4A3D10 !important; /* Darker Muuto Gold */
        color: #FFFFFF !important;
        border-color: #4A3D10 !important;
        box-shadow: 0 0 0 0.2rem rgba(91, 74, 20, 0.4) !important;
        outline: none !important;
    }
    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"]:active p,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-primary"]:focus p,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"]:active p,
    div[data-testid="stButton"] > button[data-testid="stBaseButton-secondary"]:focus p,
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"]:active p,
    div[data-testid="stDownloadButton"] > button[data-testid="stBaseButton-secondary"]:focus p {
        color: #FFFFFF !important; /* Text color on active/focus */
    }


    small {
        color: #718096; /* Keep this for other small text if needed, or override if specifically for color numbers */
        font-size:0.9em;
        display:block;
        line-height:1.1;
    }
    /* --- Multiselect Tags Styling --- */
    div[data-testid="stMultiSelect"] span[data-baseweb="tag"][aria-selected="true"] {
        background-color: #5B4A14 !important;
    }
    div[data-testid="stMultiSelect"] span[data-baseweb="tag"][aria-selected="true"] > span[title] {
        color: white !important;
        font-size: 0.85em !important;
    }
    div[data-testid="stMultiSelect"] span[data-baseweb="tag"][aria-selected="true"] span[role="button"] svg {
        fill: white !important;
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
