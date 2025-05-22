import streamlit as st
import pandas as pd
import io
import os

# --- Page Configuration (MUST BE THE FIRST STREAMLIT COMMAND) ---
st.set_page_config(
    layout="wide",
    page_title="Muuto M2O", 
    page_icon="favicon.png" 
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
    st.title("Muuto Made-to-Order Master Data Tool")

with top_col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120) 
    else:
        st.error(f"Muuto Logo not found. Expected at: {LOGO_PATH}. Please ensure 'muuto_logo.png' is in the script's directory.")


# --- App Introduction ---
st.markdown("""
Welcome to Muuto's Made-to-Order (MTO) Product Configurator!

This tool simplifies selecting MTO products and generating the data you need for your systems. Here's how it works:

* **Select Product Family & Combinations:**
    * Choose a product family to view its available products and upholstery options.
    * Select your desired product, upholstery, and color combinations directly in the matrix. You can easily switch product families to add more items to your overall selection.
* **Specify Base Colors (if applicable):**
    * For items where multiple base colors are available, you can select one or more options using the multiselect feature.
* **Confirm & Review Selections:**
    * Confirm your choices and review the final list of configured products. You can remove items from this list if needed.
* **Select Currency:**
    * Choose your preferred currency for pricing.
* **Generate Master Data File:**
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
if files_loaded_successfully and all(df is not None for df in [st.session_state.raw_df, st.session_state.wholesale_prices_df, st.session_state.retail_prices_df]) and st.session_state.template_cols:
    
    st.header("Step 1: Select Product Combinations (Product / Upholstery / Color)")
    
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
                st.warning(f"No item found for {prod_name} / {uph_type} / {uph_color}")
        else:
            if generic_item_key in st.session_state.matrix_selected_generic_items:
                del st.session_state.matrix_selected_generic_items[generic_item_key]
                if generic_item_key in st.session_state.user_chosen_base_colors_for_items:
                    del st.session_state.user_chosen_base_colors_for_items[generic_item_key]
                st.toast(f"Deselected: {prod_name} / {uph_type} / {uph_color}", icon="➖")

    # Corrected Callback for base color multiselect
    def handle_base_color_multiselect_change(item_key_for_base_select):
        multiselect_widget_key = f"ms_base_{item_key_for_base_select}" # Reconstruct the widget key
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
                                st.markdown("<div class='zoom-instruction'><br>(Click swatch in header to zoom)</div>", unsafe_allow_html=True)
                        else:
                            sw_url = data_column_map[i-1]['swatch']
                            with col_widget:
                                if sw_url: 
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
                    
                    st.markdown("---") 

                    for prod_name in products_in_family:
                        cols_product_row = st.columns([2.5] + [1] * num_data_columns)
                        with cols_product_row[0]:
                            st.markdown(f"**{prod_name}**")
                        
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
                                is_gen_selected = cb_key_str in st.session_state.matrix_selected_generic_items
                                
                                cell_container.checkbox(" ", value=is_gen_selected, key=cb_key_str,
                                            on_change=handle_matrix_cb_toggle, 
                                            args=(prod_name, current_col_uph_type_filter, current_col_uph_color_filter, cb_key_str),
                                            label_visibility="collapsed")
                            else:
                                cell_container.markdown("<div class='matrix-cell-empty'>-</div>", unsafe_allow_html=True) 
                        st.markdown("---")
        else:
            if selected_family and selected_family != DEFAULT_NO_SELECTION : st.info(f"No data found for product family: {selected_family}")


    # --- Step 2: Select Base Colors ---
    items_needing_base_choice_now = [
        item_data for key, item_data in st.session_state.matrix_selected_generic_items.items() if item_data.get('requires_base_choice')
    ]
    if items_needing_base_choice_now:
        st.header("Step 2: Select Base Colors")
        for generic_item in items_needing_base_choice_now:
            item_key = generic_item['key']
            multiselect_key = f"ms_base_{item_key}" 
            st.markdown(f"**{generic_item['product']}** ({generic_item['upholstery_type']} - {generic_item['upholstery_color']})")
            
            current_selection_for_this_item = st.session_state.user_chosen_base_colors_for_items.get(item_key, [])

            st.multiselect(
                f"Available base colors:",
                options=generic_item['available_bases'],
                default=current_selection_for_this_item,
                key=multiselect_key, 
                on_change=handle_base_color_multiselect_change, 
                args=(item_key,) 
            )
            st.markdown("---")
    
    # --- Step 3: Confirm Selections & Review List ---
    st.header("Step 3: Confirm Selections & Review List")
    # This button now only populates/updates final_items_for_download for review
    if st.button("Compile Final List for Review", key="compile_list_button"):
        st.session_state.final_items_for_download = [] 
        for key, gen_item_data in st.session_state.matrix_selected_generic_items.items():
            if not gen_item_data['requires_base_choice']:
                if gen_item_data.get('item_no_if_single_base') is not None: 
                    st.session_state.final_items_for_download.append({
                        "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']}" + (f" / Base: {gen_item_data['resolved_base_if_single']}" if pd.notna(gen_item_data['resolved_base_if_single']) else ""),
                        "item_no": gen_item_data['item_no_if_single_base'],
                        "article_no": gen_item_data['article_no_if_single_base'],
                        "key_in_matrix": key # For removal logic
                    })
            else:
                selected_bases_for_this = st.session_state.user_chosen_base_colors_for_items.get(key, [])
                if not selected_bases_for_this:
                    st.warning(f"No base color selected for {gen_item_data['product']} ({gen_item_data['upholstery_type']}/{gen_item_data['upholstery_color']}). This item will not be added to the final list until a base color is chosen.")
                    continue
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
                        st.session_state.final_items_for_download.append({
                            "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']} / Base: {bc}",
                            "item_no": actual_item['Item No'],
                            "article_no": actual_item['Article No'],
                            "key_in_matrix": key, 
                            "chosen_base": bc 
                        })
                    else:
                        st.error(f"ERROR: Could not find item for {gen_item_data['product']} ({gen_item_data['upholstery_type']}/{gen_item_data['upholstery_color']}) with base {bc}. Check data.")
        
        final_list_unique = []
        seen_item_nos_final = set()
        for item in st.session_state.final_items_for_download:
            # Create a more unique key for de-duplication if items can be identical except for description details
            unique_key_for_dedup = f"{item['item_no']}_{item.get('chosen_base', 'single_base')}" 
            if unique_key_for_dedup not in seen_item_nos_final:
                final_list_unique.append(item)
                seen_item_nos_final.add(unique_key_for_dedup)
        st.session_state.final_items_for_download = final_list_unique
        if st.session_state.final_items_for_download or not st.session_state.matrix_selected_generic_items : 
             st.success("Final list compiled for review!")
        st.rerun()

    # Display the compiled list for review
    if st.session_state.final_items_for_download:
        st.markdown("**Final Selections:**")
        for i, combo in enumerate(st.session_state.final_items_for_download):
            col1_rev, col2_rev = st.columns([0.9, 0.1])
            col1_rev.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")
            if col2_rev.button(f"Remove", key=f"final_review_remove_{i}_{combo['item_no']}_{combo.get('chosen_base','nobase')}"):
                # This removal logic needs to be robust to update matrix_selected_generic_items or user_chosen_base_colors_for_items
                # For simplicity now, it just removes from this display list. The user would need to re-compile if they want to change Step 1/2 selections.
                # A more advanced removal would trace back to the original selection and deselect it.
                st.session_state.final_items_for_download.pop(i)
                st.toast(f"Removed: {combo['description']}", icon="🗑️")
                st.rerun()
        st.markdown("---")
    elif st.session_state.matrix_selected_generic_items: # If items are selected in matrix but not compiled yet
        st.info("Click 'Compile Final List for Review' above after making all selections.")


    # --- Step 4: Select Currency ---
    st.header("Step 4: Select Currency")
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
            elif st.session_state.wholesale_prices_df.empty: 
                st.error("Price Matrix (wholesale) is empty.")
    except Exception as e: 
        st.error(f"Error with currency selection: {e}")
        selected_currency = None


    # --- Step 5: Generate Master Data File ---
    st.header("Step 5: Generate Master Data File")
    
    # This function will be called by the download button
    def prepare_and_generate_excel_for_download_final():
        if not st.session_state.final_items_for_download:
            st.warning("No items confirmed for download. Please compile your list in Step 3.")
            return None 
        
        if not st.session_state.selected_currency_session: # Check the session state variable
            st.warning("Please select a currency in Step 4.")
            return None 

        output_data = []
        current_selected_currency = st.session_state.selected_currency_session # Use the stored currency
        ws_price_col_name_dynamic = f"Wholesale price ({current_selected_currency})"
        rt_price_col_name_dynamic = f"Retail price ({current_selected_currency})"
        master_template_columns_final_output = []
        for col in st.session_state.template_cols:
            if col.lower() == "wholesale price": master_template_columns_final_output.append(ws_price_col_name_dynamic)
            elif col.lower() == "retail price": master_template_columns_final_output.append(rt_price_col_name_dynamic)
            else: master_template_columns_final_output.append(col)
        if "Wholesale price" in st.session_state.template_cols and ws_price_col_name_dynamic not in master_template_columns_final_output: master_template_columns_final_output.append(ws_price_col_name_dynamic)
        if "Retail price" in st.session_state.template_cols and rt_price_col_name_dynamic not in master_template_columns_final_output: master_template_columns_final_output.append(rt_price_col_name_dynamic)
        master_template_columns_final_output = [c for c in master_template_columns_final_output if c.lower() != "wholesale price" or c == ws_price_col_name_dynamic]
        master_template_columns_final_output = [c for c in master_template_columns_final_output if c.lower() != "retail price" or c == rt_price_col_name_dynamic]
        seen_cols = set(); unique_ordered_cols = []; 
        for col in master_template_columns_final_output:
            if col not in seen_cols: unique_ordered_cols.append(col); seen_cols.add(col)
        master_template_columns_final_output = unique_ordered_cols

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
                    if not ws_price_row_df.empty and current_selected_currency in ws_price_row_df.columns: output_row_dict[ws_price_col_name_dynamic] = ws_price_row_df.iloc[0][current_selected_currency] if pd.notna(ws_price_row_df.iloc[0][current_selected_currency]) else "N/A"
                    else: output_row_dict[ws_price_col_name_dynamic] = "Price Not Found"
                else: output_row_dict[ws_price_col_name_dynamic] = "Wholesale Matrix Empty"
                if not st.session_state.retail_prices_df.empty:
                    rt_price_row_df = st.session_state.retail_prices_df[st.session_state.retail_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                    if not rt_price_row_df.empty and current_selected_currency in rt_price_row_df.columns: output_row_dict[rt_price_col_name_dynamic] = rt_price_row_df.iloc[0][current_selected_currency] if pd.notna(rt_price_row_df.iloc[0][current_selected_currency]) else "N/A"
                    else: output_row_dict[rt_price_col_name_dynamic] = "Price Not Found"
                else: output_row_dict[rt_price_col_name_dynamic] = "Retail Matrix Empty"
                output_data.append(output_row_dict)
            else: st.warning(f"Data for Item No: {item_no_to_find} not found.")
        
        if not output_data:
            st.warning("No data to generate for the file after processing.")
            return None

        output_df = pd.DataFrame(output_data, columns=master_template_columns_final_output) 
        output_excel_buffer = io.BytesIO()
        with pd.ExcelWriter(output_excel_buffer, engine='xlsxwriter') as writer: output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
        return output_excel_buffer.getvalue()

    # The download button itself
    if st.session_state.final_items_for_download and st.session_state.selected_currency_session:
        st.download_button(
            label="Generate and Download Master Data File",
            data=prepare_and_generate_excel_for_download_final(), 
            file_name=f"masterdata_output_{st.session_state.selected_currency_session.replace(' ', '_').replace('.', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="final_download_action_button_v2", # Ensure unique key
            help="Click to generate and download your customized master data file."
        )
    elif not st.session_state.final_items_for_download:
        st.warning("Please compile your selections in Step 3 first.")
    elif not st.session_state.selected_currency_session:
        st.warning("Please select a currency in Step 4 before generating the file.")


else: 
    st.error("One or more data files could not be loaded correctly, or required columns are missing. Please check file paths, formats, and column names in your .xlsx files.")


# --- Styling (Optional) ---
st.markdown("""
<style>
    /* Apply background color to the main app container and body */
    .stApp, body {
        background-color: #EFEEEB !important;
    }
    /* More specific selector for main content area if needed */
    .main .block-container {
        background-color: #EFEEEB !important; 
    }

    h1 { color: #333; } /* App Title */
    h2 { /* Step Headers */
        color: #1E40AF; 
        /* border-bottom: 2px solid #BFDBFE; */ /* Removed blue line */
        padding-bottom: 5px; 
        margin-top: 30px; 
        margin-bottom: 15px; 
    }
    h3 { /* Product Subheaders in new layout */
        font-size: 1.1em; 
        font-weight: bold; 
        margin-top: 15px; 
        margin-bottom: 5px;
    }
    /* Styling for the matrix-like headers */
    div[data-testid="stCaptionContainer"] > div > p { 
        font-weight: bold; 
        font-size: 0.8em !important; 
        color: #4A5568 !important; 
        text-align: center; 
        white-space: normal; 
        overflow-wrap: break-word; 
        line-height: 1.2;
        padding: 2px;
    }
    /* Specifically target Upholstery Type headers for no-wrap */
    .upholstery-header {
        white-space: normal !important; /* Allow wrapping */
        overflow: visible !important; /* Allow overflow if needed, or adjust height */
        text-overflow: clip !important; /* Remove ellipsis */
        display: block; 
        max-width: 100%; 
        line-height: 1.3; /* Adjust line height for wrapped text */
    }
    div[data-testid="stCaptionContainer"] img { /* Swatch in header */
        max-height: 25px !important; 
        width: 25px !important;    
        object-fit: cover !important; /* Enforce square and cover */       
        margin-right: 3px; 
    }
    /* Placeholder for empty swatch in header */
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
    /* Placeholder for empty matrix cell */
    .matrix-cell-empty {
        height:30px; 
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.8em;
        color: #aaa;
    }
    .zoom-instruction {
        font-size: 0.75em; 
        color: #555; 
        text-align: left; 
        padding-top: 10px; 
    }
    
    /* Custom styling for the checkbox itself when checked */
    /* More specific selector to override Streamlit's default theme */
    label[data-testid="stCheckbox"] input[type="checkbox"]:checked + div {
        background-color: #5B4A14 !important; 
        border-color: #5B4A14 !important;
    }
    label[data-testid="stCheckbox"] input[type="checkbox"]:checked + div svg { /* Checkmark color */
        fill: white !important; 
    }
    /* For the little square inside the checkbox tick bar - this is often the one showing the theme color */
    label[data-testid="stCheckbox"] input[type="checkbox"]:checked + div[data-testid="stTickBar"] > div[data-testid="stTickSquare"] {
        background-color: #5B4A14 !important; 
        border-color: #5B4A14 !important; 
    }
    label[data-testid="stCheckbox"] input[type="checkbox"]:checked + div[data-testid="stTickBar"] > div[data-testid="stTickSquare"] svg {
        fill: white !important; 
    }


    hr { 
        margin-top: 0.2rem; 
        margin-bottom: 0.2rem; 
        border-top: 1px solid #e2e8f0; 
    } 
    /* Button styling for hover and active states */
    .stButton>button { /* Default button state */
        border-color: #5B4A14 !important; 
        color: #5B4A14 !important; 
        background-color: #FFFFFF !important; 
    }
    .stButton>button:hover {
        border-color: #5B4A14 !important;
        color: white !important; 
        background-color: #5B4A14 !important; 
    }
    .stButton>button:active, .stButton>button:focus, .stButton>button:focus-visible { 
        border-color: #5B4A14 !important;
        color: white !important; 
        background-color: #4A3D10 !important; 
        box-shadow: 0 0 0 0.2rem rgba(91, 74, 20, 0.5) !important; 
        outline: 2px solid #5B4A14 !important; 
        outline-offset: 2px !important;
    }
    small { 
        color: #718096; 
        font-size:0.9em; 
        display:block; 
        line-height:1.1; 
    } 
    /* Styling for selected items in st.multiselect for base colors */
    div[data-testid="stMultiSelect"] div[data-baseweb="tag"][aria-selected="true"] {
        background-color: #5B4A14 !important; 
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="tag"][aria-selected="true"] > div { 
        color: white !important;
        font-size: 0.85em !important;
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="tag"][aria-selected="true"] span[role="button"] { 
        color: white !important;
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
