import streamlit as st
import pandas as pd
import io
import os

# --- Configuration & Constants ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
PRICE_MATRIX_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx")
MASTERDATA_TEMPLATE_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")

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
st.set_page_config(layout="wide")
st.title("Product Configurator & Masterdata Generator")

# --- Initialize session state variables ---
if 'raw_df' not in st.session_state: st.session_state.raw_df = None
if 'wholesale_prices_df' not in st.session_state: st.session_state.wholesale_prices_df = None
if 'retail_prices_df' not in st.session_state: st.session_state.retail_prices_df = None
if 'template_cols' not in st.session_state: st.session_state.template_cols = None
if 'search_query_session' not in st.session_state: st.session_state.search_query_session = ""
if 'selected_family_session' not in st.session_state: st.session_state.selected_family_session = None
# New session state variables for the refined selection process
if 'checkbox_selected_items' not in st.session_state: st.session_state.checkbox_selected_items = {} # Stores {key: item_data} from checkbox matrix
if 'items_requiring_base_choice_ui' not in st.session_state: st.session_state.items_requiring_base_choice_ui = [] # List of items for Step 1.b
if 'user_chosen_base_colors_for_generic_items' not in st.session_state: st.session_state.user_chosen_base_colors_for_generic_items = {} # Stores {generic_item_key: [selected_base_colors]}
if 'final_items_for_download' not in st.session_state: st.session_state.final_items_for_download = [] # Final list of specific items

# --- Load Data Directly ---
files_loaded_successfully = True
if st.session_state.raw_df is None:
    if os.path.exists(RAW_DATA_PATH):
        try:
            st.session_state.raw_df = pd.read_excel(RAW_DATA_PATH, sheet_name=RAW_DATA_APP_SHEET)
            st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
            # Pre-calculate unique base colors for combinations to speed up matrix display
            st.session_state.raw_df['Base Color Cleaned'] = st.session_state.raw_df['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)

        except Exception as e:
            st.error(f"Error loading Raw Data: {e}"); files_loaded_successfully = False
    else: st.error(f"Raw Data file not found: {RAW_DATA_PATH}"); files_loaded_successfully = False
# Similar loading for other files (condensed for brevity in this thought block)
if st.session_state.wholesale_prices_df is None:
    if os.path.exists(PRICE_MATRIX_PATH):
        try:
            st.session_state.wholesale_prices_df = pd.read_excel(PRICE_MATRIX_PATH, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
            st.session_state.retail_prices_df = pd.read_excel(PRICE_MATRIX_PATH, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
        except Exception as e: st.error(f"Error loading Price Matrix: {e}"); files_loaded_successfully = False
    else: st.error(f"Price Matrix file not found: {PRICE_MATRIX_PATH}"); files_loaded_successfully = False
if st.session_state.template_cols is None:
    if os.path.exists(MASTERDATA_TEMPLATE_PATH):
        try:
            template_df = pd.read_excel(MASTERDATA_TEMPLATE_PATH)
            st.session_state.template_cols = template_df.columns.tolist()
            if "Wholesale price" not in st.session_state.template_cols: st.session_state.template_cols.append("Wholesale price")
            if "Retail price" not in st.session_state.template_cols: st.session_state.template_cols.append("Retail price")
        except Exception as e: st.error(f"Error loading Masterdata Template: {e}"); files_loaded_successfully = False
    else: st.error(f"Masterdata Template file not found: {MASTERDATA_TEMPLATE_PATH}"); files_loaded_successfully = False


# --- Main Application Area ---
if files_loaded_successfully and all(df is not None for df in [st.session_state.raw_df, st.session_state.wholesale_prices_df, st.session_state.retail_prices_df]) and st.session_state.template_cols:
    
    st.header("Trin 1.a: Vælg Tekstil-kombinationer")
    search_query = st.text_input("Søg på Produkt Familie eller Produkt Navn:", value=st.session_state.search_query_session, key="search_field")
    st.session_state.search_query_session = search_query
    search_query_lower = search_query.lower().strip()

    df_for_display = st.session_state.raw_df.copy()
    if search_query_lower:
        df_for_display = df_for_display[
            df_for_display.apply(lambda row: search_query_lower in str(row['Product Family']).lower() or \
                                           search_query_lower in str(row['Product Display Name']).lower(), axis=1)
        ]
        if df_for_display.empty: st.info(f"Ingen produkter fundet for søgningen: '{search_query}'")

    available_families_in_view = [DEFAULT_NO_SELECTION] + sorted(df_for_display['Product Family'].dropna().unique())
    if st.session_state.selected_family_session not in available_families_in_view: st.session_state.selected_family_session = DEFAULT_NO_SELECTION
    selected_family = st.selectbox("Vælg Produkt Familie (filtrerer nuværende visning):", options=available_families_in_view, index=available_families_in_view.index(st.session_state.selected_family_session), key="family_selector_main")
    st.session_state.selected_family_session = selected_family

    df_to_iterate_products = df_for_display.copy()
    if selected_family and selected_family != DEFAULT_NO_SELECTION:
        df_to_iterate_products = df_to_iterate_products[df_to_iterate_products['Product Family'] == selected_family]
        families_to_render = [selected_family] if not df_to_iterate_products.empty else []
    else:
        families_to_render = sorted(df_to_iterate_products['Product Family'].dropna().unique())

    # Callback for checkbox changes in Step 1.a
    def handle_matrix_checkbox_toggle(item_data_from_matrix, checkbox_key_matrix):
        is_checked_now = st.session_state[checkbox_key_matrix]
        item_key = item_data_from_matrix['key'] # Unique key for this matrix row

        if is_checked_now:
            if item_key not in st.session_state.checkbox_selected_items:
                st.session_state.checkbox_selected_items[item_key] = item_data_from_matrix
                st.toast(f"Valgt: {item_key}", icon="➕")
        else:
            if item_key in st.session_state.checkbox_selected_items:
                del st.session_state.checkbox_selected_items[item_key]
                # Also remove any dependent base color choices if this generic item is deselected
                if item_key in st.session_state.user_chosen_base_colors_for_generic_items:
                    del st.session_state.user_chosen_base_colors_for_generic_items[item_key]
                st.toast(f"Fravalgt: {item_key}", icon="➖")
        # After any checkbox change, re-evaluate what needs base color selection
        # This will be handled by the "Confirm and Proceed" button or a similar mechanism

    if not df_to_iterate_products.empty and families_to_render:
        for family_name_iter in families_to_render:
            if not (selected_family and selected_family != DEFAULT_NO_SELECTION) and len(families_to_render) > 1:
                 st.header(f"Produkt Familie: {family_name_iter}")
            
            current_family_df = df_to_iterate_products[df_to_iterate_products['Product Family'] == family_name_iter]
            products_in_current_family = sorted(current_family_df['Product Display Name'].dropna().unique())

            for product_name_disp in products_in_current_family:
                st.subheader(f"Produkt: {product_name_disp}")
                product_items_all_df = current_family_df[current_family_df['Product Display Name'] == product_name_disp]

                # Group by textile combination to determine base color options
                unique_textile_configs = product_items_all_df.groupby(
                    ['Upholstery Type', 'Upholstery Color'], dropna=False # Keep NA groups if they exist
                ).agg(
                    display_item_no=('Item No', 'first'),
                    display_article_no=('Article No', 'first'),
                    display_swatch_url=('Image URL swatch', 'first'),
                    # Get unique, non-NA base colors
                    _available_base_colors_internal=('Base Color Cleaned', lambda x: list(x.dropna().unique()))
                ).reset_index()
                
                unique_textile_configs['num_base_options'] = unique_textile_configs['_available_base_colors_internal'].apply(len)

                if unique_textile_configs.empty:
                    st.caption("Ingen tekstil konfigurationer fundet.")
                    st.markdown("---")
                    continue

                header_cols = st.columns([0.5, 0.7, 1.5, 1.2, 1.2, 1.7]) 
                with header_cols[0]: st.caption("Vælg")
                with header_cols[1]: st.caption("Swatch")
                with header_cols[2]: st.caption("Tekstil")
                with header_cols[3]: st.caption("Farve")
                with header_cols[4]: st.caption("Ben")
                with header_cols[5]: st.caption("Detaljer")

                for _, textile_row in unique_textile_configs.iterrows():
                    uph_type = textile_row['Upholstery Type']
                    uph_color = str(textile_row['Upholstery Color'])
                    num_bases = textile_row['num_base_options']
                    available_bases_for_this_textile = textile_row['_available_base_colors_internal']
                    
                    # Matrix key: identifies the row in the display matrix (generic textile combo)
                    matrix_row_key = f"{family_name_iter}_{product_name_disp}_{uph_type}_{uph_color}".replace(" ", "_")
                    
                    base_color_display_in_matrix = "N/A"
                    item_data_for_matrix_cb = {}

                    if num_bases > 1:
                        base_color_display_in_matrix = "Flere Valg"
                        item_data_for_matrix_cb = {
                            'key': matrix_row_key, 'family': family_name_iter, 'product': product_name_disp,
                            'upholstery_type': uph_type, 'upholstery_color': uph_color,
                            'has_multiple_base_options': True,
                            'available_base_colors': available_bases_for_this_textile, # Store for Step 1.b
                            'display_item_no': textile_row['display_item_no'], # For display only
                            'display_article_no': textile_row['display_article_no'], # For display only
                            'display_swatch_url': textile_row['display_swatch_url']
                        }
                    else: # 0 or 1 base option
                        # Find the specific item for this single/no base option
                        specific_item_df = product_items_all_df[
                            (product_items_all_df['Upholstery Type'].fillna("N/A") == pd.Series(uph_type).fillna("N/A").iloc[0]) &
                            (product_items_all_df['Upholstery Color'].astype(str).fillna("N/A") == pd.Series(uph_color).astype(str).fillna("N/A").iloc[0])
                        ]
                        if num_bases == 1:
                            base_color_display_in_matrix = available_bases_for_this_textile[0]
                            specific_item_df = specific_item_df[specific_item_df['Base Color Cleaned'].fillna("N/A") == base_color_display_in_matrix]
                        else: # num_bases == 0 (only N/A or empty base colors)
                             specific_item_df = specific_item_df[specific_item_df['Base Color Cleaned'].isna()]


                        if not specific_item_df.empty:
                            actual_item_row = specific_item_df.iloc[0]
                            item_data_for_matrix_cb = {
                                'key': actual_item_row['Item No'], # Use actual Item No as key
                                'item_no': actual_item_row['Item No'], 'article_no': actual_item_row['Article No'],
                                'family': family_name_iter, 'product': product_name_disp,
                                'upholstery_type': uph_type, 'upholstery_color': uph_color,
                                'base_color': base_color_display_in_matrix, # The single or N/A base color
                                'has_multiple_base_options': False,
                                'description': f"{family_name_iter} / {product_name_disp} / {uph_type} / {uph_color} / Ben: {base_color_display_in_matrix}",
                                'display_swatch_url': actual_item_row['Image URL swatch']
                            }
                        else:
                            # This case should ideally not be hit if data is consistent
                            st.caption(f"Skipping inconsistent data for {uph_type} / {uph_color}")
                            continue
                    
                    is_matrix_row_selected = item_data_for_matrix_cb.get('key') in st.session_state.checkbox_selected_items
                    
                    item_detail_cols = st.columns([0.5, 0.7, 1.5, 1.2, 1.2, 1.7])
                    with item_detail_cols[0]:
                        st.checkbox(" ", value=is_matrix_row_selected, key=f"mcb_{item_data_for_matrix_cb.get('key')}",
                                    on_change=handle_matrix_checkbox_toggle, args=(item_data_for_matrix_cb, f"mcb_{item_data_for_matrix_cb.get('key')}"))
                    with item_detail_cols[1]:
                        sw_url = item_data_for_matrix_cb.get('display_swatch_url', textile_row['display_swatch_url'])
                        if pd.notna(sw_url): st.image(sw_url, width=50)
                        else: st.markdown("<div style='width:50px; height:50px; border:1px solid #ddd; display:flex; align-items:center; justify-content:center; font-size:0.7em;'>No Swatch</div>", unsafe_allow_html=True)
                    with item_detail_cols[2]: st.markdown(f"<div style='font-size:0.9em;'>{uph_type}</div>", unsafe_allow_html=True)
                    with item_detail_cols[3]: st.markdown(f"<div style='font-size:0.9em;'>{uph_color}</div>", unsafe_allow_html=True)
                    with item_detail_cols[4]: st.markdown(f"<div style='font-size:0.9em;'>{base_color_display_in_matrix}</div>", unsafe_allow_html=True)
                    with item_detail_cols[5]: 
                        if item_data_for_matrix_cb.get('has_multiple_base_options', False):
                            st.markdown(f"<div style='font-size:0.9em;'><small><i>(Vælg ben i næste trin)</i></small></div>", unsafe_allow_html=True)
                        else:
                            st.markdown(f"<div style='font-size:0.9em;'><small><i>Vare: {item_data_for_matrix_cb.get('item_no', 'N/A')}<br>Artikel: {item_data_for_matrix_cb.get('article_no', 'N/A')}</i></small></div>", unsafe_allow_html=True)
                    st.markdown("---")
                st.markdown("---") # Separator after each product group

    # --- Step 1.b: Select Base Colors for applicable items ---
    st.session_state.items_requiring_base_choice_ui = [
        item_data for key, item_data in st.session_state.checkbox_selected_items.items() if item_data.get('has_multiple_base_options')
    ]

    if st.session_state.items_requiring_base_choice_ui:
        st.header("Trin 1.b: Vælg Benfarver for Valgte Tekstiler")
        for generic_item_data in st.session_state.items_requiring_base_choice_ui:
            item_key = generic_item_data['key']
            st.markdown(f"**Produkt:** {generic_item_data['product']} - **Tekstil:** {generic_item_data['upholstery_type']} ({generic_item_data['upholstery_color']})")
            
            available_bc_options = generic_item_data.get('available_base_colors', [])
            if not available_bc_options:
                st.warning("Ingen benfarve-valgmuligheder fundet for denne kombination, selvom det var forventet.")
                continue

            current_bc_selection = st.session_state.user_chosen_base_colors_for_generic_items.get(item_key, [])
            
            chosen_bases = st.multiselect(
                f"Vælg benfarve(r) for {generic_item_data['product']}:",
                options=available_bc_options,
                default=current_bc_selection,
                key=f"ms_{item_key}"
            )
            st.session_state.user_chosen_base_colors_for_generic_items[item_key] = chosen_bases
            st.markdown("---")

    # --- Button to finalize selections and process for download list ---
    if st.session_state.checkbox_selected_items: # Show button only if some initial selections are made
        if st.button("Bekræft Valg og Opdater Endelig Liste", key="confirm_all_selections"):
            st.session_state.final_items_for_download = [] # Reset before populating

            for key, initial_item_data in st.session_state.checkbox_selected_items.items():
                if not initial_item_data.get('has_multiple_base_options'):
                    # This is a specific item, add directly
                    st.session_state.final_items_for_download.append({
                        "description": initial_item_data.get('description', f"{initial_item_data['family']} / {initial_item_data['product']} / {initial_item_data['upholstery_type']} / {initial_item_data['upholstery_color']} / Ben: {initial_item_data['base_color']}"),
                        "item_no": initial_item_data['item_no'],
                        "article_no": initial_item_data['article_no'],
                        # Add other fields if needed for display or processing
                    })
                else:
                    # This is a generic item, check user_chosen_base_colors_for_generic_items
                    chosen_base_colors_for_this_generic = st.session_state.user_chosen_base_colors_for_generic_items.get(key, [])
                    if not chosen_base_colors_for_this_generic:
                        st.warning(f"Ingen benfarver valgt for {initial_item_data['product']} ({initial_item_data['upholstery_type']}/{initial_item_data['upholstery_color']}). Springes over.")
                        continue

                    for chosen_bc in chosen_base_colors_for_this_generic:
                        # Find the specific Item No/Article No from raw_df
                        final_item_row_df = st.session_state.raw_df[
                            (st.session_state.raw_df['Product Family'] == initial_item_data['family']) &
                            (st.session_state.raw_df['Product Display Name'] == initial_item_data['product']) &
                            (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == pd.Series(initial_item_data['upholstery_type']).fillna("N/A").iloc[0]) &
                            (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == pd.Series(initial_item_data['upholstery_color']).astype(str).fillna("N/A").iloc[0]) &
                            (st.session_state.raw_df['Base Color Cleaned'].fillna("N/A") == chosen_bc)
                        ]
                        if not final_item_row_df.empty:
                            final_item_concrete = final_item_row_df.iloc[0]
                            st.session_state.final_items_for_download.append({
                                "description": f"{initial_item_data['family']} / {initial_item_data['product']} / {initial_item_data['upholstery_type']} / {initial_item_data['upholstery_color']} / Ben: {chosen_bc}",
                                "item_no": final_item_concrete['Item No'],
                                "article_no": final_item_concrete['Article No'],
                            })
                        else:
                            st.error(f"Kunne ikke finde specifikt varenummer for {initial_item_data['product']} med benfarve {chosen_bc}. Tjek data.")
            
            # Remove duplicates from final list just in case
            temp_final_list = []
            seen_item_nos = set()
            for item in st.session_state.final_items_for_download:
                if item['item_no'] not in seen_item_nos:
                    temp_final_list.append(item)
                    seen_item_nos.add(item['item_no'])
            st.session_state.final_items_for_download = temp_final_list

            if st.session_state.final_items_for_download:
                st.success("Endelig liste er opdateret!")
            else:
                st.info("Ingen varer på den endelige liste efter bekræftelse.")
            st.rerun()


    # --- Display Current Selections (Final List) ---
    if st.session_state.final_items_for_download:
        st.header("Trin 2: Gennemse Endelige Valgte Kombinationer")
        for i, combo in enumerate(st.session_state.final_items_for_download):
            col1, col2 = st.columns([0.9, 0.1])
            col1.write(f"{i+1}. {combo['description']} (Vare: {combo['item_no']})")
            if col2.button(f"Fjern", key=f"final_remove_{i}_{combo['item_no']}"):
                # When removing from final list, we might need to update checkbox state if it was a simple item,
                # or clear selections in user_chosen_base_colors if it was derived.
                # For now, simple removal from final list. User can re-add via Step 1.
                item_to_remove_no = st.session_state.final_items_for_download.pop(i)['item_no']
                
                # Attempt to uncheck the original checkbox if this was a "simple" item
                if f"mcb_{item_to_remove_no}" in st.session_state.checkbox_selected_items:
                     # This logic is tricky because the key might be generic.
                     # A more robust way would be to rebuild the final list from scratch on removal.
                     # For now, just remove from final list and rerun.
                     pass

                st.toast(f"Fjernet fra endelig liste: {item_to_remove_no}", icon="🗑️")
                st.rerun() 
        
        st.header("Trin 3: Vælg Valuta og Generer Fil")
        try:
            article_no_col_name_ws = st.session_state.wholesale_prices_df.columns[0]
            currency_options = [col for col in st.session_state.wholesale_prices_df.columns if str(col).lower() != str(article_no_col_name_ws).lower()]
            selected_currency = st.selectbox("Vælg Valuta:", options=currency_options, key="currency_selector") if currency_options else None
            if not currency_options: st.error("Ingen valuta kolonner fundet i Pris Matrix.")
        except Exception as e: st.error(f"Kunne ikke bestemme valuta: {e}"); selected_currency = None

        if st.button("Generer Masterdata Fil", key="generate_file") and selected_currency:
            output_data = []
            master_template_columns_ordered = st.session_state.template_cols.copy()
            for combo_selection in st.session_state.final_items_for_download:
                item_no_to_find = combo_selection['item_no']
                article_no_to_find = combo_selection['article_no'] 
                item_data_row_series_df = st.session_state.raw_df[st.session_state.raw_df['Item No'] == item_no_to_find]
                if not item_data_row_series_df.empty:
                    item_data_row_series = item_data_row_series_df.iloc[0]
                    output_row_dict = {}
                    for col_template in master_template_columns_ordered:
                        if col_template in ["Wholesale price", "Retail price"]: continue
                        if col_template in item_data_row_series.index: output_row_dict[col_template] = item_data_row_series[col_template]
                        else: output_row_dict[col_template] = None 
                    
                    ws_price_row_df = st.session_state.wholesale_prices_df[st.session_state.wholesale_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                    if not ws_price_row_df.empty and selected_currency in ws_price_row_df.columns:
                        output_row_dict["Wholesale price"] = ws_price_row_df.iloc[0][selected_currency] if pd.notna(ws_price_row_df.iloc[0][selected_currency]) else "N/A"
                    else: output_row_dict["Wholesale price"] = "Pris Ikke Fundet"
                    
                    rt_price_row_df = st.session_state.retail_prices_df[st.session_state.retail_prices_df.iloc[:, 0].astype(str) == str(article
