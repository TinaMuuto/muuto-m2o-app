import streamlit as st
import pandas as pd
import io
import os

# --- Configuration & Constants ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Updated to use specific CSV file names provided by the user
RAW_DATA_CSV_PATH = os.path.join(BASE_DIR, "raw-data.xlsx - APP.csv")
PRICE_MATRIX_WHOLESALE_CSV_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx - Price matrix wholesale.csv")
PRICE_MATRIX_RETAIL_CSV_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx - Price matrix retail.csv")
MASTERDATA_TEMPLATE_CSV_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx - masterdata.csv")

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
if 'checkbox_selected_items' not in st.session_state: st.session_state.checkbox_selected_items = {}
if 'items_requiring_base_choice_ui' not in st.session_state: st.session_state.items_requiring_base_choice_ui = []
if 'user_chosen_base_colors_for_generic_items' not in st.session_state: st.session_state.user_chosen_base_colors_for_generic_items = {}
if 'final_items_for_download' not in st.session_state: st.session_state.final_items_for_download = []

# --- Load Data Directly from CSV files ---
files_loaded_successfully = True
if st.session_state.raw_df is None:
    if os.path.exists(RAW_DATA_CSV_PATH):
        try:
            st.session_state.raw_df = pd.read_csv(RAW_DATA_CSV_PATH)
            required_initial_cols = ['Product Type', 'Product Model', 'Sofa Direction', 'Base Color', 'Product Family', 'Item No', 'Article No', 'Image URL swatch', 'Upholstery Type', 'Upholstery Color'] # Added more essential columns
            missing_initial = [col for col in required_initial_cols if col not in st.session_state.raw_df.columns]
            if missing_initial:
                st.error(f"Nødvendige kolonner mangler i '{os.path.basename(RAW_DATA_CSV_PATH)}': {missing_initial}. Kan ikke fortsætte indlæsning.")
                files_loaded_successfully = False
            else:
                st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
                st.session_state.raw_df['Base Color Cleaned'] = st.session_state.raw_df['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)
        except Exception as e:
            st.error(f"Error loading Raw Data from CSV '{os.path.basename(RAW_DATA_CSV_PATH)}': {e}"); files_loaded_successfully = False
    else: st.error(f"Raw Data CSV file not found: {RAW_DATA_CSV_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.wholesale_prices_df is None :
    if os.path.exists(PRICE_MATRIX_WHOLESALE_CSV_PATH):
        try:
            st.session_state.wholesale_prices_df = pd.read_csv(PRICE_MATRIX_WHOLESALE_CSV_PATH)
        except Exception as e: st.error(f"Error loading Wholesale Price Matrix from CSV '{os.path.basename(PRICE_MATRIX_WHOLESALE_CSV_PATH)}': {e}"); files_loaded_successfully = False
    else: st.error(f"Wholesale Price Matrix CSV file not found: {PRICE_MATRIX_WHOLESALE_CSV_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.retail_prices_df is None :
    if os.path.exists(PRICE_MATRIX_RETAIL_CSV_PATH):
        try:
            st.session_state.retail_prices_df = pd.read_csv(PRICE_MATRIX_RETAIL_CSV_PATH)
        except Exception as e: st.error(f"Error loading Retail Price Matrix from CSV '{os.path.basename(PRICE_MATRIX_RETAIL_CSV_PATH)}': {e}"); files_loaded_successfully = False
    else: st.error(f"Retail Price Matrix CSV file not found: {PRICE_MATRIX_RETAIL_CSV_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.template_cols is None:
    if os.path.exists(MASTERDATA_TEMPLATE_CSV_PATH):
        try:
            template_df = pd.read_csv(MASTERDATA_TEMPLATE_CSV_PATH)
            st.session_state.template_cols = template_df.columns.tolist()
            if "Wholesale price" not in st.session_state.template_cols: st.session_state.template_cols.append("Wholesale price")
            if "Retail price" not in st.session_state.template_cols: st.session_state.template_cols.append("Retail price")
        except Exception as e: st.error(f"Error loading Masterdata Template from CSV '{os.path.basename(MASTERDATA_TEMPLATE_CSV_PATH)}': {e}"); files_loaded_successfully = False
    else: st.error(f"Masterdata Template CSV file not found: {MASTERDATA_TEMPLATE_CSV_PATH}"); files_loaded_successfully = False


# --- Main Application Area ---
if files_loaded_successfully and all(df is not None for df in [st.session_state.raw_df, st.session_state.wholesale_prices_df, st.session_state.retail_prices_df]) and st.session_state.template_cols:
    
    st.header("Trin 1.a: Vælg Tekstil-kombinationer")
    search_query = st.text_input("Søg på Produkt Familie eller Produkt Navn:", value=st.session_state.search_query_session, key="search_field")
    st.session_state.search_query_session = search_query
    search_query_lower = search_query.lower().strip()

    df_for_display = st.session_state.raw_df.copy()
    if search_query_lower:
        # Ensure columns exist before trying to apply lambda
        if 'Product Family' in df_for_display.columns and 'Product Display Name' in df_for_display.columns:
            df_for_display = df_for_display[
                df_for_display.apply(lambda row: search_query_lower in str(row['Product Family']).lower() or \
                                               search_query_lower in str(row['Product Display Name']).lower(), axis=1)
            ]
        else:
            st.warning("Kolonnerne 'Product Family' eller 'Product Display Name' mangler for at kunne søge.")

        if df_for_display.empty and ('Product Family' in st.session_state.raw_df.columns): # Check if raw_df had the column
             st.info(f"Ingen produkter fundet for søgningen: '{search_query}'")


    available_families_in_view = [DEFAULT_NO_SELECTION] + sorted(df_for_display['Product Family'].dropna().unique()) if 'Product Family' in df_for_display.columns else [DEFAULT_NO_SELECTION]
    if st.session_state.selected_family_session not in available_families_in_view: st.session_state.selected_family_session = DEFAULT_NO_SELECTION
    
    selected_family_idx = 0
    if st.session_state.selected_family_session in available_families_in_view:
        selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)
        
    selected_family = st.selectbox("Vælg Produkt Familie (filtrerer nuværende visning):", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
    st.session_state.selected_family_session = selected_family

    df_to_iterate_products = df_for_display.copy()
    if selected_family and selected_family != DEFAULT_NO_SELECTION and 'Product Family' in df_to_iterate_products.columns:
        df_to_iterate_products = df_to_iterate_products[df_to_iterate_products['Product Family'] == selected_family]
        families_to_render = [selected_family] if not df_to_iterate_products.empty else []
    elif 'Product Family' in df_to_iterate_products.columns:
        families_to_render = sorted(df_to_iterate_products['Product Family'].dropna().unique())
    else:
        families_to_render = []
        if files_loaded_successfully: # Only show if files were meant to be loaded
             st.warning("Kolonnen 'Product Family' mangler i rådata.")


    # Callback for checkbox changes in Step 1.a
    def handle_matrix_checkbox_toggle(item_data_from_matrix, checkbox_key_matrix):
        is_checked_now = st.session_state[checkbox_key_matrix]
        item_key = item_data_from_matrix['key'] 

        if is_checked_now:
            if item_key not in st.session_state.checkbox_selected_items:
                st.session_state.checkbox_selected_items[item_key] = item_data_from_matrix
                st.toast(f"Valgt: {item_key}", icon="➕")
        else:
            if item_key in st.session_state.checkbox_selected_items:
                del st.session_state.checkbox_selected_items[item_key]
                if item_key in st.session_state.user_chosen_base_colors_for_generic_items:
                    del st.session_state.user_chosen_base_colors_for_generic_items[item_key]
                st.toast(f"Fravalgt: {item_key}", icon="➖")

    if not df_to_iterate_products.empty and families_to_render and 'Product Display Name' in df_to_iterate_products.columns:
        for family_name_iter in families_to_render:
            if not (selected_family and selected_family != DEFAULT_NO_SELECTION) and len(families_to_render) > 1:
                 st.header(f"Produkt Familie: {family_name_iter}")
            
            current_family_df = df_to_iterate_products[df_to_iterate_products['Product Family'] == family_name_iter]
            products_in_current_family = sorted(current_family_df['Product Display Name'].dropna().unique())

            for product_name_disp in products_in_current_family:
                st.subheader(f"Produkt: {product_name_disp}")
                product_items_all_df = current_family_df[current_family_df['Product Display Name'] == product_name_disp]

                expected_cols_for_agg = ['Item No', 'Article No', 'Image URL swatch', 'Base Color Cleaned', 'Upholstery Type', 'Upholstery Color']
                actual_cols_in_product_items_df = product_items_all_df.columns.tolist()
                missing_cols_for_agg = [col for col in expected_cols_for_agg if col not in actual_cols_in_product_items_df]

                if missing_cols_for_agg:
                    st.error(f"FEJL for produkt '{product_name_disp}': Nødvendige kolonner mangler før gruppering: {missing_cols_for_agg}.")
                    st.info(f"Tilgængelige kolonner i data for dette produkt: {actual_cols_in_product_items_df}")
                    st.warning(f"Dette produkt springes over. Kontroller venligst kolonnenavnene i '{os.path.basename(RAW_DATA_CSV_PATH)}'.")
                    st.markdown("---")
                    continue 
                
                unique_textile_configs = product_items_all_df.groupby(
                    ['Upholstery Type', 'Upholstery Color'], dropna=False 
                ).agg(
                    display_item_no=('Item No', 'first'),
                    display_article_no=('Article No', 'first'),
                    display_swatch_url=('Image URL swatch', 'first'),
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
                    
                    matrix_row_key = f"{family_name_iter}_{product_name_disp}_{uph_type}_{uph_color}".replace(" ", "_").replace("/", "_").replace("(", "").replace(")", "") # Sanitize key
                    
                    base_color_display_in_matrix = "N/A"
                    item_data_for_matrix_cb = {}

                    if num_bases > 1:
                        base_color_display_in_matrix = "Flere Valg"
                        item_data_for_matrix_cb = {
                            'key': matrix_row_key, 'family': family_name_iter, 'product': product_name_disp,
                            'upholstery_type': uph_type, 'upholstery_color': uph_color,
                            'has_multiple_base_options': True,
                            'available_base_colors': available_bases_for_this_textile, 
                            'display_item_no': textile_row['display_item_no'], 
                            'display_article_no': textile_row['display_article_no'], 
                            'display_swatch_url': textile_row['display_swatch_url']
                        }
                    else: 
                        specific_item_df = product_items_all_df[
                            (product_items_all_df['Upholstery Type'].fillna("N/A") == pd.Series(uph_type).fillna("N/A").iloc[0]) &
                            (product_items_all_df['Upholstery Color'].astype(str).fillna("N/A") == pd.Series(uph_color).astype(str).fillna("N/A").iloc[0])
                        ]
                        if num_bases == 1:
                            base_color_display_in_matrix = available_bases_for_this_textile[0]
                            specific_item_df = specific_item_df[specific_item_df['Base Color Cleaned'].fillna("N/A") == base_color_display_in_matrix]
                        else: 
                             specific_item_df = specific_item_df[specific_item_df['Base Color Cleaned'].isna()]


                        if not specific_item_df.empty:
                            actual_item_row = specific_item_df.iloc[0]
                            item_data_for_matrix_cb = {
                                'key': actual_item_row['Item No'], 
                                'item_no': actual_item_row['Item No'], 'article_no': actual_item_row['Article No'],
                                'family': family_name_iter, 'product': product_name_disp,
                                'upholstery_type': uph_type, 'upholstery_color': uph_color,
                                'base_color': base_color_display_in_matrix, 
                                'has_multiple_base_options': False,
                                'description': f"{family_name_iter} / {product_name_disp} / {uph_type} / {uph_color} / Ben: {base_color_display_in_matrix}",
                                'display_swatch_url': actual_item_row['Image URL swatch']
                            }
                        else:
                            st.caption(f"Data-uoverensstemmelse for {uph_type} / {uph_color}. Springes over.")
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
                st.markdown("---") 
    elif not files_loaded_successfully:
        pass # Errors already shown during loading
    elif 'Product Family' not in st.session_state.raw_df.columns : # Check if raw_df was loaded but missing key column
        st.error("Nødvendig kolonne 'Product Family' mangler i rådata. Kan ikke vise produkter.")
    elif search_query_lower and df_for_display.empty:
        pass # Message "Ingen produkter fundet for søgningen" is already shown
    elif not search_query_lower and not (selected_family and selected_family != DEFAULT_NO_SELECTION):
         st.info("Indtast et søgeord eller vælg en Produkt Familie for at se tilgængelige varer.")


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
    if st.session_state.checkbox_selected_items: 
        if st.button("Bekræft Valg og Opdater Endelig Liste", key="confirm_all_selections"):
            st.session_state.final_items_for_download = [] 

            for key, initial_item_data in st.session_state.checkbox_selected_items.items():
                if not initial_item_data.get('has_multiple_base_options'):
                    st.session_state.final_items_for_download.append({
                        "description": initial_item_data.get('description', f"{initial_item_data['family']} / {initial_item_data['product']} / {initial_item_data['upholstery_type']} / {initial_item_data['upholstery_color']} / Ben: {initial_item_data['base_color']}"),
                        "item_no": initial_item_data['item_no'],
                        "article_no": initial_item_data['article_no'],
                    })
                else:
                    chosen_base_colors_for_this_generic = st.session_state.user_chosen_base_colors_for_generic_items.get(key, [])
                    if not chosen_base_colors_for_this_generic:
                        st.warning(f"Ingen benfarver valgt for {initial_item_data['product']} ({initial_item_data['upholstery_type']}/{initial_item_data['upholstery_color']}). Springes over.")
                        continue

                    for chosen_bc in chosen_base_colors_for_this_generic:
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
                            st.error(f"Kunne ikke finde specifikt varenummer for {initial_item_data['product']} ({initial_item_data['upholstery_type']}/{initial_item_data['upholstery_color']}) med benfarve '{chosen_bc}'. Tjek data.")
            
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
                item_to_remove_no = st.session_state.final_items_for_download.pop(i)['item_no']
                st.toast(f"Fjernet fra endelig liste: {item_to_remove_no}", icon="🗑️")
                st.rerun() 
        
        st.header("Trin 3: Vælg Valuta og Generer Fil")
        try:
            # Ensure wholesale_prices_df has at least one column for Article No
            if not st.session_state.wholesale_prices_df.empty:
                article_no_col_name_ws = st.session_state.wholesale_prices_df.columns[0]
                currency_options = [col for col in st.session_state.wholesale_prices_df.columns if str(col).lower() != str(article_no_col_name_ws).lower()]
            else:
                currency_options = []
                
            selected_currency = st.selectbox("Vælg Valuta:", options=currency_options, key="currency_selector") if currency_options else None
            if not currency_options and not st.session_state.wholesale_prices_df.empty: 
                st.error("Ingen valuta kolonner fundet i Pris Matrix (udover Artikel Nr. kolonnen).")
            elif st.session_state.wholesale_prices_df.empty:
                 st.error("Pris Matrix (wholesale) er tom. Kan ikke bestemme valuta.")

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
                    
                    # Ensure wholesale_prices_df and retail_prices_df are not empty before iloc
                    if not st.session_state.wholesale_prices_df.empty:
                        ws_price_row_df = st.session_state.wholesale_prices_df[st.session_state.wholesale_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                        if not ws_price_row_df.empty and selected_currency in ws_price_row_df.columns:
                            output_row_dict["Wholesale price"] = ws_price_row_df.iloc[0][selected_currency] if pd.notna(ws_price_row_df.iloc[0][selected_currency]) else "N/A"
                        else: output_row_dict["Wholesale price"] = "Pris Ikke Fundet"
                    else: output_row_dict["Wholesale price"] = "Pris Matrix (W) Tom"

                    if not st.session_state.retail_prices_df.empty:
                        rt_price_row_df = st.session_state.retail_prices_df[st.session_state.retail_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                        if not rt_price_row_df.empty and selected_currency in rt_price_row_df.columns:
                            output_row_dict["Retail price"] = rt_price_row_df.iloc[0][selected_currency] if pd.notna(rt_price_row_df.iloc[0][selected_currency]) else "N/A"
                        else: output_row_dict["Retail price"] = "Pris Ikke Fundet"
                    else: output_row_dict["Retail price"] = "Pris Matrix (R) Tom"
                    
                    output_data.append(output_row_dict)
                else: st.warning(f"Data for Vare Nr: {item_no_to_find} ikke fundet. Springes over.")
            if output_data:
                output_df = pd.DataFrame(output_data, columns=master_template_columns_ordered)
                output_excel_buffer = io.BytesIO()
                with pd.ExcelWriter(output_excel_buffer, engine='xlsxwriter') as writer: output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
                output_excel_buffer.seek(0)
                st.download_button(label="📥 Download Masterdata Excel Fil", data=output_excel_buffer, file_name=f"masterdata_output_{selected_currency.replace(' ', '_').replace('.', '')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else: st.warning("Ingen data at generere.")
        elif not selected_currency and st.session_state.final_items_for_download : st.warning("Vælg venligst en valuta.")
    elif not st.session_state.checkbox_selected_items : 
         st.info("Foretag valg i Trin 1.a for at fortsætte.")


else: 
    st.error("En eller flere datafiler kunne ikke indlæses korrekt, eller nødvendige kolonner mangler. Kontroller stier, filformater og kolonnenavne i dine CSV-filer.")
    if not os.path.exists(RAW_DATA_CSV_PATH): st.info(f"Mangler: {RAW_DATA_CSV_PATH}")
    if not os.path.exists(PRICE_MATRIX_WHOLESALE_CSV_PATH): st.info(f"Mangler: {PRICE_MATRIX_WHOLESALE_CSV_PATH}")
    if not os.path.exists(PRICE_MATRIX_RETAIL_CSV_PATH): st.info(f"Mangler: {PRICE_MATRIX_RETAIL_CSV_PATH}")
    if not os.path.exists(MASTERDATA_TEMPLATE_CSV_PATH): st.info(f"Mangler: {MASTERDATA_TEMPLATE_CSV_PATH}")


# --- Styling (Optional) ---
st.markdown("""
<style>
    h1 { /* App Title */ color: #333; }
    h2 { /* Step Headers */ color: #1E40AF; border-bottom: 2px solid #BFDBFE; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    h3 { /* Product Subheaders */ background-color: #e8f0fe; padding: 10px; border-radius: 5px; margin-top: 20px; margin-bottom: 10px; font-size: 1.15em; }
    div[data-testid="stCaptionContainer"] > div > p { font-weight: bold; font-size: 0.85em !important; color: #4A5568 !important; padding-bottom: 3px; }
    div[data-testid="stImage"] img { max-height: 45px; border: 1px solid #e2e8f0; border-radius: 3px; padding: 2px; margin: auto; display: block; }
    div.stCheckbox, div[data-testid="stMarkdownContainer"] { font-size: 0.9em; display: flex; align-items: center; height: 50px; }
    div.stCheckbox { justify-content: center; }
    hr { margin-top: 0.2rem; margin-bottom: 0.2rem; border-top: 1px solid #e2e8f0; }
    .stButton>button { font-size: 0.9em; padding: 0.3em 0.7em; }
    small { color: #718096; } 
</style>
""", unsafe_allow_html=True)
