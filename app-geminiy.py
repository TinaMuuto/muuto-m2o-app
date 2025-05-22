import streamlit as st
import pandas as pd
import io
import os

# --- Configuration & Constants ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_XLSX_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
PRICE_MATRIX_XLSX_PATH = os.path.join(BASE_DIR, "price-matrix_EUROPE.xlsx")
MASTERDATA_TEMPLATE_XLSX_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")

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
# For the new matrix selection
if 'matrix_selected_generic_items' not in st.session_state: st.session_state.matrix_selected_generic_items = {} # Stores {unique_key: data}
if 'user_chosen_base_colors_for_items' not in st.session_state: st.session_state.user_chosen_base_colors_for_items = {} # {generic_item_key: [selected_bases]}
if 'final_items_for_download' not in st.session_state: st.session_state.final_items_for_download = []


# --- Load Data Directly from XLSX files ---
files_loaded_successfully = True
# Sidebar diagnostics can be commented out if not needed
# st.sidebar.subheader("Diagnostik for Filindlæsning:")
# st.sidebar.write(f"Forventet script-mappe (BASE_DIR): `{BASE_DIR}`")
# try:
#     st.sidebar.write(f"Filer fundet i script-mappen:")
#     found_files_in_basedir = os.listdir(BASE_DIR)
#     if found_files_in_basedir:
#         for f_name in found_files_in_basedir: st.sidebar.code(f_name)
#     else: st.sidebar.warning("Ingen filer fundet i script-mappen.")
# except Exception as e_listdir: st.sidebar.error(f"Fejl ved listning af filer: {e_listdir}")

if st.session_state.raw_df is None:
    if os.path.exists(RAW_DATA_XLSX_PATH):
        try:
            st.session_state.raw_df = pd.read_excel(RAW_DATA_XLSX_PATH, sheet_name=RAW_DATA_APP_SHEET)
            required_cols = ['Product Type', 'Product Model', 'Sofa Direction', 'Base Color', 'Product Family', 'Item No', 'Article No', 'Image URL swatch', 'Upholstery Type', 'Upholstery Color', 'Market']
            missing = [col for col in required_cols if col not in st.session_state.raw_df.columns]
            if missing:
                st.error(f"Nødvendige kolonner mangler i '{os.path.basename(RAW_DATA_XLSX_PATH)}': {missing}.")
                files_loaded_successfully = False
            else:
                st.session_state.raw_df = st.session_state.raw_df[st.session_state.raw_df['Market'].astype(str).str.upper() != 'UK']
                st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
                st.session_state.raw_df['Base Color Cleaned'] = st.session_state.raw_df['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)
        except Exception as e: st.error(f"Fejl ved indlæsning af Rådata: {e}"); files_loaded_successfully = False
    else: st.error(f"Rådata fil ikke fundet: {RAW_DATA_XLSX_PATH}"); files_loaded_successfully = False

# Condensed loading for other files
if files_loaded_successfully and st.session_state.wholesale_prices_df is None:
    if os.path.exists(PRICE_MATRIX_XLSX_PATH):
        try: st.session_state.wholesale_prices_df = pd.read_excel(PRICE_MATRIX_XLSX_PATH, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
        except Exception as e: st.error(f"Fejl ved indlæsning af Wholesale Priser: {e}"); files_loaded_successfully = False
    else: st.error(f"Pris Matrix fil ikke fundet: {PRICE_MATRIX_XLSX_PATH}"); files_loaded_successfully = False

if files_loaded_successfully and st.session_state.retail_prices_df is None:
    if os.path.exists(PRICE_MATRIX_XLSX_PATH):
        try: st.session_state.retail_prices_df = pd.read_excel(PRICE_MATRIX_XLSX_PATH, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
        except Exception as e: st.error(f"Fejl ved indlæsning af Retail Priser: {e}"); files_loaded_successfully = False
    # No separate error if file itself not found, covered by wholesale check

if files_loaded_successfully and st.session_state.template_cols is None:
    if os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH):
        try:
            st.session_state.template_cols = pd.read_excel(MASTERDATA_TEMPLATE_XLSX_PATH).columns.tolist()
            if "Wholesale price" not in st.session_state.template_cols: st.session_state.template_cols.append("Wholesale price")
            if "Retail price" not in st.session_state.template_cols: st.session_state.template_cols.append("Retail price")
        except Exception as e: st.error(f"Fejl ved indlæsning af Skabelon: {e}"); files_loaded_successfully = False
    else: st.error(f"Skabelon fil ikke fundet: {MASTERDATA_TEMPLATE_XLSX_PATH}"); files_loaded_successfully = False

# --- Main Application Area ---
if files_loaded_successfully and all(df is not None for df in [st.session_state.raw_df, st.session_state.wholesale_prices_df, st.session_state.retail_prices_df]) and st.session_state.template_cols:
    
    st.header("Trin 1: Vælg Produkt Kombinationer (Produkt / Tekstil / Farve)")
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
    if st.session_state.selected_family_session not in available_families_in_view: 
        st.session_state.selected_family_session = DEFAULT_NO_SELECTION
    
    selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)
    selected_family = st.selectbox("Vælg Produkt Familie:", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
    st.session_state.selected_family_session = selected_family
    
    # Callback for matrix checkbox
    def handle_matrix_cb_toggle(prod_name, uph_type, uph_color, checkbox_key_matrix):
        is_checked = st.session_state[checkbox_key_matrix]
        # Unique key for this generic combination
        generic_item_key = f"{selected_family}_{prod_name}_{uph_type}_{uph_color}".replace(" ", "_").replace("/","_")

        if is_checked:
            # Find matching items to determine base color options
            matching_items = st.session_state.raw_df[
                (st.session_state.raw_df['Product Family'] == selected_family) &
                (st.session_state.raw_df['Product Display Name'] == prod_name) &
                (st.session_state.raw_df['Upholstery Type'].fillna("N/A") == uph_type) &
                (st.session_state.raw_df['Upholstery Color'].astype(str).fillna("N/A") == uph_color)
            ]
            if not matching_items.empty:
                unique_base_colors = matching_items['Base Color Cleaned'].dropna().unique().tolist()
                first_item_match = matching_items.iloc[0] # For swatch, item no if no base choice

                item_data = {
                    'key': generic_item_key,
                    'family': selected_family,
                    'product': prod_name,
                    'upholstery_type': uph_type,
                    'upholstery_color': uph_color,
                    'requires_base_choice': len(unique_base_colors) > 1,
                    'available_bases': unique_base_colors if len(unique_base_colors) > 1 else [],
                    'item_no_if_single_base': first_item_match['Item No'] if len(unique_base_colors) <= 1 else None,
                    'article_no_if_single_base': first_item_match['Article No'] if len(unique_base_colors) <= 1 else None,
                    'resolved_base_if_single': unique_base_colors[0] if len(unique_base_colors) == 1 else (pd.NA if not unique_base_colors else None)
                }
                st.session_state.matrix_selected_generic_items[generic_item_key] = item_data
                st.toast(f"Valgt: {prod_name} / {uph_type} / {uph_color}", icon="➕")
            else:
                st.warning(f"Ingen vare fundet for {prod_name} / {uph_type} / {uph_color}")
        else:
            if generic_item_key in st.session_state.matrix_selected_generic_items:
                del st.session_state.matrix_selected_generic_items[generic_item_key]
                if generic_item_key in st.session_state.user_chosen_base_colors_for_items:
                    del st.session_state.user_chosen_base_colors_for_items[generic_item_key]
                st.toast(f"Fravalgt: {prod_name} / {uph_type} / {uph_color}", icon="➖")


    if selected_family and selected_family != DEFAULT_NO_SELECTION:
        family_df = df_for_display[df_for_display['Product Family'] == selected_family]
        if not family_df.empty:
            products_in_family = sorted(family_df['Product Display Name'].dropna().unique())
            upholstery_types_in_family = sorted(family_df['Upholstery Type'].dropna().unique())

            if not products_in_family: st.info("Ingen produkter i denne familie.")
            elif not upholstery_types_in_family: st.info("Ingen tekstilfamilier for denne produktfamilie.")
            else:
                # --- Create data for table display ---
                # Header rows data
                header_upholstery_types = ["Produkt"]
                header_swatches = [" "] # Placeholder for product column
                header_color_numbers = [" "] # Placeholder for product column
                
                # Map (Upholstery Type, Upholstery Color) to a flat list of column headers for data cells
                data_column_map = [] # List of (uph_type, uph_color, swatch_url)

                for uph_type in upholstery_types_in_family:
                    colors_for_type_df = family_df[family_df['Upholstery Type'] == uph_type][['Upholstery Color', 'Image URL swatch']].drop_duplicates().sort_values(by='Upholstery Color')
                    if not colors_for_type_df.empty:
                        header_upholstery_types.extend([uph_type] + [""] * (len(colors_for_type_df) -1) ) # Span Upholstery Type
                        for _, color_row in colors_for_type_df.iterrows():
                            color_val = str(color_row['Upholstery Color'])
                            swatch_val = color_row['Image URL swatch']
                            header_swatches.append(swatch_val if pd.notna(swatch_val) else None)
                            header_color_numbers.append(color_val)
                            data_column_map.append({'uph_type': uph_type, 'uph_color': color_val, 'swatch': swatch_val})
                
                num_data_columns = len(data_column_map)
                if num_data_columns == 0:
                    st.info("Ingen tekstil/farve kombinationer at vise for denne familie.")
                else:
                    # --- Display Table Headers ---
                    # Upholstery Type Headers (Visually Merged)
                    cols_uph_type_header = st.columns([2.5] + [1] * num_data_columns) # Product col + data cols
                    current_uph_type_header = None
                    for i, col_widget in enumerate(cols_uph_type_header):
                        if i == 0:
                            with col_widget: st.caption("") # Empty for Product column
                        else:
                            map_entry = data_column_map[i-1]
                            if map_entry['uph_type'] != current_uph_type_header:
                                with col_widget: st.caption(f"**{map_entry['uph_type']}**")
                                current_uph_type_header = map_entry['uph_type']
                            # else: with col_widget: st.caption("") # Empty for subsequent colors of same type

                    # Swatch Headers
                    cols_swatch_header = st.columns([2.5] + [1] * num_data_columns)
                    for i, col_widget in enumerate(cols_swatch_header):
                        if i == 0: with col_widget: st.caption("")
                        else:
                            sw_url = data_column_map[i-1]['swatch']
                            with col_widget:
                                if sw_url: st.image(sw_url, width=30)
                                else: st.markdown("<div style='height:30px; width:30px;'></div>", unsafe_allow_html=True)
                    
                    # Color Number Headers
                    cols_color_num_header = st.columns([2.5] + [1] * num_data_columns)
                    for i, col_widget in enumerate(cols_color_num_header):
                        if i == 0: with col_widget: st.caption("")
                        else:
                            with col_widget: st.caption(f"<small>{data_column_map[i-1]['uph_color']}</small>", unsafe_allow_html=True)
                    st.markdown("---")

                    # --- Display Product Rows with Checkboxes ---
                    for prod_name in products_in_family:
                        cols_product_row = st.columns([2.5] + [1] * num_data_columns)
                        with cols_product_row[0]:
                            st.markdown(f"**{prod_name}**")
                        
                        for i, col_widget in enumerate(cols_product_row[1:]):
                            current_col_uph_type = data_column_map[i]['uph_type']
                            current_col_uph_color = data_column_map[i]['uph_color']
                            
                            # Check if this product is available in this specific textile/color
                            item_exists_df = family_df[
                                (family_df['Product Display Name'] == prod_name) &
                                (family_df['Upholstery Type'].fillna("N/A") == current_col_uph_type) &
                                (family_df['Upholstery Color'].astype(str).fillna("N/A") == current_col_uph_color)
                            ]
                            
                            with col_widget:
                                if not item_exists_df.empty:
                                    # Key for checkbox and session state
                                    cb_key_str = f"cb_{selected_family}_{prod_name}_{current_col_uph_type}_{current_col_uph_color}".replace(" ","_").replace("/","_")
                                    is_gen_selected = cb_key_str in st.session_state.matrix_selected_generic_items
                                    
                                    st.checkbox(" ", value=is_gen_selected, key=cb_key_str,
                                                on_change=handle_matrix_cb_toggle, 
                                                args=(prod_name, current_col_uph_type, current_col_uph_color, cb_key_str),
                                                label_visibility="collapsed")
                                else:
                                    st.markdown("<div style='height:30px;'>-</div>", unsafe_allow_html=True) # Placeholder
                        st.markdown("---")
        else:
            if selected_family and selected_family != DEFAULT_NO_SELECTION : st.info(f"Ingen data fundet for produktfamilien: {selected_family}")


    # --- Trin 2: Benfarve Valg (Base Color Selection) ---
    items_needing_base_choice_now = [
        item_data for key, item_data in st.session_state.matrix_selected_generic_items.items() if item_data.get('requires_base_choice')
    ]
    if items_needing_base_choice_now:
        st.header("Trin 2: Vælg Benfarver")
        for generic_item in items_needing_base_choice_now:
            item_key = generic_item['key']
            st.markdown(f"**{generic_item['product']}** ({generic_item['upholstery_type']} - {generic_item['upholstery_color']})")
            chosen_bases = st.multiselect(
                f"Tilgængelige benfarver:",
                options=generic_item['available_bases'],
                default=st.session_state.user_chosen_base_colors_for_items.get(item_key, []),
                key=f"ms_base_{item_key}"
            )
            st.session_state.user_chosen_base_colors_for_items[item_key] = chosen_bases
            st.markdown("---")

    # --- Trin 3: Bekræft Valg ---
    if st.session_state.matrix_selected_generic_items:
        if st.button("Bekræft Valg og Opdater Endelig Liste", key="confirm_button"):
            st.session_state.final_items_for_download = [] # Reset
            for key, gen_item_data in st.session_state.matrix_selected_generic_items.items():
                if not gen_item_data['requires_base_choice']:
                    st.session_state.final_items_for_download.append({
                        "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']}" + (f" / Ben: {gen_item_data['resolved_base_if_single']}" if pd.notna(gen_item_data['resolved_base_if_single']) else ""),
                        "item_no": gen_item_data['item_no_if_single_base'],
                        "article_no": gen_item_data['article_no_if_single_base']
                    })
                else:
                    selected_bases_for_this = st.session_state.user_chosen_base_colors_for_items.get(key, [])
                    if not selected_bases_for_this:
                        st.warning(f"Ingen benfarve valgt for {gen_item_data['product']} ({gen_item_data['upholstery_type']}/{gen_item_data['upholstery_color']}). Springes over.")
                        continue
                    for bc in selected_bases_for_this:
                        # Find specific item in raw_df
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
                                "description": f"{gen_item_data['family']} / {gen_item_data['product']} / {gen_item_data['upholstery_type']} / {gen_item_data['upholstery_color']} / Ben: {bc}",
                                "item_no": actual_item['Item No'],
                                "article_no": actual_item['Article No']
                            })
                        else:
                            st.error(f"FEJL: Kunne ikke finde vare for {gen_item_data['product']} med ben {bc}.")
            
            # Deduplicate final list
            final_list_unique = []
            seen_item_nos_final = set()
            for item in st.session_state.final_items_for_download:
                if item['item_no'] not in seen_item_nos_final:
                    final_list_unique.append(item)
                    seen_item_nos_final.add(item['item_no'])
            st.session_state.final_items_for_download = final_list_unique
            st.success("Endelig liste opdateret!")
            st.rerun()


    # --- Trin 4: Gennemse Valgte Kombinationer (Endelig Liste) ---
    if st.session_state.final_items_for_download: 
        st.header("Trin 4: Gennemse Endelige Valgte Kombinationer") 
        for i, combo in enumerate(st.session_state.final_items_for_download):
            col1, col2 = st.columns([0.9, 0.1])
            col1.write(f"{i+1}. {combo['description']} (Vare: {combo['item_no']})")
            if col2.button(f"Fjern", key=f"final_remove_{i}_{combo['item_no']}"): 
                item_to_remove_no = st.session_state.final_items_for_download.pop(i)['item_no']
                # Logic to potentially uncheck original checkbox if needed (can be complex)
                st.toast(f"Fjernet: {item_to_remove_no}", icon="🗑️") 
                st.rerun() 
        
        # --- Trin 5: Vælg Valuta ---
        st.header("Trin 5: Vælg Valuta")
        try:
            if not st.session_state.wholesale_prices_df.empty:
                article_no_col_name_ws = st.session_state.wholesale_prices_df.columns[0]
                currency_options = [col for col in st.session_state.wholesale_prices_df.columns if str(col).lower() != str(article_no_col_name_ws).lower()]
            else: currency_options = []
            selected_currency = st.selectbox("Vælg Valuta:", options=currency_options, key="currency_selector") if currency_options else None
            if not currency_options and not st.session_state.wholesale_prices_df.empty: st.error("Ingen valuta kolonner fundet.")
            elif st.session_state.wholesale_prices_df.empty: st.error("Pris Matrix (wholesale) er tom.")
        except Exception as e: st.error(f"Fejl ved valuta: {e}"); selected_currency = None

        # --- Trin 6: Generer Masterdata Fil ---
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
                    
                    if not st.session_state.wholesale_prices_df.empty:
                        ws_price_row_df = st.session_state.wholesale_prices_df[st.session_state.wholesale_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                        if not ws_price_row_df.empty and selected_currency in ws_price_row_df.columns:
                            output_row_dict["Wholesale price"] = ws_price_row_df.iloc[0][selected_currency] if pd.notna(ws_price_row_df.iloc[0][selected_currency]) else "N/A"
                        else: output_row_dict["Wholesale price"] = "Pris Ikke Fundet"
                    else: output_row_dict["Wholesale price"] = "Matrix Tom"

                    if not st.session_state.retail_prices_df.empty:
                        rt_price_row_df = st.session_state.retail_prices_df[st.session_state.retail_prices_df.iloc[:, 0].astype(str) == str(article_no_to_find)]
                        if not rt_price_row_df.empty and selected_currency in rt_price_row_df.columns:
                            output_row_dict["Retail price"] = rt_price_row_df.iloc[0][selected_currency] if pd.notna(rt_price_row_df.iloc[0][selected_currency]) else "N/A"
                        else: output_row_dict["Retail price"] = "Pris Ikke Fundet"
                    else: output_row_dict["Retail price"] = "Matrix Tom"
                    
                    output_data.append(output_row_dict)
                else: st.warning(f"Data for Vare Nr: {item_no_to_find} ikke fundet.")
            if output_data:
                output_df = pd.DataFrame(output_data, columns=master_template_columns_ordered)
                output_excel_buffer = io.BytesIO()
                with pd.ExcelWriter(output_excel_buffer, engine='xlsxwriter') as writer: output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
                output_excel_buffer.seek(0)
                st.download_button(label="📥 Download Masterdata Excel Fil", data=output_excel_buffer, file_name=f"masterdata_output_{selected_currency.replace(' ', '_').replace('.', '')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else: st.warning("Ingen data at generere.")
        elif not selected_currency and st.session_state.final_items_for_download : st.warning("Vælg venligst en valuta.")
    elif not st.session_state.matrix_selected_generic_items and not st.session_state.final_items_for_download : 
         st.info("Foretag valg i Trin 1 for at fortsætte.")


else: 
    st.error("En eller flere datafiler kunne ikke indlæses korrekt. Kontroller stier og filformater.")


# --- Styling (Optional) ---
st.markdown("""
<style>
    h1 { color: #333; }
    h2 { color: #1E40AF; border-bottom: 2px solid #BFDBFE; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    h3 { /* Product Subheaders in new layout */
        font-size: 1.1em;
        font-weight: bold;
        margin-top: 15px;
        margin-bottom: 5px;
        /* border-bottom: 1px solid #eee; */
        /* padding-bottom: 3px; */
    }
    /* Styling for the new matrix/list layout */
    .color-item-row { /* Class for each color row under a product */
        display: flex;
        align-items: center;
        padding: 3px 0;
        border-bottom: 1px solid #f0f0f0; /* Light separator between color items */
    }
    .color-item-row > div { /* Direct children of the row for alignment */
        padding: 0 5px;
    }
    .color-item-row img { /* Swatch image */
        max-height: 25px;
        max-width: 25px;
        border: 1px solid #ddd;
        border-radius: 3px;
    }
    /* Captions for the matrix-like headers */
    div[data-testid="stCaptionContainer"] > div > p { 
        font-weight: bold; font-size: 0.8em !important; color: #4A5568 !important; 
        text-align: center; white-space: normal; overflow-wrap: break-word; line-height: 1.2;
        padding: 2px;
    }
     div[data-testid="stCaptionContainer"] img { /* Swatch in header */
        max-height: 20px !important; margin-right: 3px; 
    }

    hr { margin-top: 0.2rem; margin-bottom: 0.2rem; border-top: 1px solid #e2e8f0; } 
    .stButton>button { font-size: 0.9em; padding: 0.3em 0.7em; }
    small { color: #718096; font-size:0.9em; display:block; line-height:1.1; } 
</style>
""", unsafe_allow_html=True)
