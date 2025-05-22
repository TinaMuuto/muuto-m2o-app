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
if 'selected_upholstery_type_session' not in st.session_state: st.session_state.selected_upholstery_type_session = None # New session state for upholstery type
if 'selected_items_for_download' not in st.session_state: st.session_state.selected_items_for_download = []

# --- Load Data Directly from XLSX files ---
files_loaded_successfully = True
st.sidebar.subheader("Diagnostik for Filindlæsning:")
st.sidebar.write(f"Forventet script-mappe (BASE_DIR): `{BASE_DIR}`")
try:
    st.sidebar.write(f"Filer fundet i script-mappen:")
    found_files_in_basedir = os.listdir(BASE_DIR)
    if found_files_in_basedir:
        for f_name in found_files_in_basedir:
            st.sidebar.code(f_name)
    else:
        st.sidebar.warning("Ingen filer fundet i script-mappen ifølge os.listdir().")
except Exception as e_listdir:
    st.sidebar.error(f"Fejl ved listning af filer i BASE_DIR: {e_listdir}")


if st.session_state.raw_df is None:
    if os.path.exists(RAW_DATA_XLSX_PATH):
        try:
            st.session_state.raw_df = pd.read_excel(RAW_DATA_XLSX_PATH, sheet_name=RAW_DATA_APP_SHEET)
            required_initial_cols = ['Product Type', 'Product Model', 'Sofa Direction', 'Base Color', 'Product Family', 'Item No', 'Article No', 'Image URL swatch', 'Upholstery Type', 'Upholstery Color']
            missing_initial = [col for col in required_initial_cols if col not in st.session_state.raw_df.columns]
            
            if 'Market' not in st.session_state.raw_df.columns:
                st.warning("Kolonnen 'Market' blev ikke fundet i rådata. UK-filtrering springes over.")
            else:
                st.session_state.raw_df = st.session_state.raw_df[st.session_state.raw_df['Market'].astype(str).str.upper() != 'UK']
                st.info("Rådata er filtreret for at ekskludere 'UK' markedet.")

            if missing_initial:
                st.error(f"Nødvendige kolonner mangler i '{os.path.basename(RAW_DATA_XLSX_PATH)}' (Ark: {RAW_DATA_APP_SHEET}): {missing_initial}. Kan ikke fortsætte indlæsning.")
                files_loaded_successfully = False
            else:
                st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
                st.session_state.raw_df['Base Color Cleaned'] = st.session_state.raw_df['Base Color'].astype(str).str.strip().replace("N/A", pd.NA)
        except Exception as e:
            st.error(f"Error loading Raw Data from '{os.path.basename(RAW_DATA_XLSX_PATH)}' (Sheet: {RAW_DATA_APP_SHEET}): {e}"); files_loaded_successfully = False
    else: 
        st.error(f"Raw Data Excel file not found at: {RAW_DATA_XLSX_PATH}")
        files_loaded_successfully = False

if files_loaded_successfully and st.session_state.wholesale_prices_df is None :
    if os.path.exists(PRICE_MATRIX_XLSX_PATH):
        try:
            st.session_state.wholesale_prices_df = pd.read_excel(PRICE_MATRIX_XLSX_PATH, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
        except Exception as e: st.error(f"Error loading Wholesale Price Matrix from '{os.path.basename(PRICE_MATRIX_XLSX_PATH)}' (Sheet: {PRICE_MATRIX_WHOLESALE_SHEET}): {e}"); files_loaded_successfully = False
    else: 
        st.error(f"Price Matrix Excel file not found: {PRICE_MATRIX_XLSX_PATH}")
        files_loaded_successfully = False

if files_loaded_successfully and st.session_state.retail_prices_df is None :
    if os.path.exists(PRICE_MATRIX_XLSX_PATH): 
        try:
            st.session_state.retail_prices_df = pd.read_excel(PRICE_MATRIX_XLSX_PATH, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
        except Exception as e: st.error(f"Error loading Retail Price Matrix from '{os.path.basename(PRICE_MATRIX_XLSX_PATH)}' (Sheet: {PRICE_MATRIX_RETAIL_SHEET}): {e}"); files_loaded_successfully = False
    else: 
        if os.path.exists(PRICE_MATRIX_XLSX_PATH): 
             st.error(f"Price Matrix Excel file found, but error loading sheet: {PRICE_MATRIX_RETAIL_SHEET}")
        files_loaded_successfully = False


if files_loaded_successfully and st.session_state.template_cols is None:
    if os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH):
        try:
            template_df = pd.read_excel(MASTERDATA_TEMPLATE_XLSX_PATH) 
            st.session_state.template_cols = template_df.columns.tolist()
            if "Wholesale price" not in st.session_state.template_cols: st.session_state.template_cols.append("Wholesale price")
            if "Retail price" not in st.session_state.template_cols: st.session_state.template_cols.append("Retail price")
        except Exception as e: st.error(f"Error loading Masterdata Template from '{os.path.basename(MASTERDATA_TEMPLATE_XLSX_PATH)}': {e}"); files_loaded_successfully = False
    else: 
        st.error(f"Masterdata Template Excel file not found: {MASTERDATA_TEMPLATE_XLSX_PATH}")
        files_loaded_successfully = False


# --- Main Application Area ---
if files_loaded_successfully and all(df is not None for df in [st.session_state.raw_df, st.session_state.wholesale_prices_df, st.session_state.retail_prices_df]) and st.session_state.template_cols:
    
    st.header("Trin 1: Vælg Produkt Kombinationer")
    search_query = st.text_input("Søg på Produkt Familie eller Produkt Navn:", value=st.session_state.search_query_session, key="search_field")
    st.session_state.search_query_session = search_query
    search_query_lower = search_query.lower().strip()

    df_for_display = st.session_state.raw_df.copy() 
    if search_query_lower:
        if 'Product Family' in df_for_display.columns and 'Product Display Name' in df_for_display.columns:
            df_for_display = df_for_display[
                df_for_display.apply(lambda row: search_query_lower in str(row['Product Family']).lower() or \
                                               search_query_lower in str(row['Product Display Name']).lower(), axis=1)
            ]
        else:
            st.warning("Kolonnerne 'Product Family' eller 'Product Display Name' mangler for at kunne søge.")
        if df_for_display.empty and ('Product Family' in st.session_state.raw_df.columns):
             st.info(f"Ingen produkter fundet for søgningen: '{search_query}'")

    available_families_in_view = [DEFAULT_NO_SELECTION] + sorted(df_for_display['Product Family'].dropna().unique()) if 'Product Family' in df_for_display.columns else [DEFAULT_NO_SELECTION]
    if st.session_state.selected_family_session not in available_families_in_view: 
        st.session_state.selected_family_session = DEFAULT_NO_SELECTION
        st.session_state.selected_upholstery_type_session = None # Reset upholstery if family changes
    
    selected_family_idx = 0
    if st.session_state.selected_family_session in available_families_in_view:
        selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)
        
    selected_family = st.selectbox("Vælg Produkt Familie:", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
    if selected_family != st.session_state.selected_family_session: # If family changed, reset upholstery
        st.session_state.selected_upholstery_type_session = None
    st.session_state.selected_family_session = selected_family
    

    # --- Upholstery Type Selection ---
    selected_upholstery_type = None
    if selected_family and selected_family != DEFAULT_NO_SELECTION and 'Upholstery Type' in df_for_display.columns:
        family_specific_df_for_upholstery = df_for_display[df_for_display['Product Family'] == selected_family]
        available_upholstery_types = [DEFAULT_NO_SELECTION] + sorted(family_specific_df_for_upholstery['Upholstery Type'].dropna().unique())
        
        if st.session_state.selected_upholstery_type_session not in available_upholstery_types:
            st.session_state.selected_upholstery_type_session = DEFAULT_NO_SELECTION

        upholstery_type_idx = 0
        if st.session_state.selected_upholstery_type_session in available_upholstery_types:
            upholstery_type_idx = available_upholstery_types.index(st.session_state.selected_upholstery_type_session)

        selected_upholstery_type = st.selectbox(
            "Vælg Tekstil Familie (Upholstery Type):", 
            options=available_upholstery_types, 
            index=upholstery_type_idx,
            key="upholstery_type_selector"
        )
        st.session_state.selected_upholstery_type_session = selected_upholstery_type
    
    # Callback for checkbox changes
    def handle_item_checkbox_toggle(item_details_dict, checkbox_key):
        is_checked_now = st.session_state[checkbox_key]
        current_selection_item_nos = [sel['item_no'] for sel in st.session_state.selected_items_for_download]

        if is_checked_now:
            if item_details_dict['item_no'] not in current_selection_item_nos:
                st.session_state.selected_items_for_download.append(item_details_dict)
                st.toast(f"Tilføjet: {item_details_dict['item_no']}", icon="✅")
        else:
            if item_details_dict['item_no'] in current_selection_item_nos:
                st.session_state.selected_items_for_download = [
                    sel for sel in st.session_state.selected_items_for_download if sel['item_no'] != item_details_dict['item_no']
                ]
                st.toast(f"Fjernet: {item_details_dict['item_no']}", icon="❌")

    # --- Display Product Table ---
    if selected_family and selected_family != DEFAULT_NO_SELECTION and \
       selected_upholstery_type and selected_upholstery_type != DEFAULT_NO_SELECTION and \
       'Product Family' in df_for_display.columns:
        
        # Filter by selected product family AND upholstery type
        table_df = df_for_display[
            (df_for_display['Product Family'] == selected_family) &
            (df_for_display['Upholstery Type'].fillna("N/A_Type").astype(str) == selected_upholstery_type)
        ]

        if not table_df.empty and 'Product Display Name' in table_df.columns and \
           'Upholstery Color' in table_df.columns:
            
            unique_products_table = sorted(table_df['Product Display Name'].dropna().unique())
            
            # Determine color columns ONLY from the filtered table_df (selected family AND upholstery type)
            color_combos_df = table_df[['Upholstery Color']].drop_duplicates().copy() # Only Color now
            color_combos_df['Upholstery Color'] = color_combos_df['Upholstery Color'].fillna("N/A_Color")
            # Headers will be just the color names/codes
            sorted_color_headers = sorted(color_combos_df['Upholstery Color'].astype(str).unique())
            
            if not unique_products_table:
                st.info(f"Ingen produkter fundet for '{selected_family}' med tekstil '{selected_upholstery_type}'.")
            elif not sorted_color_headers:
                st.info(f"Ingen farver fundet for '{selected_upholstery_type}' i familien '{selected_family}'.")
            else:
                num_color_cols = len(sorted_color_headers)
                col_widths_table = [3.0] + [1.0] * num_color_cols # Product name col, then color cols
                
                header_row_table = st.columns(col_widths_table)
                with header_row_table[0]:
                    st.caption(f"**Produkt ({selected_upholstery_type})**") # Indicate selected upholstery
                for i, color_header_text in enumerate(sorted_color_headers):
                    with header_row_table[i+1]:
                        # Find a representative swatch for this color within the current Upholstery Type and Family
                        rep_swatch_series_color = table_df[
                            table_df['Upholstery Color'].fillna("N/A_Color").astype(str) == color_header_text
                        ]['Image URL swatch'].dropna()
                        
                        header_cell_content_cols = st.columns([0.4, 1])
                        with header_cell_content_cols[0]:
                            if not rep_swatch_series_color.empty and pd.notna(rep_swatch_series_color.iloc[0]):
                                st.image(rep_swatch_series_color.iloc[0], width=25)
                            else:
                                st.markdown("<div style='width:25px; height:25px;'></div>", unsafe_allow_html=True)
                        with header_cell_content_cols[1]:
                            st.caption(f"**{color_header_text}**")
                st.markdown("---")

                for product_name_table_iter in unique_products_table:
                    product_row_table = st.columns(col_widths_table)
                    with product_row_table[0]:
                        st.markdown(f"**{product_name_table_iter}**")

                    product_df_current_row_table = table_df[table_df['Product Display Name'] == product_name_table_iter]

                    for i, color_col_header_text in enumerate(sorted_color_headers):
                        with product_row_table[i+1]:
                            color_filter_table = color_col_header_text if color_col_header_text != "N/A_Color" else pd.NA
                            
                            if pd.isna(color_filter_table): color_cond_table = product_df_current_row_table['Upholstery Color'].isna()
                            else: color_cond_table = (product_df_current_row_table['Upholstery Color'].astype(str) == str(color_filter_table))
                            
                            # Items matching Product + selected Upholstery Type + current Color column
                            cell_items_table_df = product_df_current_row_table[color_cond_table]

                            if cell_items_table_df.empty:
                                st.markdown("<div style='height: 50px; display:flex; align-items:center; justify-content:center;'>-</div>", unsafe_allow_html=True)
                            else:
                                # In this simplified model, we assume one item per Product/UpholsteryType/Color combo,
                                # or we pick the first if multiple (due to base color differences now ignored for selection)
                                first_item_in_cell_table = cell_items_table_df.iloc[0]
                                item_no_cell_table = first_item_in_cell_table['Item No']
                                article_no_cell_table = first_item_in_cell_table['Article No']
                                base_color_for_desc_table = str(first_item_in_cell_table.get('Base Color', "N/A"))

                                desc_parts_cell_table = [selected_family, product_name_table_iter, selected_upholstery_type, color_col_header_text]
                                full_desc_cell_table = " / ".join(map(str, desc_parts_cell_table)) 

                                item_data_for_cb_cell_table = {
                                    "description": full_desc_cell_table, "item_no": item_no_cell_table, "article_no": article_no_cell_table,
                                    "family": selected_family, "product": product_name_table_iter,
                                    "textile_family": selected_upholstery_type, "textile_color": color_col_header_text,
                                    "base_color": base_color_for_desc_table 
                                }
                                is_selected_cell_table = any(sel['item_no'] == item_no_cell_table for sel in st.session_state.selected_items_for_download)

                                with st.container():
                                    st.checkbox(" ", value=is_selected_cell_table, key=f"cb_table_{item_no_cell_table}", 
                                                on_change=handle_item_checkbox_toggle, args=(item_data_for_cb_cell_table, f"cb_table_{item_no_cell_table}"),
                                                label_visibility="collapsed")
                    st.markdown("---") 
        elif selected_family and selected_family != DEFAULT_NO_SELECTION and not (selected_upholstery_type and selected_upholstery_type != DEFAULT_NO_SELECTION):
            st.info("Vælg venligst en Tekstil Familie for at se produkter og farver.")
        elif 'Product Display Name' not in family_specific_df.columns and not family_specific_df.empty :
             st.warning("Nødvendig kolonne 'Product Display Name' mangler for at vise tabel.")

    elif not files_loaded_successfully:
        pass 
    elif 'Product Family' not in st.session_state.raw_df.columns : 
        st.error("Nødvendig kolonne 'Product Family' mangler i rådata. Kan ikke vise produkter.")
    elif search_query_lower and df_for_display.empty and ('Product Family' in st.session_state.raw_df.columns):
        pass 
    elif not search_query_lower and not (selected_family and selected_family != DEFAULT_NO_SELECTION):
         st.info("Indtast et søgeord eller vælg en Produkt Familie for at se tilgængelige varer.")


    # --- Display Current Selections (Final List) ---
    if st.session_state.selected_items_for_download: 
        st.header("Trin 2: Gennemse Valgte Kombinationer") 
        for i, combo in enumerate(st.session_state.selected_items_for_download):
            col1, col2 = st.columns([0.9, 0.1])
            col1.write(f"{i+1}. {combo['description']} (Vare: {combo['item_no']})")
            if col2.button(f"Fjern", key=f"sel_remove_{i}_{combo['item_no']}"): 
                item_to_remove_no = st.session_state.selected_items_for_download.pop(i)['item_no']
                st.toast(f"Fjernet fra valgte liste: {item_to_remove_no}", icon="🗑️") 
                st.rerun() 
        
        st.header("Trin 3: Vælg Valuta og Generer Fil")
        try:
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
            for combo_selection in st.session_state.selected_items_for_download: 
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
        elif not selected_currency and st.session_state.selected_items_for_download : st.warning("Vælg venligst en valuta.") 
    elif not st.session_state.selected_items_for_download : 
         st.info("Foretag valg i Trin 1 for at fortsætte.")


else: 
    st.error("En eller flere datafiler kunne ikke indlæses korrekt, eller nødvendige kolonner mangler. Kontroller stier, filformater og kolonnenavne i dine .xlsx-filer.")
    if not os.path.exists(RAW_DATA_XLSX_PATH): st.sidebar.error(f"FEJL: {RAW_DATA_XLSX_PATH} ikke fundet.")
    if not os.path.exists(PRICE_MATRIX_XLSX_PATH): st.sidebar.error(f"FEJL: {PRICE_MATRIX_XLSX_PATH} (for wholesale/retail) ikke fundet.")
    if not os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH): st.sidebar.error(f"FEJL: {MASTERDATA_TEMPLATE_XLSX_PATH} ikke fundet.")


# --- Styling (Optional) ---
st.markdown("""
<style>
    h1 { /* App Title */ color: #333; }
    h2 { /* Step Headers */ color: #1E40AF; border-bottom: 2px solid #BFDBFE; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    
    /* Styling for table headers (captions) */
    div[data-testid="stCaptionContainer"] > div > p { 
        font-weight: bold; 
        font-size: 0.75em !important; 
        color: #4A5568 !important; 
        padding-bottom: 1px;
        text-align: center; 
        white-space: normal; 
        overflow-wrap: break-word; 
        line-height: 1.1; 
        display: flex; 
        align-items: center;
        justify-content: center;
        min-height: 40px; 
    }
    /* Swatch in header */
    div[data-testid="stCaptionContainer"] img {
        max-height: 25px !important; 
        margin-right: 5px; 
    }

    div[data-testid="stImage"] img { 
        max-height: 30px; 
        border: 1px solid #e2e8f0; 
        border-radius: 3px; 
        padding: 1px; 
        margin: auto; 
        display: block; 
    }
    /* Product name column styling */
    div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] > div:first-child > div[data-testid="stMarkdownContainer"] {
         align-items: flex-start !important; 
         padding-top: 10px; 
         padding-left: 5px;
         font-size: 0.9em; 
         justify-content: flex-start !important;
         text-align: left !important;
    }
    /* Styling for data cells (checkbox only) */
    div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] > div.stCheckbox {
        display: flex; 
        align-items: center; 
        justify-content: center; 
        height: auto; 
        min-height: 45px; 
        padding: 1px; 
        border-left: 1px solid #f0f0f0; 
    }
     div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] > div:first-child > div[data-testid="stMarkdownContainer"] {
        border-left: none; 
    }

    hr { margin-top: 0.1rem; margin-bottom: 0.1rem; border-top: 1px solid #e2e8f0; } 
    .stButton>button { font-size: 0.9em; padding: 0.3em 0.7em; }
    small { color: #718096; font-size:0.9em; display:block; line-height:1.1; } 
</style>
""", unsafe_allow_html=True)
