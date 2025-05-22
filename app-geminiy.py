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
            required_initial_cols = ['Product Type', 'Product Model', 'Sofa Direction', 'Base Color', 'Product Family', 'Item No', 'Article No', 'Image URL swatch', 'Upholstery Type', 'Upholstery Color'] # Base Color still needed for internal logic if we re-add it
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
                # Base Color Cleaned might still be useful if we re-introduce base color logic later.
                # For now, it's not directly used in the simplified table display.
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
    if st.session_state.selected_family_session not in available_families_in_view: st.session_state.selected_family_session = DEFAULT_NO_SELECTION
    
    selected_family_idx = 0
    if st.session_state.selected_family_session in available_families_in_view:
        selected_family_idx = available_families_in_view.index(st.session_state.selected_family_session)
        
    selected_family = st.selectbox("Vælg Produkt Familie (filtrerer nuværende visning):", options=available_families_in_view, index=selected_family_idx, key="family_selector_main")
    st.session_state.selected_family_session = selected_family

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

    if selected_family and selected_family != DEFAULT_NO_SELECTION and 'Product Family' in df_for_display.columns:
        family_specific_df = df_for_display[df_for_display['Product Family'] == selected_family]

        if not family_specific_df.empty and 'Product Display Name' in family_specific_df.columns and \
           'Upholstery Type' in family_specific_df.columns and 'Upholstery Color' in family_specific_df.columns:
            
            unique_products = sorted(family_specific_df['Product Display Name'].dropna().unique())
            
            # Determine textile/color combinations ONLY from the selected family_specific_df
            textile_color_combos_df = family_specific_df[['Upholstery Type', 'Upholstery Color']].drop_duplicates().copy()
            textile_color_combos_df['Upholstery Type'] = textile_color_combos_df['Upholstery Type'].fillna("N/A_Type")
            textile_color_combos_df['Upholstery Color'] = textile_color_combos_df['Upholstery Color'].fillna("N/A_Color")
            textile_color_combos_df['Combined Header'] = textile_color_combos_df['Upholstery Type'].astype(str) + " - " + textile_color_combos_df['Upholstery Color'].astype(str)
            sorted_textile_color_headers = sorted(textile_color_combos_df['Combined Header'].unique())
            
            if not unique_products:
                st.info(f"Ingen produkter fundet for familien: {selected_family}")
            elif not sorted_textile_color_headers:
                st.info(f"Ingen tekstil/farve kombinationer fundet for den valgte familie: {selected_family}")
            else:
                num_data_cols = len(sorted_textile_color_headers)
                # Adjusted column widths: Wider product column, slightly narrower textile/color columns
                col_widths = [3.0] + [1.5] * num_data_cols # Increased product column width further
                
                header_row = st.columns(col_widths)
                with header_row[0]:
                    st.caption("**Produkt**")
                for i, header_text in enumerate(sorted_textile_color_headers):
                    with header_row[i+1]:
                        # Make header text wrap and display swatch
                        uph_type_h, uph_color_h = header_text.split(" - ", 1)
                        # Find a representative swatch for this header (first one found in the family for this combo)
                        representative_swatch_series = family_specific_df[
                            (family_specific_df['Upholstery Type'].fillna("N/A_Type").astype(str) == uph_type_h) &
                            (family_specific_df['Upholstery Color'].fillna("N/A_Color").astype(str) == uph_color_h)
                        ]['Image URL swatch'].dropna()
                        
                        # Display swatch and text in the header cell
                        header_cell_cols = st.columns([0.4, 1]) # Swatch, Text
                        with header_cell_cols[0]:
                             if not representative_swatch_series.empty and pd.notna(representative_swatch_series.iloc[0]):
                                st.image(representative_swatch_series.iloc[0], width=25)
                             else: # Placeholder if no swatch
                                st.markdown("<div style='width:25px; height:25px; font-size:0.5em; border:1px solid #ddd;'></div>", unsafe_allow_html=True)
                        with header_cell_cols[1]:
                            st.caption(f"**{header_text}**")


                st.markdown("---")

                for product_name_iter in unique_products:
                    product_row = st.columns(col_widths)
                    with product_row[0]:
                        st.markdown(f"**{product_name_iter}**")

                    product_df_current_row = family_specific_df[family_specific_df['Product Display Name'] == product_name_iter]

                    for i, combined_header_text in enumerate(sorted_textile_color_headers):
                        try:
                            uph_type_col_header, uph_color_col_header = combined_header_text.split(" - ", 1)
                            uph_type_filter = uph_type_col_header if uph_type_col_header != "N/A_Type" else pd.NA
                            uph_color_filter = uph_color_col_header if uph_color_col_header != "N/A_Color" else pd.NA
                        except ValueError:
                            st.error(f"Fejl i header dekonstruktion: {combined_header_text}")
                            continue
                        
                        with product_row[i+1]:
                            if pd.isna(uph_type_filter): type_condition = product_df_current_row['Upholstery Type'].isna()
                            else: type_condition = (product_df_current_row['Upholstery Type'].astype(str) == str(uph_type_filter))
                            
                            if pd.isna(uph_color_filter): color_condition = product_df_current_row['Upholstery Color'].isna()
                            else: color_condition = (product_df_current_row['Upholstery Color'].astype(str) == str(uph_color_filter))

                            cell_items_df = product_df_current_row[type_condition & color_condition]

                            if cell_items_df.empty:
                                st.markdown("<div style='height: 50px; display:flex; align-items:center; justify-content:center;'>-</div>", unsafe_allow_html=True)
                            else:
                                first_item_in_cell = cell_items_df.iloc[0]
                                item_no_cell = first_item_in_cell['Item No']
                                article_no_cell = first_item_in_cell['Article No']
                                base_color_for_desc = str(first_item_in_cell.get('Base Color', "N/A")) 
                                # Swatch is now part of the column header, not repeated in cell unless desired
                                # swatch_url_cell = first_item_in_cell.get('Image URL swatch')

                                desc_parts_cell = [selected_family, product_name_iter, uph_type_col_header, uph_color_col_header]
                                full_desc_cell = " / ".join(map(str, desc_parts_cell)) 

                                item_data_for_cb_cell = {
                                    "description": full_desc_cell, "item_no": item_no_cell, "article_no": article_no_cell,
                                    "family": selected_family, "product": product_name_iter,
                                    "textile_family": uph_type_col_header, "textile_color": uph_color_col_header,
                                    "base_color": base_color_for_desc 
                                }
                                is_selected_cell = any(sel['item_no'] == item_no_cell for sel in st.session_state.selected_items_for_download)

                                with st.container():
                                    # Cell now only needs the checkbox
                                    # The swatch is in the header. If multiple items per cell (due to base color diff), this simplified model picks the first.
                                    st.checkbox(" ", value=is_selected_cell, key=f"cb_cell_{item_no_cell}_{product_name_iter}_{i}", 
                                                on_change=handle_item_checkbox_toggle, args=(item_data_for_cb_cell, f"cb_cell_{item_no_cell}_{product_name_iter}_{i}"),
                                                label_visibility="collapsed")
                                    # Optionally, show Item No subtly if needed for user reference
                                    # st.markdown(f"<small><i>({item_no_cell})</i></small>", unsafe_allow_html=True)
                    st.markdown("---") 
        elif 'Product Display Name' not in family_specific_df.columns and not family_specific_df.empty:
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
        font-size: 0.75em !important; /* Smaller caption for table headers */
        color: #4A5568 !important; 
        padding-bottom: 1px;
        text-align: center; 
        white-space: normal; 
        overflow-wrap: break-word; 
        line-height: 1.1; /* Tighter line height for headers */
        display: flex; /* For aligning swatch and text */
        align-items: center;
        justify-content: center;
        min-height: 40px; /* Ensure header cells have some height */
    }
    /* Swatch in header */
    div[data-testid="stCaptionContainer"] img {
        max-height: 25px !important; /* Smaller swatch in header */
        margin-right: 5px; /* Space between swatch and text in header */
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
         padding-top: 10px; /* More padding for product name */
         padding-left: 5px;
         font-size: 0.9em; /* Slightly larger font for product name */
         justify-content: flex-start !important;
         text-align: left !important;
    }
    /* Styling for data cells (checkbox only) */
    div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] > div.stCheckbox {
        display: flex; 
        align-items: center; 
        justify-content: center; 
        height: auto; 
        min-height: 45px; /* Min height for data cells */
        padding: 1px; 
        border-left: 1px solid #f0f0f0; 
    }
     div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] > div:first-child > div[data-testid="stMarkdownContainer"] {
        border-left: none; /* No left border for the first (product name) column */
    }

    hr { margin-top: 0.1rem; margin-bottom: 0.1rem; border-top: 1px solid #e2e8f0; } 
    .stButton>button { font-size: 0.9em; padding: 0.3em 0.7em; }
    small { color: #718096; font-size:0.9em; display:block; line-height:1.1; } 
</style>
""", unsafe_allow_html=True)
