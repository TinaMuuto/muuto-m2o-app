import streamlit as st
import pandas as pd
import io

# --- Configuration & Constants ---
RAW_DATA_APP_SHEET = "APP"
PRICE_MATRIX_WHOLESALE_SHEET = "Price matrix wholesale"
PRICE_MATRIX_RETAIL_SHEET = "Price matrix retail"
DEFAULT_NO_SELECTION = "---Please Select---"


# --- Helper Function to Construct Product Display Name ---
def construct_product_display_name(row):
    name_parts = []
    product_type = row.get('Product Type')
    product_model = row.get('Product Model')
    sofa_direction = row.get('Sofa Direction')

    if pd.notna(product_type) and str(product_type).strip().upper() != "N/A":
        name_parts.append(str(product_type))
    if pd.notna(product_model) and str(product_model).strip().upper() != "N/A":
        name_parts.append(str(product_model))

    if str(product_type) == "Sofa Chaise Longue": # Check actual value, not just if it exists
        if pd.notna(sofa_direction) and str(sofa_direction).strip().upper() != "N/A":
            name_parts.append(str(sofa_direction))
    return " - ".join(name_parts) if name_parts else "Unnamed Product"

# --- Main App Logic ---
st.set_page_config(layout="wide")
st.title("Product Configurator & Masterdata Generator")

# --- Initialize session state variables ---
if 'selected_combinations' not in st.session_state:
    st.session_state.selected_combinations = []
if 'raw_df' not in st.session_state:
    st.session_state.raw_df = None
if 'wholesale_prices_df' not in st.session_state:
    st.session_state.wholesale_prices_df = None
if 'retail_prices_df' not in st.session_state:
    st.session_state.retail_prices_df = None
if 'template_cols' not in st.session_state:
    st.session_state.template_cols = None
if 'current_selections' not in st.session_state:
    st.session_state.current_selections = {
        "family": None, "product_display_name": None, "textile_family": None,
        "textile_color": None, "base_color": None
    }

# --- File Uploads ---
st.sidebar.header("Upload Data Files")
raw_data_file = st.sidebar.file_uploader("Upload Raw Data Excel File (e.g., raw-data.xlsx)", type=["xlsx", "xls"])
price_matrix_file = st.sidebar.file_uploader("Upload Price Matrix Excel File (e.g., price-matrix_EUROPE.xlsx)", type=["xlsx", "xls"])
masterdata_template_file = st.sidebar.file_uploader("Upload Masterdata Output Template Excel File", type=["xlsx", "xls"])

# --- Load and Cache Data ---
# Use flags to ensure data is loaded only once or if files change
if raw_data_file and (st.session_state.raw_df is None or st.session_state.get('raw_data_file_name') != raw_data_file.name):
    try:
        st.session_state.raw_df = pd.read_excel(raw_data_file, sheet_name=RAW_DATA_APP_SHEET)
        st.session_state.raw_df['Product Display Name'] = st.session_state.raw_df.apply(construct_product_display_name, axis=1)
        st.session_state.raw_data_file_name = raw_data_file.name # Store filename to detect changes
        st.sidebar.success("Raw Data loaded.")
    except Exception as e:
        st.sidebar.error(f"Error loading Raw Data: {e}")
        st.session_state.raw_df = None

if price_matrix_file and (st.session_state.wholesale_prices_df is None or st.session_state.get('price_matrix_file_name') != price_matrix_file.name):
    try:
        st.session_state.wholesale_prices_df = pd.read_excel(price_matrix_file, sheet_name=PRICE_MATRIX_WHOLESALE_SHEET)
        st.session_state.retail_prices_df = pd.read_excel(price_matrix_file, sheet_name=PRICE_MATRIX_RETAIL_SHEET)
        st.session_state.price_matrix_file_name = price_matrix_file.name
        st.sidebar.success("Price Matrix loaded.")
    except Exception as e:
        st.sidebar.error(f"Error loading Price Matrix: {e}")
        st.session_state.wholesale_prices_df = None
        st.session_state.retail_prices_df = None

if masterdata_template_file and (st.session_state.template_cols is None or st.session_state.get('masterdata_template_file_name') != masterdata_template_file.name):
    try:
        template_df = pd.read_excel(masterdata_template_file)
        st.session_state.template_cols = template_df.columns.tolist()
        # Ensure mandatory price columns exist in the template or add them
        if "Wholesale price" not in st.session_state.template_cols:
            st.session_state.template_cols.append("Wholesale price")
            st.sidebar.warning("Added 'Wholesale price' to template columns.")
        if "Retail price" not in st.session_state.template_cols:
            st.session_state.template_cols.append("Retail price")
            st.sidebar.warning("Added 'Retail price' to template columns.")
        st.session_state.masterdata_template_file_name = masterdata_template_file.name
        st.sidebar.success("Masterdata Template loaded.")
    except Exception as e:
        st.sidebar.error(f"Error loading Masterdata Template: {e}")
        st.session_state.template_cols = None


# --- Main Application Area ---
if st.session_state.raw_df is not None and \
   st.session_state.wholesale_prices_df is not None and \
   st.session_state.retail_prices_df is not None and \
   st.session_state.template_cols is not None:

    st.header("1. Select Product Combination")
    cs = st.session_state.current_selections # shorthand for current selections

    # --- Product Family Selection ---
    available_families = [DEFAULT_NO_SELECTION] + sorted(st.session_state.raw_df['Product Family'].dropna().unique())
    cs['family'] = st.selectbox("Select Product Family:", options=available_families, key="product_family_selector",
                                index=available_families.index(cs['family']) if cs['family'] in available_families else 0)

    if cs['family'] and cs['family'] != DEFAULT_NO_SELECTION:
        family_df = st.session_state.raw_df[st.session_state.raw_df['Product Family'] == cs['family']].copy()

        # --- Product Selection ---
        available_products = [DEFAULT_NO_SELECTION] + sorted(family_df['Product Display Name'].dropna().unique())
        cs['product_display_name'] = st.selectbox("Select Product:", options=available_products, key="product_selector",
                                                  index=available_products.index(cs['product_display_name']) if cs['product_display_name'] in available_products else 0)

        if cs['product_display_name'] and cs['product_display_name'] != DEFAULT_NO_SELECTION:
            product_df = family_df[family_df['Product Display Name'] == cs['product_display_name']].copy()

            # --- Textile Family (Upholstery Type) Selection ---
            available_textile_families = [DEFAULT_NO_SELECTION] + sorted(product_df['Upholstery Type'].dropna().unique())
            cs['textile_family'] = st.selectbox("Select Textile Family (Upholstery Type):", options=available_textile_families, key="textile_family_selector",
                                                index=available_textile_families.index(cs['textile_family']) if cs['textile_family'] in available_textile_families else 0)

            if cs['textile_family'] and cs['textile_family'] != DEFAULT_NO_SELECTION:
                textile_family_df = product_df[product_df['Upholstery Type'] == cs['textile_family']].copy()

                # --- Textile Color (Upholstery Color) Selection with Swatch ---
                color_options_data = textile_family_df[['Upholstery Color', 'Image URL swatch']].drop_duplicates().dropna(subset=['Upholstery Color'])
                
                if not color_options_data.empty:
                    color_display_options = [DEFAULT_NO_SELECTION]
                    # Using a dictionary to map display name to actual color name for selection
                    st.session_state.color_name_map = {DEFAULT_NO_SELECTION: None}

                    for _, row in color_options_data.iterrows():
                        color_name = str(row['Upholstery Color'])
                        # Ensure unique display names if color names are not unique (e.g. if '1' appears multiple times for different swatches)
                        # For now, assuming Upholstery Color is distinct enough for selection here.
                        display_name = color_name
                        color_display_options.append(display_name)
                        st.session_state.color_name_map[display_name] = color_name


                    # Radio button for textile color selection
                    st.write("Select Textile Color (Upholstery Color):")
                    cols = st.columns(5) # Adjust number of columns for swatches
                    selected_textile_color_display = None

                    # Temporary state for radio button, as direct selection isn't straightforward with complex items
                    if 'selected_radio_color' not in st.session_state:
                        st.session_state.selected_radio_color = DEFAULT_NO_SELECTION

                    # Create a list of (display_name, swatch_url) for format_func or custom display
                    radio_options_list = []
                    swatch_map = {}
                    for _, row in color_options_data.iterrows():
                        color_name = str(row['Upholstery Color'])
                        swatch_url = row['Image URL swatch']
                        radio_options_list.append(color_name) # Use actual color name for radio options
                        swatch_map[color_name] = swatch_url
                    
                    # Filter out default from actual options for radio
                    actual_color_options = [opt for opt in radio_options_list if opt != DEFAULT_NO_SELECTION]
                    
                    # Use index to manage radio button selection state
                    current_color_index = 0 # Default to first actual color or "None"
                    if cs['textile_color'] and cs['textile_color'] in actual_color_options:
                        current_color_index = actual_color_options.index(cs['textile_color']) +1 # +1 because of DEFAULT_NO_SELECTION prepended for display
                    
                    # Prepend default for the radio button options list
                    full_radio_options = [DEFAULT_NO_SELECTION] + actual_color_options


                    cs['textile_color'] = st.radio(
                        "Available Colors:",
                        options=full_radio_options,
                        index=0, # Default to "---Please Select---"
                        key="textile_color_radio",
                        format_func=lambda x: f"{x} " + (f" (View Swatch)" if swatch_map.get(x) and pd.notna(swatch_map.get(x)) else "(No Swatch)")
                    )
                    
                    # Display swatches below the radio group
                    if actual_color_options:
                        st.write("Color Swatches:")
                        num_swatch_cols = 5 
                        swatch_cols = st.columns(num_swatch_cols)
                        for i, color_name_for_swatch in enumerate(actual_color_options):
                            swatch_url = swatch_map.get(color_name_for_swatch)
                            with swatch_cols[i % num_swatch_cols]:
                                if swatch_url and pd.notna(swatch_url):
                                    st.image(swatch_url, caption=str(color_name_for_swatch), width=60)
                                else:
                                    st.markdown(f"**{color_name_for_swatch}**<br>(No Swatch)", unsafe_allow_html=True)
                                if cs['textile_color'] == color_name_for_swatch:
                                     st.markdown(f"<span style='color:green;font-weight:bold;'>Selected</span>", unsafe_allow_html=True)


                    if cs['textile_color'] and cs['textile_color'] != DEFAULT_NO_SELECTION:
                        # Actual selected color name (not display name)
                        actual_selected_color = cs['textile_color']
                        color_df = textile_family_df[textile_family_df['Upholstery Color'].astype(str) == actual_selected_color].copy()

                        # --- Base Color Selection (Conditional) ---
                        available_base_colors_series = color_df['Base Color'].dropna()
                        # Ensure "N/A" string is treated as NA if necessary, or filter out
                        available_base_colors = sorted([str(bc) for bc in available_base_colors_series.unique() if str(bc).strip().upper() != "N/A"])
                        
                        cs['base_color'] = None # Reset
                        if len(available_base_colors) > 1:
                            base_color_options = [DEFAULT_NO_SELECTION] + available_base_colors
                            cs['base_color'] = st.radio("Select Base Color:", options=base_color_options, key="base_color_selector",
                                                        index=0)
                        elif len(available_base_colors) == 1:
                            cs['base_color'] = available_base_colors[0]
                            st.write(f"Base Color (auto-selected): {cs['base_color']}")
                        else:
                            st.write("No specific base color options for this selection.") # No selectable base colors

                        # --- Add to Selections Button ---
                        if st.button("Add Combination to List", key="add_combination"):
                            final_filter_df = color_df.copy()

                            if cs['base_color'] and cs['base_color'] != DEFAULT_NO_SELECTION:
                                final_filter_df = final_filter_df[final_filter_df['Base Color'].astype(str) == cs['base_color']]
                            elif not available_base_colors: # No base colors were available or selected
                                # If Base Color column might have N/A or blanks for items without distinct base colors
                                final_filter_df = final_filter_df[final_filter_df['Base Color'].isna() | (final_filter_df['Base Color'].astype(str).str.upper() == "N/A")]
                            # If len(available_base_colors) == 1, it's already set in cs['base_color'] and used for filtering if needed.
                            # If cs['base_color'] is DEFAULT_NO_SELECTION and len(available_base_colors) > 1, it means user didn't pick one, so this might be an invalid state for adding.
                            
                            if cs['base_color'] == DEFAULT_NO_SELECTION and len(available_base_colors) > 1:
                                st.error("Please select a Base Color when multiple options are available.")
                            elif not final_filter_df.empty:
                                if len(final_filter_df) == 1:
                                    item_row = final_filter_df.iloc[0]
                                    item_no = item_row['Item No']
                                    article_no = item_row['Article No']
                                    
                                    desc_parts = [cs['family'], cs['product_display_name'], cs['textile_family'], actual_selected_color]
                                    final_base_color_desc = cs['base_color'] if cs['base_color'] and cs['base_color'] != DEFAULT_NO_SELECTION else "N/A"
                                    if final_base_color_desc != "N/A": # Only add if it's a specific color
                                        desc_parts.append(final_base_color_desc)
                                    
                                    combination_desc = " / ".join(map(str, desc_parts))

                                    st.session_state.selected_combinations.append({
                                        "description": combination_desc,
                                        "item_no": item_no,
                                        "article_no": article_no,
                                        "family": cs['family'],
                                        "product": cs['product_display_name'],
                                        "textile_family": cs['textile_family'],
                                        "textile_color": actual_selected_color,
                                        "base_color": final_base_color_desc
                                    })
                                    st.success(f"Added: {combination_desc} (Item No: {item_no})")
                                    # Optionally reset some selections for the next item
                                    st.session_state.current_selections['textile_color'] = None
                                    st.session_state.current_selections['base_color'] = None
                                    if 'textile_color_radio' in st.session_state:
                                        st.session_state.textile_color_radio = DEFAULT_NO_SELECTION # Attempt to reset radio
                                    if 'base_color_selector' in st.session_state:
                                        st.session_state.base_color_selector = DEFAULT_NO_SELECTION


                                    st.experimental_rerun()
                                elif len(final_filter_df) > 1:
                                    st.error(f"Multiple items found ({len(final_filter_df)}) for this exact combination. Data ambiguity. Items: {final_filter_df['Item No'].tolist()}")
                                else:
                                    st.error("No specific item found for the selected combination. Please check data or selections.")
                            else:
                                st.error("No item found matching current criteria.")
                else:
                    st.warning("No color options available for the selected textile family.")
    else:
        if cs['family'] == DEFAULT_NO_SELECTION and any(val for key, val in cs.items() if key != 'family' and val and val != DEFAULT_NO_SELECTION):
            # Reset downstream selections if family is changed back to default
            for key in ['product_display_name', 'textile_family', 'textile_color', 'base_color']:
                st.session_state.current_selections[key] = None
            st.experimental_rerun()


    # --- Display Current Selections ---
    if st.session_state.selected_combinations:
        st.header("2. Review Selected Combinations")
        for i, combo in enumerate(st.session_state.selected_combinations):
            col1, col2 = st.columns([0.9, 0.1])
            col1.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']}, Article: {combo['article_no']})")
            if col2.button(f"Remove", key=f"remove_{i}_{combo['item_no']}"): # Unique key
                st.session_state.selected_combinations.pop(i)
                st.experimental_rerun()
        
        st.header("3. Select Currency and Generate File")
        # --- Currency Selection ---
        try:
            # Assuming currency columns start after 'Article No'
            # And that wholesale and retail have the same currency columns (use wholesale as primary)
            currency_options = [col for col in st.session_state.wholesale_prices_df.columns if str(col).lower() not in ['article no', 'article_no']]
            if not currency_options:
                 st.error("No currency columns found in the Price Matrix file (excluding 'Article No'). Please check the file format.")
                 selected_currency = None
            else:
                selected_currency = st.selectbox("Select Currency:", options=currency_options, key="currency_selector")

        except Exception as e:
            st.error(f"Could not determine currency options: {e}")
            selected_currency = None


        if st.button("Generate Masterdata File", key="generate_file") and selected_currency:
            output_data = []
            # Use template_cols from session state, which has Wholesale and Retail price columns ensured
            master_template_columns = st.session_state.template_cols
            raw_data_cols_to_copy = [col for col in master_template_columns if col not in ["Wholesale price", "Retail price"]]

            for combo in st.session_state.selected_combinations:
                item_no_to_find = combo['item_no']
                # Ensure article_no is treated as it is in the Excel file (string or number)
                article_no_to_find = combo['article_no'] 
                
                item_data_row_df = st.session_state.raw_df[st.session_state.raw_df['Item No'] == item_no_to_find]
                
                if not item_data_row_df.empty:
                    item_data_row = item_data_row_df.iloc[0]
                    output_row = {}
                    for col in master_template_columns: # Iterate through template columns to maintain order
                        if col == "Wholesale price": continue # Handle separately
                        if col == "Retail price": continue # Handle separately
                        if col in item_data_row:
                            output_row[col] = item_data_row[col]
                        else:
                            output_row[col] = None # Or pd.NA or ""

                    # Fetch Wholesale Price (ensure Article No types match for merging/lookup)
                    ws_article_col = st.session_state.wholesale_prices_df.columns[0] # Assume first col is Article No
                    ws_price_row = st.session_state.wholesale_prices_df[
                        st.session_state.wholesale_prices_df[ws_article_col].astype(str) == str(article_no_to_find)
                    ]
                    if not ws_price_row.empty and selected_currency in ws_price_row.columns:
                        price_val = ws_price_row.iloc[0][selected_currency]
                        output_row["Wholesale price"] = price_val if pd.notna(price_val) else "N/A"
                    else:
                        output_row["Wholesale price"] = "Price Not Found"

                    # Fetch Retail Price
                    rt_article_col = st.session_state.retail_prices_df.columns[0] # Assume first col is Article No
                    rt_price_row = st.session_state.retail_prices_df[
                        st.session_state.retail_prices_df[rt_article_col].astype(str) == str(article_no_to_find)
                    ]
                    if not rt_price_row.empty and selected_currency in rt_price_row.columns:
                        price_val = rt_price_row.iloc[0][selected_currency]
                        output_row["Retail price"] = price_val if pd.notna(price_val) else "N/A"
                    else:
                        output_row["Retail price"] = "Price Not Found"
                    
                    output_data.append(output_row)
                else:
                    st.warning(f"Could not find data for Item No: {item_no_to_find} in raw_data during output generation.")

            if output_data:
                output_df = pd.DataFrame(output_data, columns=master_template_columns)
                
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                    output_df.to_excel(writer, index=False, sheet_name='Masterdata Output')
                output_excel.seek(0)

                st.download_button(
                    label="📥 Download Masterdata Excel File",
                    data=output_excel,
                    file_name=f"masterdata_output_{selected_currency.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No data to generate. Please select and add combinations.")
        elif not selected_currency and st.session_state.selected_combinations : # Only show if generate was implicitly or explicitly not run due to no currency
             st.warning("Please select a currency to enable file generation.")


    elif st.session_state.raw_df is not None: # Files loaded, but no selections made yet
        st.info("Select options above and click 'Add Combination to List' to build your masterdata.")

else:
    st.info("👋 Welcome! Please upload all three required Excel files using the sidebar to begin configuring products.")

# --- Styling (Optional) ---
st.markdown("""
<style>
    /* Enlarge radio button labels slightly for better readability */
    div[data-testid="stRadio"] label span {
        font-size: 1.05em;
    }
    /* Ensure images in columns are reasonably sized */
    div[data-testid="stImage"] img {
        object-fit: contain;
        max-height: 70px;
        border: 1px solid #eee;
        border-radius: 4px;
        padding: 2px;
    }
    .stButton>button {
        width: auto; /* Ensure buttons are not overly wide */
    }
</style>
""", unsafe_allow_html=True)
