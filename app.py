import streamlit as st
import pandas as pd
from io import BytesIO

# Load Excel files
@st.cache_data
def load_data():
    raw = pd.read_excel("raw-data.xlsx", sheet_name="APP")
    wholesale = pd.read_excel("price-matrix_EUROPE.xlsx", sheet_name="Price matrix wholesale")
    retail = pd.read_excel("price-matrix_EUROPE.xlsx", sheet_name="Price matrix retail")
    template = pd.read_excel("Masterdata-output-template.xlsx")
    return raw, wholesale, retail, template

raw_data, wholesale_prices, retail_prices, output_template = load_data()

# Initialize session state to store selections
if "selections" not in st.session_state:
    st.session_state.selections = []

st.title("Muuto Configurator Tool")
st.markdown("""
Welcome to the Muuto configurator.

This tool helps you:
1. Select multiple product and textile combinations
2. Choose a currency
3. Download a complete masterdata file with prices.

Let’s begin!
""")

# Step 1: Product family
product_families = sorted(raw_data['Product Family'].dropna().unique())
selected_family = st.selectbox("Step 1: Choose a product family", product_families)

# Step 2: Product in family
filtered_family = raw_data[raw_data['Product Family'] == selected_family]
valid_products = filtered_family[filtered_family['Product Type'].notna() & (filtered_family['Product Type'] != "N/A")].copy()

def build_product_label(row):
    label = f"{row['Product Type']} {row['Product Model']}".strip()
    if row['Product Type'] == "Sofa Chaise Longue" and pd.notna(row['Sofa Direction']):
        label += f" - {row['Sofa Direction']}"
    return label

valid_products["Display Name"] = valid_products.apply(build_product_label, axis=1)
product_options = sorted(valid_products["Display Name"].unique())
selected_product_label = st.selectbox("Step 2: Choose a specific product", product_options)

# Step 3: Upholstery type
product_row = valid_products[valid_products["Display Name"] == selected_product_label]
upholstery_families = product_row["Upholstery Type"].dropna().unique()
selected_upholstery = st.selectbox("Step 3: Choose a textile family", upholstery_families)

# Step 4: Upholstery color
color_rows = product_row[product_row["Upholstery Type"] == selected_upholstery]
color_options = color_rows["Upholstery Color"].dropna().unique()
selected_color = st.selectbox("Step 4: Choose a color", color_options)

# Step 5: Base color if applicable
base_color_rows = color_rows[color_rows["Upholstery Color"] == selected_color]
base_colors = base_color_rows["Base Color"].dropna().unique()
selected_base = None
if len(base_colors) > 1:
    selected_base = st.selectbox("Step 5: Choose a base color", base_colors)
elif len(base_colors) == 1:
    selected_base = base_colors[0]

# Step 6: Confirm and show selected Item No
filtered_final = base_color_rows[base_color_rows["Base Color"] == selected_base] if selected_base else base_color_rows
item_numbers = filtered_final["Item No"].dropna().unique()
selected_item_no = item_numbers[0] if len(item_numbers) == 1 else st.selectbox("Confirm Item No", item_numbers)

# Add to selection list
if st.button("Add to selection"):
    article_no = filtered_final[filtered_final["Item No"] == selected_item_no]["Article No"].values[0]
    selection = {
        "Item No": selected_item_no,
        "Article No": article_no
    }
    if selection not in st.session_state.selections:
        st.session_state.selections.append(selection)
        st.success(f"Added {selected_item_no} to selection list.")
    else:
        st.warning("This combination is already added.")

# Show selection list
if st.session_state.selections:
    st.subheader("Selected combinations")
    for idx, sel in enumerate(st.session_state.selections):
        st.markdown(f"- **{sel['Item No']}** (Article No: {sel['Article No']})")

    # Step 7: Choose currency once for all
    available_currencies = [col for col in wholesale_prices.columns if col != "Article No."]
    selected_currency = st.selectbox("Step 6: Choose your currency", available_currencies)

    # Step 8: Generate masterdata file for all selections
    if st.button("Generate masterdata file for all"):
        export_rows = []
        for sel in st.session_state.selections:
            item_no = sel["Item No"]
            article_no = sel["Article No"]
            ws_price = wholesale_prices.loc[wholesale_prices["Article No."] == article_no, selected_currency].values
            rt_price = retail_prices.loc[retail_prices["Article No."] == article_no, selected_currency].values
            matched_row = raw_data[raw_data["Item No"] == item_no].copy()
            if not matched_row.empty:
                output_row = output_template.copy()
                for col in output_template.columns:
                    if col in matched_row.columns:
                        output_row.loc[0, col] = matched_row.iloc[0][col]
                output_row.loc[0, "Wholesale price"] = ws_price[0] if len(ws_price) else ""
                output_row.loc[0, "Retail Price"] = rt_price[0] if len(rt_price) else ""
                export_rows.append(output_row)

        # Concatenate all selections
        final_export = pd.concat(export_rows, ignore_index=True)
        output = BytesIO()
        final_export.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="📥 Download masterdata file with all selections",
            data=output,
            file_name=f"Muuto_masterdata_{selected_currency}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
