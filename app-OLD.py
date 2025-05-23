import streamlit as st
import pandas as pd
from io import BytesIO

# Load data
@st.cache_data
def load_data():
    raw = pd.read_excel("raw-data.xlsx", sheet_name="APP")
    wholesale = pd.read_excel("price-matrix_EUROPE.xlsx", sheet_name="Price matrix wholesale")
    retail = pd.read_excel("price-matrix_EUROPE.xlsx", sheet_name="Price matrix retail")
    template = pd.read_excel("Masterdata-output-template.xlsx")
    return raw, wholesale, retail, template

raw_data, wholesale_prices, retail_prices, output_template = load_data()

# Session state for selections
if "color_selections" not in st.session_state:
    st.session_state.color_selections = []

if "final_selections" not in st.session_state:
    st.session_state.final_selections = []

st.title("Muuto Configurator Matrix Tool")
st.markdown("""
This tool allows you to choose one product family at a time and select specific textile-color combinations via a matrix view. Swatches are shown. Base color selection will follow for products with multiple base options.
""")

# Step 1: Select a product family
families = sorted(raw_data['Product Family'].dropna().unique())
selected_family = st.selectbox("Step 1: Choose a product family", families)

# Filter for selected family
family_data = raw_data[raw_data['Product Family'] == selected_family].copy()
family_data = family_data[family_data['Product Type'].notna() & (family_data['Product Type'] != "N/A")]

# Create product label
def product_label(row):
    label = f"{row['Product Type']} {row['Product Model']}".strip()
    if row['Product Type'] == "Sofa Chaise Longue" and pd.notna(row['Sofa Direction']):
        label += f" - {row['Sofa Direction']}"
    return label

family_data["Product Label"] = family_data.apply(product_label, axis=1)

# Identify textile-family/color combinations
textiles = family_data[["Upholstery Type", "Upholstery Color", "Image URL swatch"]].dropna().drop_duplicates()
textiles["Textile+Color"] = textiles["Upholstery Type"] + " - " + textiles["Upholstery Color"]

# Step 2: Show matrix with swatches
st.subheader("Step 2: Select color combinations")

cols = st.columns([2] + [1 for _ in range(len(textiles))])
cols[0].markdown("**Product**")
for idx, (_, row) in enumerate(textiles.iterrows()):
    cols[idx+1].markdown(f"**{row['Upholstery Color']}**")
    cols[idx+1].image(row["Image URL swatch"], width=40)

product_labels = family_data["Product Label"].unique()
for prod in product_labels:
    row_data = family_data[family_data["Product Label"] == prod]
    cols = st.columns([2] + [1 for _ in range(len(textiles))])
    cols[0].markdown(f"**{prod}**")

    for idx, (_, tex_row) in enumerate(textiles.iterrows()):
        match_rows = row_data[
            (row_data["Upholstery Type"] == tex_row["Upholstery Type"]) &
            (row_data["Upholstery Color"] == tex_row["Upholstery Color"])
        ]
        if not match_rows.empty:
            item_base_preview = match_rows[["Item No", "Base Color"]].dropna()
            if not item_base_preview.empty:
                item_no = item_base_preview["Item No"].values[0]
                checkbox_key = f"{prod}_{tex_row['Upholstery Color']}_{item_no}"
                if st.checkbox("Select", key=checkbox_key):
                    entry = {
                        "Product Label": prod,
                        "Upholstery Type": tex_row["Upholstery Type"],
                        "Upholstery Color": tex_row["Upholstery Color"]
                    }
                    if entry not in st.session_state.color_selections:
                        st.session_state.color_selections.append(entry)

# Clear all selections
if st.button("Clear all selections"):
    st.session_state.color_selections = []
    st.session_state.final_selections = []
    st.experimental_rerun()

# Step 3: Base color selection
if st.session_state.color_selections:
    st.subheader("Step 3: Choose base colors if applicable")
    st.markdown("You can choose multiple base colors for each product+color combination.")

    for i, sel in enumerate(st.session_state.color_selections):
        st.markdown(f"**{sel['Product Label']} – {sel['Upholstery Color']}**")
        match_rows = family_data[
            (family_data["Product Label"] == sel["Product Label"]) &
            (family_data["Upholstery Type"] == sel["Upholstery Type"]) &
            (family_data["Upholstery Color"] == sel["Upholstery Color"])
        ]
        base_colors = match_rows["Base Color"].dropna().unique()
        for base in base_colors:
            item_match = match_rows[match_rows["Base Color"] == base]
            if not item_match.empty:
                item_no = item_match["Item No"].values[0]
                article_no = item_match["Article No"].values[0]
                checkbox_key = f"base_{i}_{item_no}"
                if st.checkbox(f"{base}", key=checkbox_key):
                    final = {"Item No": item_no, "Article No": article_no}
                    if final not in st.session_state.final_selections:
                        st.session_state.final_selections.append(final)

# Step 4: Show final list and download
if st.session_state.final_selections:
    st.subheader("Selected combinations")
    for sel in st.session_state.final_selections:
        st.markdown(f"- **{sel['Item No']}** (Article No: {sel['Article No']})")

    currencies = [c for c in wholesale_prices.columns if c != "Article No."]
    selected_currency = st.selectbox("Step 4: Choose your currency", currencies)

    if st.button("Download masterdata file"):
        export_rows = []
        for sel in st.session_state.final_selections:
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

        final_export = pd.concat(export_rows, ignore_index=True)
        output = BytesIO()
        final_export.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="\ud83d\udcc5 Download masterdata file",
            data=output,
            file_name=f"Muuto_matrix_output_{selected_currency}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
