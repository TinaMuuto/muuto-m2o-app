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
if "selections" not in st.session_state:
    st.session_state.selections = []

st.title("Muuto Configurator Matrix Tool")
st.markdown("""
This tool allows you to choose one product family at a time and select specific textile-color combinations via a matrix view. Base color options are shown where applicable.
""")

# Step 1: Select a product family
families = sorted(raw_data['Product Family'].dropna().unique())
selected_family = st.selectbox("Step 1: Choose a product family", families)

# Filter for selected family
family_data = raw_data[raw_data['Product Family'] == selected_family].copy()
family_data = family_data[family_data['Product Type'].notna() & (family_data['Product Type'] != "N/A")]

# Create readable product labels
def product_label(row):
    label = f"{row['Product Type']} {row['Product Model']}".strip()
    if row['Product Type'] == "Sofa Chaise Longue" and pd.notna(row['Sofa Direction']):
        label += f" - {row['Sofa Direction']}"
    return label

family_data["Product Label"] = family_data.apply(product_label, axis=1)

# Identify textile-family/color combinations
textiles = family_data[["Upholstery Type", "Upholstery Color", "Image URL swatch"]].dropna().drop_duplicates()
textiles["Textile+Color"] = textiles["Upholstery Type"] + " - " + textiles["Upholstery Color"]

# Step 2: Show matrix of product x textile/color
st.subheader("Step 2: Select combinations")

# Matrix header with swatches
cols = st.columns([2] + [1 for _ in range(len(textiles))])
cols[0].markdown("**Product**")
for idx, (_, row) in enumerate(textiles.iterrows()):
    cols[idx+1].image(row["Image URL swatch"], width=40, caption=row["Upholstery Color"])

# Matrix rows
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
            base_colors = match_rows["Base Color"].dropna().unique()
            base_colors = [c for c in base_colors if c != "N/A"]
            if len(base_colors) > 1:
                with cols[idx+1]:
                    matched = match_rows[match_rows["Base Color"] == base_colors[0]]
                    if not matched.empty:
                        item_no = matched["Item No"].values[0]
                        selectbox_key = f"{prod}_{tex_row['Upholstery Color']}_{item_no}"
                        selected = st.selectbox(
                            f"{prod[:12]}-{tex_row['Upholstery Color']}",
                            options=[""] + list(base_colors),
                            key=selectbox_key
                        )
                        if selected:
                            matched = match_rows[match_rows["Base Color"] == selected]
                            if not matched.empty:
                                item_no = matched["Item No"].values[0]
                                article_no = matched["Article No"].values[0]
                                new_selection = {"Item No": item_no, "Article No": article_no}
                                if new_selection not in st.session_state.selections:
                                    st.session_state.selections.append(new_selection)
            elif len(base_colors) == 1:
                matched = match_rows[match_rows["Base Color"] == base_colors[0]]
                if not matched.empty:
                    item_no = matched["Item No"].values[0]
                    article_no = matched["Article No"].values[0]
                    checkbox_key = f"{prod}_{tex_row['Upholstery Color']}_{item_no}"
                    with cols[idx+1]:
                        if st.checkbox(f"{base_colors[0]}", key=checkbox_key):
                            new_selection = {"Item No": item_no, "Article No": article_no}
                            if new_selection not in st.session_state.selections:
                                st.session_state.selections.append(new_selection)

# --- Show selection summary and export ---
if st.session_state.selections:
    st.subheader("Selected combinations")

    for idx, sel in enumerate(st.session_state.selections):
        col1, col2 = st.columns([6, 1])
        with col1:
            st.markdown(f"- **{sel['Item No']}** (Article No: {sel['Article No']})")
        with col2:
            if st.button("\u274C", key=f"remove_{idx}"):
                st.session_state.selections.pop(idx)
                st.experimental_rerun()

    # Step 3: Choose currency
    currencies = [c for c in wholesale_prices.columns if c != "Article No."]
    selected_currency = st.selectbox("Step 3: Choose your currency", currencies)

    # Step 4: Download masterdata file
    if st.button("Download masterdata file with all selections"):
        export_rows = []
        for sel in st.session_state.selections:
            item_no = sel["Item No"]
            article_no = sel["Article No"]

            ws_price = wholesale_prices.loc[
                wholesale_prices["Article No."] == article_no, selected_currency
            ].values
            rt_price = retail_prices.loc[
                retail_prices["Article No."] == article_no, selected_currency
            ].values

            matched_row = raw_data[raw_data["Item No"] == item_no].copy()
            if not matched_row.empty:
                output_row = output_template.copy()
                for col in output_template.columns:
                    if col in matched_row.columns:
                        output_row.loc[0, col] = matched_row.iloc[0][col]
                output_row.loc[0, "Wholesale price"] = ws_price[0] if len(ws_price) else ""
                output_row.loc[0, "Retail Price"] = rt_price[0] if len(rt_price) else ""
                export_rows.append(output_row)

        # Concatenate and export
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
