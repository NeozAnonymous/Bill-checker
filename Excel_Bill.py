import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import copy

# Fixed buyer information
buyer_info = {
    "buyer_name": "CÔNG TY TNHH MAI KA",
    "buyer_tax_code": "3700769325",
    "buyer_address": "Số 10, Đường Đồng Minh, Khu phố Bình Minh 1, Phường Dĩ An, Thành phố Dĩ An, Tỉnh Bình Dương, Việt Nam"
}

# Columns to select from Excel
chosen_cols = [1, 8, 9, 10, 11, 12]

# State for DOCX cell handling
prior_tc = None
shift = 0

def replace_text_in_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

def replace_text_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_text_in_paragraph(para, replacements)

def replace_text_in_doc(doc, replacements):
    for para in doc.paragraphs:
        replace_text_in_paragraph(para, replacements)
    for table in doc.tables:
        replace_text_in_table(table, replacements)

def set_cell_text(row, j, text):
    global prior_tc, shift
    cell = row.cells[j + shift]
    this_tc = cell._tc
    if this_tc is prior_tc:
        shift = 1
        cell = row.cells[j + shift]
    for p in cell.paragraphs:
        if p.runs:
            p.runs[0].text = text
            for run in p.runs[1:]:
                run.text = ''
        else:
            p.add_run(text)
    prior_tc = this_tc

# Streamlit UI
st.title("Excel to DOCX Invoice Generator with Row Selection")
st.write("Upload an Excel file and a DOCX template. Select the rows you want to include before generating the invoice.")

excel_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])
docx_template = st.file_uploader("Upload DOCX Template", type="docx")

if excel_file and docx_template:
    try:
        # Read the Excel without header
        df = pd.read_excel(excel_file, header=None)
        # Extract only primary table rows: second column must be integer-like
        primary_df = df[df.iloc[:, 1].apply(lambda x: str(x).isdigit())]
        primary_df = primary_df.iloc[:, chosen_cols].reset_index(drop=True)

        st.subheader("Primary Data Table Preview")
        st.dataframe(primary_df)

        # Let user pick rows by index
        selected_indices = st.multiselect(
            "Select rows to include",
            options=primary_df.index.tolist(),
            default=primary_df.index.tolist()
        )
        filtered_df = primary_df.loc[selected_indices]

        st.subheader("Rows to be Used")
        st.dataframe(filtered_df)

        if st.button("Generate DOCX"):
            # Load DOCX template
            doc = Document(docx_template)

            # Replace buyer information
            replacements = {
                "{buyer_name}": buyer_info["buyer_name"],
                "{buyer_tax_code}": buyer_info["buyer_tax_code"],
                "{buyer_address}": buyer_info["buyer_address"],
            }
            replace_text_in_doc(doc, replacements)

            # Populate the first table in the template
            table = doc.tables[0]
            example_tr = table.rows[2]._tr if len(table.rows) > 2 else None

            # Fill or append rows
            for i, row_data in enumerate(filtered_df.itertuples(index=False)):
                if i + 2 < len(table.rows):
                    row = table.rows[i + 2]
                else:
                    if example_tr:
                        new_tr = copy.deepcopy(example_tr)
                        table._tbl.append(new_tr)
                        row = table.rows[-1]
                    else:
                        row = table.add_row()
                for j, value in enumerate(row_data):
                    set_cell_text(row, j, str(value))

            # Clear any surplus rows beyond the used range
            total_used = len(filtered_df) + 2
            for r in range(total_used, len(table.rows)-4):
                for cell in table.rows[r].cells:
                    for p in cell.paragraphs:
                        p.clear()

            # Perform calculations
            # Assume last chosen column is Amount
            amounts = pd.to_numeric(filtered_df.iloc[:, -1], errors='coerce').fillna(0)
            total_ex_vat = amounts.sum()
            vat_rate = 0.05  # 5%
            vat_amount = total_ex_vat * vat_rate
            total_inc_vat = total_ex_vat + vat_amount

            for r in range(len(table.rows)-4, len(table.rows)):
                row = table.rows[r]
                if r==len(table.rows)-4:
                    set_cell_text(row, -2, f"{total_ex_vat:,.0f}")
                if r==len(table.rows)-3:
                    set_cell_text(row, -2, f"{vat_amount:,.0f}")
                if r==len(table.rows)-2:
                    set_cell_text(row, -2, f"{total_inc_vat:,.0f}")
                if r==len(table.rows)-1:
                    for cell in table.rows[r].cells:
                        for p in cell.paragraphs:
                            p.clear()

            # Save to BytesIO and provide download
            output = BytesIO()
            doc.save(output)
            output.seek(0)

            st.download_button(
                label="Download DOCX",
                data=output,
                file_name="invoice.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload both an Excel file and a DOCX template to proceed.")
