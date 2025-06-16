import streamlit as st
import pandas as pd
import docx  # use python-docx
from io import BytesIO
import copy
import zipfile
from lxml import etree

# Fixed buyer information
buyer_info = {
    "buyer_name": "CÔNG TY TNHH MAI KA",
    "buyer_tax_code": "3700769325",
    "buyer_address": "Số 10, Đường Đồng Minh, Khu phố Bình Minh 1, Phường Dĩ An, Thành phố Dĩ An, Tỉnh Bình Dương, Việt Nam"
}

# Columns to select from Excel
table_cols = [0, 5, 6, 7, 8, 9]

# State for DOCX cell handling
prior_tc = None
shift = 0

def replace_text_in_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            if paragraph.text == key:
                if paragraph.runs:
                    paragraph.runs[0].text = value
                    for run in paragraph.runs[1:]:
                        run.text = ''
                else:
                    paragraph.text = value
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)


def replace_text_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_text_in_paragraph(para, replacements)


def replace_text_in_doc(doc, replacements):
    # replaces in normal paragraphs & tables
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


def replace_text_in_textboxes_stream(input_bytes, replacements):
    """
    Reads a DOCX from input_bytes (BytesIO or bytes), replaces text in all textboxes, and returns new BytesIO
    """
    buffer_in = BytesIO(input_bytes)
    buffer_out = BytesIO()
    with zipfile.ZipFile(buffer_in) as zin:
        with zipfile.ZipFile(buffer_out, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    # parse and replace in textboxes
                    tree = etree.fromstring(data)
                    ns = {
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
                        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                    }
                    for tb in tree.findall('.//w:txbxContent', ns):
                        for t in tb.xpath('.//w:t', namespaces=ns):
                            if t.text:
                                for old, new in replacements.items():
                                    if old in t.text:
                                        t.text = t.text.replace(old, new)
                    new_xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone='yes')
                    zout.writestr(item, new_xml)
                else:
                    zout.writestr(item, data)
    buffer_out.seek(0)
    return buffer_out

# Streamlit UI
st.title("Excel to DOCX Invoice Generator with Textbox Support")
st.write("Upload an Excel file and a DOCX template. Select rows and generate the invoice, including textboxes.")

excel_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])
docx_template = st.file_uploader("Upload DOCX Template", type="docx")

if excel_file and docx_template:
    try:
        # Read the Excel without header
        df = pd.read_excel(excel_file, header=None)
        primary_df = df[df.iloc[:, 1].apply(lambda x: str(x).isdigit())].reset_index(drop=True)
        primary_df = primary_df.dropna(axis=1)
        primary_df.columns = range(len(primary_df.columns))

        primary_df[7] = primary_df[7].astype(float).apply(lambda x: f"{x:.2f}".replace(".", ","))
        primary_df[8] = primary_df[8].apply(lambda x: f"{x:,}".replace(",", "."))
        primary_df[9] = primary_df[9].apply(lambda x: f"{x:,}".replace(",", "."))

        st.subheader("Primary Data Table Preview")
        st.dataframe(primary_df)

        selected_indices = st.multiselect(
            "Select rows to include",
            options=primary_df.index.tolist(),
            default=primary_df.index.tolist()
        )
        filtered_df = primary_df.iloc[selected_indices]

        st.subheader("Rows to be Used")
        st.dataframe(filtered_df)

        if st.button("Generate DOCX"):
            # Read template into bytes
            template_bytes = docx_template.read()
            # Replacement mapping for textboxes and normal body
            day = primary_df[2].iloc[0].day
            month = f"{primary_df[2].iloc[0].month:02d}"
            year = primary_df[2].iloc[0].year

            No = primary_df[1].iloc[0]
            Seller_tax_code = primary_df[4].iloc[0]
            vats = pd.to_numeric(primary_df.iloc[selected_indices][10], errors='coerce').fillna(0)
            vat = vats.iloc[0]

            rep_str = f"Ngày (Date) {day} tháng (month) {month} năm (year) {year}"
            replacements = {
                "Ngày (Date) 22 tháng (month) 04 năm (year) 2025": rep_str,
                "00000017" : f"{No:08d}",
                "22/04/2025": f"{day}/{month}/{year}",
                "0318012656": f"{Seller_tax_code:010}",
                "5%" : f"{vat*100:,.0f}%",
            }

            # First replace text in textboxes
            boxed_bytes = replace_text_in_textboxes_stream(template_bytes, replacements)
            # Then load into python-docx for tables & paragraphs
            doc = docx.Document(boxed_bytes)

            # Also replace in body
            replace_text_in_doc(doc, replacements)

            # Populate table rows
            table = doc.tables[0]
            filtered_df = filtered_df[table_cols]
            example_tr = table.rows[2]._tr if len(table.rows) > 2 else None
            for i, row_data in enumerate(filtered_df.itertuples(index=False)):
                if i + 2 < len(table.rows):
                    row = table.rows[i + 2]
                else:
                    new_tr = copy.deepcopy(example_tr)
                    table._tbl.append(new_tr)
                    row = table.rows[-1]
                for j, value in enumerate(row_data):
                    set_cell_text(row, j, str(value))

            # Clear surplus and calculate totals
            total_rows = len(filtered_df) + 2
            for r in range(total_rows, len(table.rows)-4):
                for cell in table.rows[r].cells:
                    for p in cell.paragraphs:
                        p.clear()

            amounts = pd.to_numeric(filtered_df[9].str.replace('.', '', regex=False), errors='coerce').fillna(0)
            vat_amounts = amounts * vats
            vat_amount = vat_amounts.sum()
            total_ex_vat = amounts.sum()
            total_inc_vat = total_ex_vat + vat_amount

            for idx, label_val in enumerate([total_ex_vat, vat_amount, total_inc_vat]):
                row = table.rows[len(table.rows)-4 + idx]
                set_cell_text(row, -2, f"{label_val:,.0f}".replace(",", "."))
            set_cell_text(table.rows[len(table.rows)-1], -1, "")

            # Save and offer download
            output = BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button(
                label="Download Updated Invoice",
                data=output,
                file_name="invoice.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload both an Excel file and a DOCX template to proceed.")

# Note: ensure environment has python-docx, lxml installed
# pip uninstall docx
# pip install python-docx lxml
