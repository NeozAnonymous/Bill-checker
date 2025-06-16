import streamlit as st
import pandas as pd
import docx
from io import BytesIO
import copy
import zipfile
from lxml import etree

# Column mappings for the invoice table
cols_mapping = {
    "STT": 0,
    "Mặt hàng": 1,
    "ĐVT": 3,
    "số Lượng": 4,
    "Đơn giá  ": 5,
    "Doanh số mua chưa có thuế": 6,
}


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
    for para in doc.paragraphs:
        replace_text_in_paragraph(para, replacements)
    for table in doc.tables:
        replace_text_in_table(table, replacements)


def set_cell_text(row, j, text):
    cell = row.cells[j]
    for p in cell.paragraphs:
        if p.runs:
            p.runs[0].text = text
            for run in p.runs[1:]:
                run.text = ''
        else:
            p.add_run(text)


def replace_text_in_textboxes_stream(input_bytes, replacements):
    buffer_in = BytesIO(input_bytes)
    buffer_out = BytesIO()
    with zipfile.ZipFile(buffer_in) as zin:
        with zipfile.ZipFile(buffer_out, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
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


def load_and_process_excel(excel_file):
    excel_data = pd.read_excel(excel_file, header=None)
    invoice_items = excel_data[excel_data.iloc[:, 1].apply(lambda x: str(x).isdigit())]
    col_names = excel_data.ffill().iloc[invoice_items.index.values[0] - 4]
    invoice_items = invoice_items.reset_index(drop=True)
    invoice_items.columns = col_names
    invoice_items = invoice_items.dropna(axis=1)

    if "số Lượng" in invoice_items.columns:
        invoice_items["số Lượng"] = invoice_items["số Lượng"].astype(float).apply(
            lambda x: f"{x:.2f}".replace(".", ","))
    if "Đơn giá  " in invoice_items.columns:
        invoice_items["Đơn giá  "] = invoice_items["Đơn giá  "].apply(lambda x: f"{x:,}".replace(",", "."))
    if "Doanh số mua chưa có thuế" in invoice_items.columns:
        invoice_items["Doanh số mua chưa có thuế"] = invoice_items["Doanh số mua chưa có thuế"].apply(
            lambda x: f"{x:,}".replace(",", "."))

    return invoice_items


def prepare_replacements(invoice_items):
    replacements = {}

    if "Ngày, tháng, năm lập hóa đơn" in invoice_items.columns:
        time = invoice_items["Ngày, tháng, năm lập hóa đơn"].iloc[0]
        day = time.day
        month = f"{time.month:02d}"
        year = time.year
        replacements[
            "Ngày (Date) 22 tháng (month) 04 năm (year) 2025"] = f"Ngày (Date) {day} tháng (month) {month} năm (year) {year}"
        replacements["22/04/2025"] = f"{day}/{month}/{year}"
    else:
        replacements["Ngày (Date) 22 tháng (month) 04 năm (year) 2025"] = ""
        replacements["22/04/2025"] = ""

    if "Số hoá đơn" in invoice_items.columns:
        invoice_number = invoice_items["Số hoá đơn"].iloc[0]
        replacements["00000017"] = f"{invoice_number:08d}"
    else:
        replacements["00000017"] = ""

    if "Mã số thuế người bán" in invoice_items.columns:
        seller_tax_code = invoice_items["Mã số thuế người bán"].iloc[0]
        replacements["0318012656"] = f"{int(seller_tax_code):010}"
    else:
        replacements["0318012656"] = ""

    vat_rates = pd.to_numeric(invoice_items["Thuế suất"], errors='coerce').fillna(0)
    vat_rate = vat_rates.iloc[0]
    replacements["5%"] = f"{vat_rate * 100:,.0f}%"

    return replacements


def populate_document(doc, selected_items, replacements, cols_mapping):
    replace_text_in_doc(doc, replacements)

    table = doc.tables[0]
    table_cols = [col for col in cols_mapping.keys() if col in selected_items.columns]
    missing_table_cols = [cols_mapping[col] for col in cols_mapping.keys() if col not in selected_items.columns]
    table_data = selected_items[table_cols].copy()
    table_data["STT"] = range(1, len(table_data) + 1)
    example_row = table.rows[2]._tr if len(table.rows) > 2 else None

    for i, (_, row_data) in enumerate(table_data.iterrows()):
        if i + 2 < len(table.rows):
            row = table.rows[i + 2]
        else:
            new_row = copy.deepcopy(example_row)
            table._tbl.append(new_row)
            row = table.rows[-1]
        for col in row_data.index:
            value = row_data[col]
            col_index = cols_mapping[col]
            set_cell_text(row, col_index, str(value))

    for r in range(2, len(table.rows) - 4):
        for j in missing_table_cols:
            if j < len(table.rows[r].cells):
                cell = table.rows[r].cells[j]
                for p in cell.paragraphs:
                    p.clear()

    total_rows = len(table_data) + 2
    for r in range(total_rows, len(table.rows) - 4):
        for cell in table.rows[r].cells:
            for p in cell.paragraphs:
                p.clear()

    amounts = pd.to_numeric(selected_items["Doanh số mua chưa có thuế"].str.replace('.', '', regex=False),
                            errors='coerce').fillna(0)
    vat_rates = pd.to_numeric(selected_items["Thuế suất"], errors='coerce').fillna(0)
    vat_amounts = amounts * vat_rates
    vat_total = vat_amounts.sum()
    total_ex_vat = amounts.sum()
    total_inc_vat = total_ex_vat + vat_total

    for idx, value in enumerate([total_ex_vat, vat_total, total_inc_vat]):
        row = table.rows[len(table.rows) - 4 + idx]
        set_cell_text(row, -1, f"{value:,.0f}".replace(",", "."))
    set_cell_text(table.rows[len(table.rows) - 1], -1, "")


# Streamlit UI
st.title("Excel to DOCX Invoice Generator with Textbox Support")
st.write("Upload an Excel file and a DOCX template. Select rows and generate the invoice, including textboxes.")

excel_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])
docx_template = st.file_uploader("Upload DOCX Template", type="docx")

if excel_file and docx_template:
    try:
        invoice_items = load_and_process_excel(excel_file)
        st.subheader("Primary Data Table Preview")
        st.dataframe(invoice_items)

        selected_indices = st.multiselect(
            "Select rows to include",
            options=invoice_items.index.tolist(),
            default=invoice_items.index.tolist()
        )
        selected_items = invoice_items.iloc[selected_indices]
        st.subheader("Rows to be Used")
        st.dataframe(selected_items)

        if st.button("Generate DOCX"):
            template_bytes = docx_template.read()
            replacements = prepare_replacements(invoice_items)
            updated_bytes = replace_text_in_textboxes_stream(template_bytes, replacements)
            doc = docx.Document(updated_bytes)
            populate_document(doc, selected_items, replacements, cols_mapping)
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
