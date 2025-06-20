import streamlit as st
import docx
import pandas as pd
import re
from io import BytesIO
import openpyxl
import copy
from datetime import datetime

excel_cols = [
    'STT', 'Số hóa đơn', "Ngày, tháng, năm lập hóa đơn", "Tên người bán", "Mã số thuế người bán", "Mặt hàng",
    "ĐVT", "số Lượng", "Đơn giá  ", "Doanh số mua chưa có thuế", "Thuế suất", "Thuế GTGT", "Ghi chú",
]
col_map = {
    1: "STT",
    2: "Ký hiệu mẫu hóa đơn",
    3: "Ký hiệu hoá đơn",
    4: "Số hoá đơn",
    5: "Ngày, tháng, năm lập hóa đơn",
    6: "Tên người bán",
    7: "Mã số thuế người bán",
    8: "Mặt hàng",
    9: "ĐVT",
    10: "số Lượng",
    11: "Đơn giá",
    12: "Doanh số mua chưa có thuế",
    13: "Thuế suất",
    14: "Thuế GTGT",
    15: "Ghi chú"
}

def get_writable_cell(ws, row, column):

    cell_coord = ws.cell(row=row, column=column).coordinate
    for merged_range in ws.merged_cells.ranges:
        if cell_coord in merged_range:
            # Return the top-left cell of the merged range
            return ws[merged_range.min_row][merged_range.min_col - 1]
    # If not merged, return the original cell
    return ws.cell(row=row, column=column)

def extract_from_docx(file):
    doc = docx.Document(file)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip() != '']

    # Find the index of invoice title
    for i, p in enumerate(paragraphs):
        if "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in p or "VAT INVOICE" in p or p in "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" or p in "VAT INVOICE":
            idx = i
            break
    else:
        raise ValueError("Cannot find invoice title")

    # Extract seller name (first paragraph)
    seller_name = ""

    # Extract tax code
    tax_code = None
    for p in paragraphs[:]:
        if "Mã số thuế" in p or "Tax code" in p:
            tax_code = p.split(":")[-1].strip().replace("-", "").replace(" ", "")
            break
    if tax_code is None:
        tax_code = ''

    # Extract date, serial, and number
    date = serial = number = None
    for p in paragraphs[idx + 1:]:
        if date is None:
            match = re.search(r"Ngày(?:\s*\([^)]*\))?\s+(\d{1,2})\s+tháng(?:\s*\([^)]*\))?\s+(\d{1,2})\s+năm(?:\s*\([^)]*\))?\s+(\d{4})", p)
            if match:
                day = match.group(1)
                month = match.group(2)
                year = match.group(3)
                date = f"{day}/{month}/{year}"
        if serial is None and ("Ký hiệu" in p or "Serial" in p):
            serial = p.split(":")[1].strip()
        if number is None and ("Số" in p or "Invoice No." in p):
            number = p.split(":")[1].strip()
        if date and serial and number:
            break
    if date is None:
        date = ''
    if serial is None:
        serial = ''
    if number is None:
        number = ''

    # Extract first table
    table = doc.tables[0]
    item_rows = []
    start_idx = 1
    row = table.rows[2]
    cells = row.cells
    if cells[0].text.strip()=="1":
        start_idx = 2
    else:
        start_idx = 1
    for row in table.rows[start_idx:]:
        cells = row.cells
        if cells[0].text.strip().isdigit():
            cols_used = []
            j = 0
            prior_tc = None
            for cell in cells:
                this_tc = cell._tc
                if this_tc is prior_tc:
                    j += 1
                    continue
                cols_used.append(cells[j].text.strip())
                j += 1
                prior_tc = this_tc
            item_rows.append(cols_used)
        else:
            break

    # Extract tax rate
    tax_rate = None
    for cell in table.rows[-3].cells:
        match = re.search(r"\d{1,3}%", cell.text)
        if match:
            tax_rate = float(match.group(0).replace("%", "")) / 100
            break
    if tax_rate is None:
        tax_rate = ''

    # Process item rows
    data = []
    for row in item_rows:
        stt, ten_hang_hoa, don_vi_tinh, so_luong_str, don_gia_str, thanh_tien_str = row[:6]

        # Parse quantities and prices
        so_luong = int(so_luong_str.replace('.', '').split(',')[0]) if ',' in so_luong_str else int(so_luong_str.replace('.', ''))
        don_gia = int(don_gia_str.replace('.', '').replace(',', ''))
        thanh_tien = int(thanh_tien_str.replace('.', '').replace(',', ''))
        thue_gtgt = thanh_tien * tax_rate

        data.append({
            'STT': stt,
            'Tên người bán': seller_name,
            'Mã số thuế người bán': tax_code,
            'Mặt hàng': ten_hang_hoa,
            'ĐVT': don_vi_tinh,
            'số Lượng': so_luong,
            'Đơn giá': don_gia,
            'Doanh số mua chưa có thuế': thanh_tien,
            'Thuế suất': tax_rate,
            'Thuế GTGT': thue_gtgt,
            'Ghi chú': '',
            'Ký hiệu mẫu hóa đơn': '',
            'Ký hiệu hoá đơn': serial,
            'Số hoá đơn': number,
            'Ngày, tháng, năm lập hóa đơn': date
        })
    return data

def process_files(docx_files, excel_file):
    all_data = []
    for docx_file in docx_files:
        all_data.extend(extract_from_docx(docx_file))

    # Sort by invoice date
    all_data.sort(key=lambda x: datetime.strptime(x['Ngày, tháng, năm lập hóa đơn'], "%d/%m/%Y"))

    c = 1
    for i in range(len(all_data)):
        all_data[i]["STT"] = c
        c+=1

    new_df = pd.DataFrame(all_data)
    st.write("### Extracted Data from DOCX Files")
    st.dataframe(new_df)

    # Load workbook with openpyxl
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active

    # Find insert_idx
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        val = row[1]  # column B
        if val is not None and str(val).strip().isdigit():
            insert_idx = row_idx
            break
    else:
        raise ValueError("Cannot find starting row")

    # Store styles from the first data row (row insert_idx) for columns 1 to 16
    # Before applying styles, store them with copy()
    column_styles = []
    for col in range(1, 17):  # Adjust range as per your columns
        source_cell = ws.cell(row=insert_idx, column=col)
        column_styles.append({
            'font': copy.copy(source_cell.font),
            'border': copy.copy(source_cell.border),
            'fill': copy.copy(source_cell.fill),
            'number_format': copy.copy(source_cell.number_format),
            'protection': copy.copy(source_cell.protection),
            'alignment': copy.copy(source_cell.alignment)
        })

    # Find end_idx
    end_idx = insert_idx
    while True:
        val_b = ws.cell(row=end_idx, column=2).value
        if not (isinstance(val_b, int) or val_b is None):
            break
        end_idx += 1
        if end_idx > ws.max_row:
            break

    # Delete existing data rows
    num_delete = end_idx - insert_idx
    if num_delete > 0:
        ws.delete_rows(insert_idx, amount=num_delete)

    # Insert new rows
    for cell in ws[insert_idx]:
        print(cell.value)
    num_new_rows = len(all_data)
    ws.insert_rows(insert_idx, amount=num_new_rows)

    # Set values and styles for new rows
    for i, item in enumerate(all_data):
        row_idx = insert_idx + i
        row = ['']
        for key in range(1, 16):
            col_name = col_map[key]
            value = item.get(col_name, '')
            row.append(value)
        for j, value in enumerate(row):
            col = j + 1  # Columns 1 to 16
            target_cell = get_writable_cell(ws, row_idx, col)
            style = column_styles[j]
            target_cell.font = style['font']
            target_cell.border = style['border']
            target_cell.fill = style['fill']
            target_cell.number_format = style['number_format']
            target_cell.protection = style['protection']
            target_cell.alignment = style['alignment']
            target_cell.value = value

    # Update summary rows
    row_idx = insert_idx + len(all_data)
    total_doanh_so = sum(item.get('Doanh số mua chưa có thuế', 0) for item in all_data)
    total_thue_gtgt = sum(item.get('Thuế GTGT', 0) for item in all_data)
    # Use get_writable_cell for robustness
    get_writable_cell(ws, row_idx, 13).value = total_doanh_so
    get_writable_cell(ws, row_idx, 15).value = total_thue_gtgt

    row_idx += 5
    get_writable_cell(ws, row_idx, 8).value = total_doanh_so

    row_idx += 1
    get_writable_cell(ws, row_idx, 8).value = total_thue_gtgt

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("Invoice Data Processor")
    docx_files = st.file_uploader("Upload DOCX Files", type="docx", accept_multiple_files=True)
    excel_file = st.file_uploader("Upload Excel File", type="xlsx", accept_multiple_files=False)

    if docx_files and excel_file:
        if st.button("Process Files"):
            result_excel = process_files(docx_files, excel_file)
            st.success("Processing complete!")
            st.download_button(
                label="Download Result Excel",
                data=result_excel,
                file_name="result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
