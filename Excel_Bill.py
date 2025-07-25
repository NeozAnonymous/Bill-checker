import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io


def parse_num(x):
    try:
        return int(x)
    except:
        return float(x)

def extract_invoice_info(tree):
    """
    Parse a Vietnamese VAT invoice XML ElementTree and extract key information.
    Returns a dict with header, parties, line items, and totals.
    """
    root = tree.getroot()

    # Header

    hdon = root.find('.//DLHDon')
    chung = hdon.find('.//TTChung')
    invoice_series = chung.findtext('KHHDon')
    invoice_number = chung.findtext('SHDon')
    invoice_date = chung.findtext('NLap').split("-")
    invoice_date = f"{invoice_date[2]}/{invoice_date[1]}/{invoice_date[0]}"
    invoice_exchange_rate = parse_num(chung.findtext('TGia') or 1)

    # Parties
    seller_el = root.find('.//NBan')
    seller = {
        'name': seller_el.findtext('Ten'),
        'tax_code': seller_el.findtext('MST'),
        'address': seller_el.findtext('DChi')
    }
    buyer_el = root.find('.//NMua')
    buyer = {
        'name': buyer_el.findtext('Ten'),
        'tax_code': buyer_el.findtext('MST'),
        'address': buyer_el.findtext('DChi')
    }

    # Line items
    items = []
    for line in root.findall('.//DSHHDVu/HHDVu'):
        tax_amount = 0.0
        for tt in line.findall('TTKhac/TTin'):
            if tt.findtext('TTruong') == 'VATAmount':
                tax_amount = float(tt.findtext('DLieu') or 0)
        items.append({
            'description': line.findtext('THHDVu'),
            'quantity': parse_num(line.findtext('SLuong') or 0),
            'unit': line.findtext('DVTinh'),
            'unit_price': parse_num(line.findtext('DGia') or 0),
            'line_total': parse_num(line.findtext('ThTien') or 0),
            'tax_rate': line.findtext('TSuat'),
            'tax_amount': tax_amount
        })

    total = root.find('.//TToan')
    total_vat = parse_num(total.findtext('TgTThue'))

    return {
        'header': {
            'series': invoice_series,
            'number': invoice_number,
            'date': invoice_date,
            'exchange_rate': invoice_exchange_rate
        },
        'seller': seller,
        'buyer': buyer,
        'items': items,
        'total_vat': total_vat,
    }


# Streamlit App
st.set_page_config(page_title="Invoice XML to Excel", layout="wide")
st.title("Invoice XML to Excel Exporter")

# Allow multiple file uploads
uploaded_files = st.file_uploader(
    "Upload one or more invoice XML files", type="xml", accept_multiple_files=True
)

if uploaded_files:
    all_rows = []

    for uploaded_file in uploaded_files:
        try:
            tree = ET.parse(uploaded_file)
            info = extract_invoice_info(tree)

            # Build rows for this invoice
            for idx, item in enumerate(info['items'], start=1):
                total_line = item['line_total'] + item['tax_amount']
                all_rows.append({
                    'STT': idx,
                    'NGÀY CHỨNG TỪ': info['header']['date'],
                    'SỐ CHỨNG TỪ': "",
                    'NGÀY HÓA ĐƠN': info['header']['date'],
                    'SỐ HÓA ĐƠN': info['header']['number'],
                    'TÊN MẶT HÀNG': item['description'],
                    'SỐ LƯỢNG': item['quantity'],
                    'ĐƠN VỊ': item['unit'],
                    'GIÁ': item['unit_price'],
                    'TỶ GIÁ': info['header']['exchange_rate'],
                    'THÀNH TIỀN NGUYÊN TỆ': item['line_total'],
                    'THÀNHTIỀN(VND)': item['line_total'] * info['header']['exchange_rate'],
                    'NỢ': "",
                    'CÓ': "",
                    'HẠNG MỤC': "",
                    'TÊN NGƯỜI BÁN': info['seller']['name'],
                    'MÃ SỐ THUẾ NGƯỜI BÁN': info['seller']['tax_code'],
                })
            all_rows.append({
                'STT': idx+1,
                'NGÀY CHỨNG TỪ': info['header']['date'],
                'SỐ CHỨNG TỪ': "",
                'NGÀY HÓA ĐƠN': info['header']['date'],
                'SỐ HÓA ĐƠN': info['header']['number'],
                'TÊN MẶT HÀNG': "THUẾ GTGT",
                'SỐ LƯỢNG': "",
                'ĐƠN VỊ': "",
                'GIÁ': "",
                'TỶ GIÁ': info['header']['exchange_rate'],
                'THÀNH TIỀN NGUYÊN TỆ': info["total_vat"],
                'THÀNHTIỀN(VND)': info["total_vat"] * info['header']['exchange_rate'],
                'NỢ': "",
                'CÓ': "",
                'HẠNG MỤC': "",
                'TÊN NGƯỜI BÁN': info['seller']['name'],
                'MÃ SỐ THUẾ NGƯỜI BÁN': info['seller']['tax_code'],
            })
        except Exception as e:
            st.error(f"Failed to parse {uploaded_file.name}: {e}")

    total_1 = sum([x["THÀNH TIỀN NGUYÊN TỆ"] for x in all_rows])
    total_2 = sum([x["THÀNHTIỀN(VND)"] for x in all_rows])

    all_rows.append({k:"" for k in all_rows[0].keys()})

    d = {k: "" for k in all_rows[0].keys()}
    d["THÀNH TIỀN NGUYÊN TỆ"] = total_1
    d["THÀNHTIỀN(VND)"] = total_2
    all_rows.append(d)

    # Create DataFrame with all rows
    df_export = pd.DataFrame(all_rows)
    df_export["STT"] = list(range(1, len(df_export) - 1)) + ["", ""]

    # Show preview
    st.subheader("Preview of Excel Format for All Invoices")
    st.dataframe(df_export)

    # Generate Excel in-memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Invoices')
    buffer.seek(0)

    # Download button
    st.download_button(
        label="Download all invoices as Excel",
        data=buffer,
        file_name="invoices_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Please upload at least one XML invoice to convert to Excel.")
