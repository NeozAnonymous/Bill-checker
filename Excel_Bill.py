import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject
import streamlit as st
import streamlit.components.v1 as components
import io
import base64


def fill_pdf_bytes(csv_bytes: bytes,
                   pdf_bytes: bytes,
                   row_index: int = 0) -> bytes:
    """
    Reads CSV and PDF bytes, fills the PDF form with values from the specified CSV row,
    flattens the form, and returns filled PDF bytes.
    """
    df = pd.read_csv(io.BytesIO(csv_bytes))
    if row_index < 0 or row_index >= len(df):
        raise IndexError(f"Row index {row_index} out of bounds (0 to {len(df) - 1})")
    data = df.iloc[row_index].fillna("").astype(str).to_dict()

    # Read PDF and fill form fields
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    writer.clone_reader_document_root(reader)

    for page in reader.pages:
        writer.add_page(page)

    # Update form fields with data
    for page in writer.pages:
        writer.update_page_form_field_values(page, data)

    # FLATTEN THE FORM (critical addition)
    for page in writer.pages:
        if '/Annots' in page:
            for annot in page['/Annots']:
                writer_annot = annot.get_object()
                writer_annot.update({
                    NameObject("/Ff"): NameObject(1)  # Set field flag to ReadOnly
                })

    # Remove interactive form elements
    if NameObject('/AcroForm') in writer._root_object:
        del writer._root_object[NameObject('/AcroForm')]

    # Write final PDF to buffer
    out_buffer = io.BytesIO()
    writer.write(out_buffer)
    return out_buffer.getvalue()


def display_pdf(pdf_bytes: bytes):
    """
    Embeds PDF bytes by creating a Blob URL and opening it in a new browser tab to avoid data-URI blocking in Chrome/Opera.
    """
    import base64
    # Encode PDF and generate Blob URL client-side
    b64 = base64.b64encode(pdf_bytes).decode('utf-8')
    pdf_html = f"""
    <script>
    (function() {{
        const base64 = "{b64}";
        const binary = atob(base64);
        const len = binary.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
        const blob = new Blob([bytes], {{ type: 'application/pdf' }});
        const url = URL.createObjectURL(blob);
        // Open PDF in new tab
        window.open(url, '_blank');
        // Fallback: display link if pop-up blocked
        const link = document.createElement('a');
        link.href = url;
        link.textContent = 'Click here to view the PDF';
        link.target = '_blank';
        document.body.appendChild(link);
    }})();
    </script>
    """
    components.html(pdf_html, height=100)  # small height for script/link display


def main():
    st.title("📝 Bill checker")
    st.markdown("Upload a CSV and a PDF form template, then choose which row to check.")
    csv_file = st.file_uploader("Upload CSV file", type=["csv"] )
    pdf_file = st.file_uploader("Upload PDF template", type=["pdf"] )

    if csv_file and pdf_file:
        try:
            df = pd.read_csv(csv_file)
            st.success(f"Loaded CSV with {len(df)} rows and {len(df.columns)} columns.")
            st.dataframe(df.head(), height=200)
            row_index = st.number_input(
                label="Select row index for form data",
                min_value=0,
                max_value=len(df) - 1,
                value=0,
                step=1
            )
            if st.button("Generate and Display PDF"):
                try:
                    csv_file.seek(0)
                    pdf_file.seek(0)
                    filled_pdf = fill_pdf_bytes(csv_file.read(), pdf_file.read(), int(row_index))
                    display_pdf(filled_pdf)
                except Exception as e:
                    st.error(f"Error filling PDF: {e}")
        except Exception as e:
            st.error(f"Failed to read CSV: {e}")
    else:
        st.info("Please upload both CSV and PDF files to continue.")

if __name__ == "__main__":
    main()
