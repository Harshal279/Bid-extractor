import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import re
import io
import requests
from PyPDF2 import PdfReader

st.set_page_config(page_title="Bid Extractor", layout="centered")
st.title("Bid Extractor")
st.markdown("Upload PDF/Image")


pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def get_zoho_access_token():
    url = "https://accounts.zoho.in/oauth/v2/token"
    params = {
        "refresh_token": st.secrets["zoho"]["refresh_token"],
        "client_id": st.secrets["zoho"]["client_id"],
        "client_secret": st.secrets["zoho"]["client_secret"],
        "grant_type": "refresh_token"
    }
    response = requests.post(url, params=params)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        raise Exception(f"Token fetch failed: {response.text}")

def extract_text_from_pdf_bytes(pdf_bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n\n"
    return text

def extract_text_from_image(image):
    return pytesseract.image_to_string(image, config='--psm 6')

def parse_bid_table(full_text):
    lines = [line.strip() for line in full_text.splitlines() if line.strip()]
    
    rfx_number = "UNKNOWN"
    for line in lines:
        if re.search(r'RFx\s*number|RFx\s*No', line, re.IGNORECASE):
            match = re.search(r'\d{10,}', line)
            if match:
                rfx_number = match.group(0)
                break

    start_idx = 0
    for i, line in enumerate(lines):
        if "Bid Details" in line:
            start_idx = i + 1  
            break

    rows = []
    current = None

    for line in lines[start_idx:]:
        if any(h in line.upper() for h in ["ITEM", "MATERIAL NO.", "DESCRIPTION", "QTY/UNIT"]):
            continue

        item_match = re.match(r'^(\d{1,4})\s+', line)
        if item_match:
            if current:
                rows.append(current)

            item_no = item_match.group(1)
            rest = line[item_match.end():].strip()

            qty_match = re.search(r'(\d+\s+NO)$', rest, re.IGNORECASE)
            if qty_match:
                qty = qty_match.group(1).upper()
                rest_no_qty = rest[:qty_match.start()].strip()
            else:
                qty = ""
                rest_no_qty = rest

            mat_match = re.match(r'(\d{10,})\s+(.*)', rest_no_qty)
            if mat_match:
                material = mat_match.group(1)
                description = mat_match.group(2)
            else:
                material = ""
                description = rest_no_qty

            current = {
                "RFx_Number": rfx_number,
                "Item_No": item_no,
                "Material_No": material,
                "Description": description,
                "Qty_Unit": qty
            }
        else:
            if current and line.strip():
                cleaned = line.strip()

                if re.search(r'[A-Z0-9\-& ,.]+', cleaned) and any(kw in cleaned.upper() for kw in ["SOURIAU", "RADIALL", "& CIE", "PRIVATE LIMITED", "-847"]):
                    if current["Material_No"]:
                        current["Material_No"] += " / " + cleaned
                    else:
                        current["Material_No"] = cleaned
                else:
                    
                    current["Description"] += " " + cleaned

    if current:
        rows.append(current)

    return rows

uploaded_file = st.file_uploader("Upload Bid(PDF or Image)", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    if uploaded_file.type == "application/pdf":
        pdf_bytes = uploaded_file.getvalue()
        st.success("PDF uploaded successfully")
        text = extract_text_from_pdf_bytes(pdf_bytes)
    else:
        image = Image.open(uploaded_file)
        st.image(image, caption="Uploaded Image", use_column_width=True)
        text = extract_text_from_image(image)

    with st.expander("Extracted Raw Text (OCR)"):
        st.text_area("Raw Text", text, height=300)

    rows = parse_bid_table(text)

    if rows:
        df = pd.DataFrame(rows)
        rfx = rows[0]["RFx_Number"]
        st.success(f"Extracted {len(rows)} items | RFx number: {rfx}")
        st.dataframe(df[['Item_No', 'Material_No', 'Description', 'Qty_Unit']], use_container_width=True)

        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Extracted Data as CSV", csv, f"BEL_RFx_{rfx}.csv", "text/csv")


        st.markdown("### Upload to Zoho CRM")
        st.info("Creates 1 record in **Bid Details** with all line items in the **Bid Items** subform.")

        if st.button("Upload to Zoho CRM", type="primary"):
            with st.spinner("Uploading to Zoho CRM..."):
                try:
                    token = get_zoho_access_token()
                    headers = {
                        "Authorization": f"Zoho-oauthtoken {token}",
                        "Content-Type": "application/json"
                    }

                    subform_items = [
                        {
                            "Item": row["Item_No"],
                            "Material_No": row["Material_No"],
                            "Description": row["Description"],
                            "Qty_Unit": row["Qty_Unit"]
                        }
                        for row in rows
                    ]

                    payload = {
                        "data": [{
                            "Name": rfx,                       
                            "Bid_Items": subform_items
                        }]
                    }

                    url = "https://www.zohoapis.in/crm/v6/Bid_Details"

                    response = requests.post(url, headers=headers, json=payload)

                    if response.status_code == 201:
                        record_id = response.json()["data"][0]["details"]["id"]
                        st.success("Record created in Zoho CRM")
                        st.balloons()
                        st.write(f"**RFx number**: {rfx}")
                        st.write(f"**Record ID**: {record_id}")
                        st.write(f"**Items added**: {len(rows)} in Bid Items subform")
                        st.info("Check Zoho CRM → Bid Details to see the new record!")
                    else:
                        st.error("Upload failed")
                        st.code(response.text)

                except Exception as e:
                    st.error(f"Error: {str(e)}")

    else:
        st.warning("No items detected. Try a higher quality scan/PDF.")

else:
    st.info("Upload bid document to get started.")
