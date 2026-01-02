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

def parse_bid_info(full_text):
    # Remove 
    full_text = re.sub(r'\s*\|\s*', ' ', full_text)
    lines = [line.strip() for line in full_text.splitlines() if line.strip()]

    # Bid No
    bid_no = "UNKNOWN"
    for line in lines:
        if "RFx number" in line:
            match = re.search(r'\d{10,}', line)
            if match:
                bid_no = match.group(0)
                break

    # Customer Name
    customer_name = "Unknown"
    company_found = False
    for line in lines:
        if "Company" in line:
            company_found = True
            continue
        if company_found and line and not line.startswith(("C-", "Information", "Description", "RFx")):
            # The line after "Company" is usually the customer name (ASEEM ELECTRONICS)
            customer_name = line.upper()  # Often in caps
            break

    # Closing Date
    closing_date = "Unknown"
    for line in lines:
        if "Submission period:" in line:
            match = re.search(r'-\s*(\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2})', line)
            if match:
                closing_date = match.group(1)
            break

    # Bid Email
    bid_email = "Unknown"
    for line in lines:
        match = re.search(r'[\w\.-]+@bel\.co\.in', line)
        if match:
            bid_email = match.group(0)
            break

    # Find table start
    start_idx = 0
    for i, line in enumerate(lines):
        if "Bid Details" in line:
            # Header usually next
            if i + 1 < len(lines) and "Item" in lines[i+1]:
                start_idx = i + 2
            break

    rows = []
    current_item = None
    vendor_lines = []

    for line in lines[start_idx:]:
        # Skip unwanted
        if any(kw in line for kw in ["MSME", "@gmail.com", "@aseemelectronics.com"]):
            continue

        # Item line
        item_match = re.match(r'^(\d{1,4})\s+(\d{10,})\s+(.*?)\s+(\d+(?:\.\d+)? \s*[A-Z]+)$', line)
        if item_match:
            # Save previous
            if current_item:
                for v_line in vendor_lines:
                    if '-' in v_line:
                        parts = v_line.split('-', 1)
                        man = parts[0].strip()
                        mpn = parts[1].strip() if len(parts) > 1 else ""
                        rows.append({
                            "Item": current_item["Item"],
                            "Customer_Part_No": current_item["Customer_Part_No"],
                            "Item_Description": current_item["Item_Description"],
                            "Quantity": current_item["Quantity"],
                            "Manufacturer": man,
                            "MPN": mpn
                        })
                if not vendor_lines:
                    rows.append(current_item)

            current_item = {
                "Item": item_match.group(1),
                "Customer_Part_No": item_match.group(2),
                "Item_Description": item_match.group(3).strip(),
                "Quantity": item_match.group(4),
                "Manufacturer": "",
                "MPN": ""
            }
            vendor_lines = []
        else:
            if current_item:
                vendor_lines.append(line)

    # Save last
    if current_item:
        for v_line in vendor_lines:
            if '-' in v_line:
                parts = v_line.split('-', 1)
                man = parts[0].strip()
                mpn = parts[1].strip() if len(parts) > 1 else ""
                rows.append({
                    "Item": current_item["Item"],
                    "Customer_Part_No": current_item["Customer_Part_No"],
                    "Item_Description": current_item["Item_Description"],
                    "Quantity": current_item["Quantity"],
                    "Manufacturer": man,
                    "MPN": mpn
                })
        if not vendor_lines:
            rows.append(current_item)

    return bid_no, customer_name, closing_date, bid_email, rows

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

    '''with st.expander("Extracted Raw Text (OCR)"):
        st.text_area("Raw Text", text, height=300)'''

    bid_no, customer_name, closing_date, bid_email, rows = parse_bid_info(text)

    if rows:
        df = pd.DataFrame(rows)
        st.success(f"Extracted {len(rows)} items | Bid No: {bid_no}")
        st.dataframe(df[['Item','Customer_Part_No', 'Item_Description', 'Manufacturer', 'MPN', 'Quantity']], use_container_width=True)

        '''csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Extracted Data as CSV", csv, f"Bid_{bid_no}.csv", "text/csv")'''

        st.markdown("### Upload to Zoho CRM")
        '''st.info("Creates 1 record")'''

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
                            "Item":row["Item"],
                            "Material_No": row["Customer_Part_No"],
                            "Description": row["Item_Description"],
                            "manufacturer": row["Manufacturer"],
                            "MPN": row["MPN"],
                            "Qty_Unit": row["Quantity"]
                        }
                        for row in rows
                    ]

                    payload = {
                        "data": [{
                            "Name": bid_no,                       
                            "Customer_Name": customer_name,
                            "Closing_Date": closing_date,
                            "Bid_email": bid_email,
                            "Bid_Items": subform_items
                        }]
                    }

                    url = "https://www.zohoapis.in/crm/v6/Bid_Details"

                    response = requests.post(url, headers=headers, json=payload)

                    if response.status_code == 201:
                        record_id = response.json()["data"][0]["details"]["id"]
                        st.success("Record created in Zoho CRM")
                        st.write(f"**Bid No**: {bid_no}")
                        st.write(f"**Record ID**: {record_id}")
                        st.write(f"**Items added**: {len(rows)} in Bid Items subform")
                        st.info("Check Zoho CRM → Bid Details Updated!")
                    else:
                        st.error("Upload failed")
                        st.code(response.text)

                except Exception as e:
                    st.error(f"Error: {str(e)}")

    else:
        st.warning("No items detected. Try a higher quality scan/PDF.")

else:
    st.info("Upload bid document")

