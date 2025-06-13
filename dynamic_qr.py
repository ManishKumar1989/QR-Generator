import os
import qrcode
import pandas as pd
from PIL import Image, ImageDraw
from flask import Flask, jsonify, abort

def create_vcard(
    first_name, last_name, designation, department, org, phone, email, linkedin=None, office_address=None, website=None
):
    vcard = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"N:{last_name};{first_name};;;",
        f"FN:{first_name} {last_name}",
    ]
    # Add department as part of ORG for compatibility (ORG:Company;Department)
    if org and department:
        vcard.append(f"ORG:{org};{department}")
    elif org:
        vcard.append(f"ORG:{org}")
    elif department:
        vcard.append(f"ORG:;{department}")
    if designation:
        vcard.append(f"TITLE:{designation}")
    if office_address:
        vcard.append(f"ADR;TYPE=WORK:;;{office_address};;;;;")
    if phone:
        vcard.append(f"TEL;CELL:{phone}")
    if email:
        vcard.append(f"EMAIL:{email}")
    if website:
        vcard.append(f"URL:{website}")
    if linkedin:
        vcard.append(f"URL:{linkedin}")
    vcard.append("END:VCARD")
    return "\n".join(vcard)

def generate_qr(url, filename):
    qr = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#284181", back_color="white").convert('RGB')

    logo_path = "logo.jpg"
    if os.path.exists(logo_path):
        logo = Image.open(logo_path).convert("RGBA")
        qr_width, qr_height = img.size
        max_logo_width = int(qr_width * 0.25)
        logo.thumbnail((max_logo_width, max_logo_width), Image.LANCZOS)
        logo_width, logo_height = logo.size

        # Adjust padding: more padding above, less below
        padding_top = int(logo_height * 0.05)
        padding_bottom = int(logo_height * 0.05)
        total_logo_height = logo_height + padding_top + padding_bottom

        # Create new image to fit QR code and logo below
        new_height = qr_height + total_logo_height
        combined_img = Image.new("RGB", (qr_width, new_height), "white")
        combined_img.paste(img, (0, 0))

        # Center the logo horizontally below the QR code, move it up by reducing bottom padding
        logo_x = (qr_width - logo_width) // 2
        logo_y = qr_height + padding_top
        combined_img.paste(logo, (logo_x, logo_y), mask=logo if logo.mode == 'RGBA' else None)

        combined_img.save(filename)
    else:
        print(f"Logo not found at {logo_path}, saving QR without logo.")
        img.save(filename)
    print(f"QR code saved as {filename}")

# --- Flask app to serve employee info locally ---
app = Flask(__name__)
df = pd.read_excel("contacts1.xlsx")  # Load once at startup

@app.route("/employees/<pen>")
def get_employee(pen):
    row = df[df["PEN"].astype(str) == str(pen)]
    if row.empty:
        abort(404)
    data = row.iloc[0].to_dict()
    return jsonify(data)

if __name__ == "__main__":
    # Create "Dynamic qr" folder if it doesn't exist
    os.makedirs("Dynamic qr", exist_ok=True)

    # Generate QR codes with local URL
    for idx, row in df.iterrows():
        pen = row.get("PEN", "")
        first_name = row.get("first_name", "")
        last_name = row.get("last_name", "")
        url = f"http://localhost:5000/employees/{pen}"
        filename = os.path.join("Dynamic qr", f"{pen} - {first_name} {last_name}.png")
        generate_qr(url, filename)

    # Start the Flask server
    print("Starting local server at http://localhost:5000 ...")
    app.run(debug=False)
