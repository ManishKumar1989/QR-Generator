# Simple and Dynamic Excel based QR Code Generator

This project generates QR codes for contacts using data from an Excel file. It supports two modes:
- **Static vCard QR Codes**: Generates QR codes containing vCard contact information.
- **Dynamic QR Codes with Local Server**: Generates QR codes that link to a local Flask server serving employee details.

## Features

- Reads contact data from `contacts.xlsx` or `contacts1.xlsx`.
- Generates QR codes with embedded vCard data or dynamic URLs.
- Adds a logo below each QR code (if `logo.jpg` is present).
- Saves QR codes in organized folders (`QR Codes/` and `Dynamic qr/`).
- Local Flask server to serve employee details as JSON.

## Requirements

- Python 3.x
- Packages:
  - `qrcode`
  - `pandas`
  - `Pillow`
  - `flask` (for dynamic QR codes)

Install dependencies with:

```sh
pip install qrcode[pil] pandas pillow flask
```

## Usage

### 1. Static vCard QR Codes

Generates QR codes with embedded vCard data for each contact.

```sh
python generate.py
```

- Reads from `contacts.xlsx`.
- Outputs PNG files to the `QR Codes/` directory.

### 2. Dynamic QR Codes with Local Server

Generates QR codes that point to a local server endpoint for each employee.

```sh
python dynamic_qr.py
```

- Reads from `contacts1.xlsx`.
- Outputs PNG files to the `Dynamic qr/` directory.
- Starts a Flask server at [http://localhost:5000](http://localhost:5000).
- Each QR code links to `/employees/<PEN>` endpoint.

## File Structure

- `generate.py` - Script for static vCard QR code generation.
- `dynamic_qr.py` - Script for dynamic QR code generation and local server.
- `contacts.xlsx` - Excel file with contact data for static QR codes.
- `contacts1.xlsx` - Excel file with contact data for dynamic QR codes.
- `logo.jpg` - Logo image to be added below QR codes (optional).
- `QR Codes/` - Output folder for static QR codes.
- `Dynamic qr/` - Output folder for dynamic QR codes.

## Customization

- Update the Excel files with your contact data.
- Replace `logo.jpg` with your organization's logo (optional).

## License

MIT License
