"""
Certificate Bulk Mailer — Zoom Participant Export Edition
==========================================================
Requirements:
    pip install pillow reportlab

Files needed (same folder as this script):
    - certificate_template.png        ← blank cert from Canva
    - participants_85326044081_...     ← the Zoom export CSV (rename or update ZOOM_FILE below)

Gmail App Password setup:
    1. myaccount.google.com > Security > 2-Step Verification → enable
    2. myaccount.google.com/apppasswords → create one → paste below
"""

import os, re, csv, io, smtplib
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO

# ─── CONFIG ──────────────────────────────────────────────────────────────────

GMAIL_ADDRESS      = os.getenv("GMAIL_ADDRESS")         # ← change this
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")           # ← paste app password here

TEMPLATE_FILE      = "certificate_template.png"     # ← Canva export
ZOOM_FILE          = "recipients.xlsx"              # ← the Zoom CSV (keep whatever filename)

# Name text styling
FONT_SIZE          = 80
FONT_COLOR         = (18, 53, 120)                  # dark blue

# Position — fraction of image (0.0–1.0). Y=0.52 = just above gold line
NAME_X_FRACTION    = 0.50
NAME_Y_FRACTION    = 0.44

EMAIL_SUBJECT      = "Your Certificate of Participation – UNITE 2026"
EMAIL_BODY         = """\
Dear {name},

Congratulations! Please find attached your Certificate of Participation for the \
South Asia Presidents and Secretaries Virtual Meet-up – "UNITE 2026".

It was a pleasure having you with us.

Warm regards,
Rotaract South Asia MDIO (RSA MDIO)
"""

OUTPUT_FOLDER      = "generated_certificates"

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def extract_clean_name(raw: str) -> str:
    """
    Handles Zoom display names like:
      '3233- Varsha Sriram - Sec (Varsha Sriram)' → 'Varsha Sriram'
      'PDRR Darryl - President RSA (Darryl Dsouza)' → 'Darryl Dsouza'
      'Rithika T' → 'Rithika T'
    If there's text in (brackets), use that. Otherwise use raw name.
    """
    match = re.search(r'\(([^)]+)\)$', raw.strip())
    if match:
        return match.group(1).strip()
    return raw.strip()


def load_recipients(filepath: str):
    """Parse Zoom CSV export, deduplicate by email, return list of (name, email)."""
    with open(filepath, encoding='utf-8-sig') as f:
        lines = f.read().split('\n')

    # Find the participant data header row
    header_idx = next(
        (i for i, l in enumerate(lines) if l.startswith('Name (original name)')),
        None
    )
    if header_idx is None:
        raise ValueError("Could not find participant data in file. Expected 'Name (original name)' column.")

    data = '\n'.join(lines[header_idx:])
    reader = csv.DictReader(io.StringIO(data))

    seen = set()
    recipients = []
    for row in reader:
        email = row.get('Email', '').strip()
        raw_name = row.get('Name (original name)', '').strip()
        if not email or email in seen:
            continue
        seen.add(email)
        clean = extract_clean_name(raw_name)
        recipients.append((clean, email))

    return recipients


def load_font(size):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
        "C:/Windows/Fonts/Georgia.ttf",
        "C:/Windows/Fonts/times.ttf",
        "/Library/Fonts/Georgia.ttf",
        "/System/Library/Fonts/Supplemental/Georgia.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def stamp_name_on_cert(name: str, output_pdf: str):
    img = Image.open(TEMPLATE_FILE).convert("RGBA")
    draw = ImageDraw.Draw(img)
    font = load_font(FONT_SIZE)

    img_w, img_h = img.size
    bbox = draw.textbbox((0, 0), name, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    x = (img_w * NAME_X_FRACTION) - (text_w / 2)
    y = (img_h * NAME_Y_FRACTION) - (text_h / 2)

    draw.text((x, y), name, font=font, fill=FONT_COLOR)

    img_rgb = img.convert("RGB")
    img_bytes = BytesIO()
    img_rgb.save(img_bytes, format="PNG")
    img_bytes.seek(0)

    c = rl_canvas.Canvas(output_pdf, pagesize=(img_w, img_h))
    c.drawImage(ImageReader(img_bytes), 0, 0, width=img_w, height=img_h)
    c.save()


def send_email(to_email: str, name: str, pdf_path: str):
    msg = MIMEMultipart()
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = to_email
    msg["Subject"] = EMAIL_SUBJECT
    msg.attach(MIMEText(EMAIL_BODY.format(name=name), "plain"))

    with open(pdf_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="Certificate_{name}.pdf"')
    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_email, msg.as_string())


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    recipients = load_recipients(ZOOM_FILE)
    total = len(recipients)
    print(f"Found {total} unique recipients.\n")

    failed = []

    for i, (name, email) in enumerate(recipients):
        print(f"[{i+1}/{total}] {name} <{email}>")

        safe_name = "".join(c for c in name if c.isalnum() or c in " _-").strip()
        pdf_path  = os.path.join(OUTPUT_FOLDER, f"{safe_name}.pdf")

        try:
            stamp_name_on_cert(name, pdf_path)
            send_email(email, name, pdf_path)
            print(f"  ✓ Sent")
        except Exception as e:
            print(f"  ✗ Failed: {e}")
            failed.append((name, email, str(e)))

    print(f"\n─── Done ───")
    print(f"Sent: {total - len(failed)}/{total}")
    if failed:
        print(f"\nFailed ({len(failed)}):")
        for name, email, err in failed:
            print(f"  {name} <{email}> → {err}")


if __name__ == "__main__":
    main()
