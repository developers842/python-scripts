import os
import re
import fitz  # PyMuPDF
import pandas as pd
from datetime import datetime, timedelta

# Path to folder with PDFs
CV_FOLDER = "E:\\CVs"
OUTPUT_FILE = 'cv_name_summary.xlsx'

# Regex for name extraction
NAME_PATTERN = re.compile(r'(?:Name\s*[:\-]?\s*)([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)')

# Natural sort key
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

# List of PDFs in natural order
pdf_files = sorted(
    [f for f in os.listdir(CV_FOLDER) if f.lower().endswith('.pdf')],
    key=natural_sort_key
)

# Set starting time
start_time = datetime.strptime("10:00", "%H:%M")

# Store extracted data
data = []

for i, filename in enumerate(pdf_files):
    serial = os.path.splitext(filename)[0]
    filepath = os.path.join(CV_FOLDER, filename)

    try:
        doc = fitz.open(filepath)
        text = ""
        for page in doc:
            text += page.get_text()

        # Try extracting name
        name_match = NAME_PATTERN.search(text)
        name = name_match.group(1) if name_match else None

        # Fallback to first line
        if not name:
            first_page_text = doc[0].get_text("text")
            for line in first_page_text.split("\n"):
                cleaned = line.strip()
                if len(cleaned.split()) >= 2 and cleaned[0].isupper():
                    name = cleaned
                    break

        if not name:
            name = "Not Found"

        # Calculate time slot
        time_slot = (start_time + timedelta(minutes=15 * i)).strftime("%I:%M %p")

        data.append({
            "S.No": serial,
            "Extracted Name": name,
            "Time Slot": time_slot
        })

    except Exception as e:
        print(f"❌ Error processing {filename}: {e}")

# Save to Excel
df = pd.DataFrame(data)
df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Done! Names and time slots saved to: {OUTPUT_FILE}")
