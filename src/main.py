# Step 1: Install dependencies
!sudo apt install tesseract-ocr -y
!pip install pytesseract opencv-python pillow pandas openpyxl

# Step 2: Import libraries
import cv2
import pytesseract
import pandas as pd

# Step 3: Define file paths
blank_form_path = "/content/form-1.png"        # path to blank form
filled_form_path = "assets/sample_input.jpeg"  # path to filled handwritten form
output_excel_path = "/content/final_form_data.xlsx"

# Step 4: Load both images
blank_img = cv2.imread(blank_form_path)
filled_img = cv2.imread(filled_form_path)

# Check image loading
if blank_img is None or filled_img is None:
    raise Exception("❌ Error: Could not load one or both images. Check file paths.")

# Step 5: Coordinates for printed text (on blank form)
column_coords = {
    "Last name": (120, 310, 380, 360),
    "First name": (430, 310, 740, 360),
    "Middle name(s)": (790, 310, 1080, 360),
    "Home address": (120, 420, 1080, 470),
    "Email": (120, 540, 1080, 590),
    "Home phone number": (120, 660, 520, 710),
    "Mobile phone number": (560, 660, 1080, 710),
    "Date": (120, 850, 300, 890),
    "Signature": (720, 850, 1080, 890)
}

# Step 6: Coordinates for handwritten text (filled form only)
handwritten_coords = {
    "Last name": (120, 360, 380, 420),
    "First name": (430, 360, 740, 420),
    "Middle name(s)": (790, 360, 1080, 420),
    "Home address": (120, 470, 1080, 530),
    "Email": (120, 590, 1080, 650),
    "Home phone number": (120, 710, 520, 770),
    "Mobile phone number": (560, 710, 1080, 770),
    "Date": (120, 890, 300, 940),
    "Signature": (720, 890, 1080, 940)
}

# Step 7: Extract column names from blank form
print("📋 Extracting column labels from blank form...\n")
columns = []
for field, (x1, y1, x2, y2) in column_coords.items():
    roi = blank_img[y1:y2, x1:x2]
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    # Only use the given field names, not OCR text
    columns.append(field)
print("✅ Column headers:", columns)

# Step 8: Extract only handwritten text (ignore printed labels)
print("\n✍ Extracting handwritten text from filled form...\n")
handwritten_data = {}
for field, (x1, y1, x2, y2) in handwritten_coords.items():
    roi = filled_img[y1:y2, x1:x2]
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

    # OCR extraction (no printed labels)
    text = pytesseract.image_to_string(gray, config="--oem 3 --psm 6 -l eng").strip()
    handwritten_data[field] = text
    print(f"{field}: {text}")

# Step 9: Save combined result to Excel
df = pd.DataFrame([handwritten_data], columns=columns)
df.to_excel(output_excel_path, index=False)

print("\n✅ Handwritten data extracted and saved successfully!")
print("📂 Excel file saved at:", output_excel_path)
