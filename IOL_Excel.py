import os
import pdfplumber
import re
import pandas as pd
from tkinter import filedialog

# Define the data fields to extract
FIELDS = [
    'Name', 'ID',
    'LS_OD', 'LS_OS', 'VS_OD', 'VS_OS','LVC_OD', 'LVC_OS','Target_ref_OD', 'Target_ref_OS', 'AL_OD', 'AL_OS', 'ACD_OD', 'ACD_OS', 'LT_OD', 'LT_OS',
    'WTW_OD', 'WTW_OS', 'SE_OD', 'SE_OS','K1_OD', 'K1_OS', 'ΔK_OD', 'ΔK_OS', 'K2_OD', 'K2_OS', 'Box_1_Name', 'Box_2_Name', 'Box_3_Name', 'Box_4_Name',
    'Formula','Box_1_IOL_OD', 'Box_1_REF_OD','Box_2_IOL_OD', 'Box_2_REF_OD', 'Box_1_IOL_OS', 'Box_1_REF_OS','Box_2_IOL_OS', 'Box_2_REF_OS',
    'Box_3_IOL_OD', 'Box_3_REF_OD','Box_4_IOL_OD', 'Box_4_REF_OD','Box_3_IOL_OS', 'Box_3_REF_OS','Box_4_IOL_OS', 'Box_4_REF_OS'
]

# Extract values using regex patterns
def extract_data(text):
    data = dict.fromkeys(FIELDS)

    # Normalize the text to make parsing easier
    fullpage = text.replace('\n', ' ')
    # Find patient name and ID
    match_name = re.search(r'Patient ([A-Z]+, [A-Z]+)', fullpage)
    if match_name:
        data['Name'] = match_name.group(1).strip()

    match_id = re.search(r'Patient ID\s+(\d+)', fullpage)
    if match_id:
        data['ID'] = match_id.group(1)

    # Eye-specific values
    pairs = {
        'LS': r'LS:\s*(\w+)', 'VS': r'VS:\s*([A-Za-z ]+[a-z]+?)(?:\s|Ref|LVC|Target ref|$)', 'LVC': r'LVC:\s*(\w+)', 'Target_ref': r'Target ref\.:\s*([+\-]?\d+\.\d+\s*D)', 
        'AL': r'AL:\s*(\d+\.\d+mm)','ACD': r'ACD:\s*(\d+\.\d+mm)', 'LT': r'LT:\s*(\d+\.\d+mm)', 'WTW': r'WTW:\s*(---|\d+\.\d+mm)', 'SE': r'SE:\s*(\d+\.\d+D)', 
        'K1': r'K1:\s*(\d+\.\d+D)','ΔK': r'ΔK:\s*([+\-]?\d+\.\d+D)', 'K2': r'K2:\s*(\d+\.\d+D)'
    }

    # Extract values for each eye
    for key, pattern in pairs.items():
        match = re.findall(pattern, text)
        if match:
                data[f"{key}_OD"] = match[0]
                data[f"{key}_OS"] = match[1]
            
    # Use ' K/TK ' as a delimiter to split model name blocks
    candidates = re.split(r'\s*(K|TK)\s', fullpage)
    box_matches = []
    for block in candidates:
        match = re.match(r'(Alcon(?:/[A-Za-z]+)?|AMO|baush & lomb)\s+[A-Z0-9]+(?:\s+[A-Z0-9]+)*', block.strip(), re.IGNORECASE)
        if match:
            box_matches.append(match.group(0).strip())

    # Assign box names based on the matches found 
    data['Box_1_Name'] = box_matches[0]
    data['Box_2_Name'] = box_matches[1]
    data['Box_3_Name'] = box_matches[4]
    if len(box_matches) >= 8:
        data['Box_4_Name'] = box_matches[5]
    else:
         data['Box_4_Name'] = ""


    # Find formula matches
    formula_matches = re.findall(r'-\s+([A-Za-z®/\-\s]+?)\s+-', fullpage)
    if formula_matches:
        data['Formula'] = formula_matches[0]

    # Extract IOL and REF values from the blocks
    blocks = re.findall(r'IOL \(D\)(.*?)Emmetropia', fullpage, re.DOTALL)
    numbers = re.findall(r'[+-]?\d+\.\d+', blocks[0])
    Box1_Box2_numbers = [numbers[i:i+8] for i in range(0, len(numbers), 8)]
    Box1_Box2_numbers = Box1_Box2_numbers[2]
    
    numbers = re.findall(r'[+-]?\d+\.\d+', blocks[1])
    if len(box_matches) >= 8:
        Box3_Box4_numbers = [numbers[i:i+8] for i in range(0, len(numbers), 8)]   
    else:
        Box3_Box4_numbers = [numbers[i:i+4] for i in range(0, len(numbers), 4)]
        
    Box3_Box4_numbers = Box3_Box4_numbers[2]

    data["Box_1_IOL_OD"] = Box1_Box2_numbers[0]
    data["Box_1_REF_OD"] = Box1_Box2_numbers[1]
    data["Box_2_IOL_OD"] = Box1_Box2_numbers[2]
    data["Box_2_REF_OD"] = Box1_Box2_numbers[3]
    data["Box_1_IOL_OS"] = Box1_Box2_numbers[4]
    data["Box_1_REF_OS"] = Box1_Box2_numbers[5]
    data["Box_2_IOL_OS"] = Box1_Box2_numbers[6]
    data["Box_2_REF_OS"] = Box1_Box2_numbers[7]
    data["Box_3_IOL_OD"] = Box3_Box4_numbers[0]
    data["Box_3_REF_OD"] = Box3_Box4_numbers[1]
    if len(box_matches) >= 8:
        data["Box_4_IOL_OD"] = Box3_Box4_numbers[2]
        data["Box_4_REF_OD"] = Box3_Box4_numbers[3]
        data["Box_3_IOL_OS"] = Box3_Box4_numbers[4]
        data["Box_3_REF_OS"] = Box3_Box4_numbers[5]
        data["Box_4_IOL_OS"] = Box3_Box4_numbers[6]
        data["Box_4_REF_OS"] = Box3_Box4_numbers[7]
    else:
        data["Box_3_IOL_OS"] = Box3_Box4_numbers[2]
        data["Box_3_REF_OS"] = Box3_Box4_numbers[3]

    return data

# Main processing function
def process_pdfs(folder):
    extracted_data = []

    for filename in os.listdir(folder):
        if filename.lower().endswith(".pdf"):
            path = os.path.join(folder, filename)
            try:
                with pdfplumber.open(path) as pdf:
                    page_text = pdf.pages[0].extract_text(x_tolerance=3, x_tolerance_ratio=None, y_tolerance=3, layout=False, x_density=7.25, y_density=13, line_dir_render=None, char_dir_render=None)
                    values = extract_data(page_text)
                    extracted_data.append(values)
                    print(f"Extracted data from {filename}")
            except Exception as e:
                print(f"Failed to read {filename}: {e}")

    return extracted_data

# Save to Excel
def export_to_excel(data_list, out_path):
    headers = [
        'Name', 'ID', 'EYE', 'LS', 'VS', 'LVC', 'Target Ref', 'AL', 'ACD', 'LT', 'WTW', 'K1', 'K2', 'ΔK', 'SE',
        'Box_1_IOL', 'Box_1_REF', 'Box_2_IOL', 'Box_2_REF', 'Box_3_IOL', 'Box_3_REF', 'Box_4_IOL', 'Box_4_REF',
    ]

    rows = []
    for data in data_list:
        Box_1_Name = data.get("Box_1_Name")
        Box_2_Name = data.get("Box_2_Name")
        Box_3_Name = data.get("Box_3_Name") 
        Box_4_Name = data.get("Box_4_Name")
        formula = data.get("Formula")
        # Prepare data for each eye
        for eye in ["OD", "OS"]:
            row = {
                "Name": data.get("Name") if eye == "OD" else "",
                "ID": data.get("ID") if eye == "OD" else "",
                "EYE": eye,
                "LS": data.get(f"LS_{eye}"),
                "VS": data.get(f"VS_{eye}"),
                "LVC": data.get(f"LVC_{eye}"),
                "Target Ref": data.get(f"Target_ref_{eye}"),
                "AL": data.get(f"AL_{eye}"),
                "ACD": data.get(f"ACD_{eye}"),
                "LT": data.get(f"LT_{eye}"),
                "WTW": data.get(f"WTW_{eye}"),
                "K1": data.get(f"K1_{eye}"),
                "K2": data.get(f"K2_{eye}"),
                "ΔK": data.get(f"ΔK_{eye}"),
                "SE": data.get(f"SE_{eye}"),
                "Box_1_IOL": Box_1_Name + " " + formula + " " + data.get("Box_1_IOL_" + eye),
                "Box_1_REF": Box_1_Name + " " + formula + " " + data.get("Box_1_REF_" + eye),
                "Box_2_IOL": Box_2_Name + " " + formula + " " + data.get("Box_2_IOL_" + eye),
                "Box_2_REF": Box_2_Name + " " + formula + " " + data.get("Box_2_REF_" + eye),
                "Box_3_IOL": Box_3_Name + " " + formula + " " + data.get("Box_3_IOL_" + eye),
                "Box_3_REF": Box_3_Name + " " + formula + " " + data.get("Box_3_REF_" + eye),
                "Box_4_IOL": Box_4_Name + " " + formula + " " + data.get("Box_4_IOL_" + eye) if Box_4_Name else "",
                "Box_4_REF": Box_4_Name + " " + formula + " " + data.get("Box_4_REF_" + eye) if Box_4_Name else ""
                }
            rows.append(row)

    df = pd.DataFrame(rows, columns=headers)
    df.to_excel(out_path, index=False)

# Run script
def main():
    folder = filedialog.askdirectory(title="Select the folder containing PDF files")
    if not folder:
        print("No folder selected.")
        return

    data = process_pdfs(folder)
    if data:
        output_file = os.path.join(folder, "extracted_data.xlsx")
        export_to_excel(data, output_file)
    else:
        print("No data extracted.")

if __name__ == "__main__":
    main()
