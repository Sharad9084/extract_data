import os
import re
import pdfplumber
import pandas as pd

# ==========================================
# PATHS
# ==========================================
INPUT_FOLDER = "input_pdfs"
OUTPUT_FOLDER = "output"
OUTPUT_FILE = os.path.join(OUTPUT_FOLDER, "extracted_data.xlsx")

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ==========================================
# HELPERS
# ==========================================
def clean(x):
    if x is None:
        return ""
    return str(x).replace("\n", " ").strip()


def get_value(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        return clean(match.group(1))
    return ""


# ==========================================
# HEADER EXTRACTION
# ==========================================
def extract_header(full_text):
    data = {}

    # -------------------------
    # CLIENT NAME
    # -------------------------
    client = ""
    lines = full_text.split("\n")

    for i, line in enumerate(lines):
        if "Service provided To" in line:
            if i + 1 < len(lines):
                client = clean(lines[i + 1])

                # unwanted text remove
                client = client.replace("Activity Details", "")
                client = client.replace("State Code:", "")
                client = client.strip()

            break

    data["Client Name"] = client

    # -------------------------
    # Other fields
    # -------------------------
    data["Bill Date"] = get_value(
        r"Date\s*:\s*(.*?)\n",
        full_text
    )

    data["Activity Month"] = get_value(
        r"Activity month\s*:\s*(.*?)\n",
        full_text
    )

    data["Brand"] = get_value(
        r"Brand Name\s*:\s*(.*?)\n",
        full_text
    )

    data["Estimate No"] = get_value(
        r"Estimate Number\s*:\s*(.*?)\n",
        full_text
    )

    data["PAN Number"] = get_value(
        r"PAN Number\s*:\s*([A-Z0-9]+)",
        full_text
    )

    data["Campaign Name"] = get_value(
        r"Campaign Name\s*:\s*(.*?)\s*(Plan No:|\n)",
        full_text
    )

    return data


# ==========================================
# CHANNEL NAME
# ==========================================
def find_channel_name(table):
    for row in table:
        if row and row[0]:
            val = clean(row[0])

            bad_words = [
                "channel", "program", "title",
                "spot", "page", "wed", "thu",
                "fri", "sat", "sun", "mon",
                "tue", "total"
            ]

            if val and not any(x in val.lower() for x in bad_words):
                return val

    return ""


# ==========================================
# DAY COLUMN TO SPOT DATE
# ==========================================
def build_spot_dates(row):
    result = []
    total = 0

    day_start = 11

    for day in range(1, 31):
        idx = day_start + (day - 1)

        if idx < len(row):
            val = clean(row[idx])

            if val.isdigit():
                count = int(val)

                if count > 0:
                    result.append(f"{day}({count})")
                    total += count

    return ", ".join(result), total


# ==========================================
# PROCESS PDF
# ==========================================
def process_pdf(pdf_path):
    final_rows = []

    with pdfplumber.open(pdf_path) as pdf:

        # Full text for header
        full_text = ""

        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                full_text += txt + "\n"

        header = extract_header(full_text)

        # Page loop
        for page_no, page in enumerate(pdf.pages, start=1):

            tables = page.extract_tables()

            if not tables:
                continue

            for table_no, table in enumerate(tables, start=1):

                if len(table) < 2:
                    continue

                channel_name = find_channel_name(table)

                for row in table:

                    if not row:
                        continue

                    row = [clean(x) for x in row]
                    row_text = " ".join(row).lower()

                    # -------------------------
                    # Skip header rows
                    # -------------------------
                    skip_words = [
                        "program time",
                        "net cost",
                        "rate per",
                        "spot dur",
                        "position"
                    ]

                    if any(x in row_text for x in skip_words):
                        continue

                    # Skip total row
                    if "total" in row_text:
                        continue

                    if len(row) < 10:
                        continue

                    # ----------------------------------
                    # Column Mapping
                    # ----------------------------------
                    # 0 Channel
                    # 1 Program
                    # 2 Program Time
                    # 3 Title
                    # 4 Spot Dur
                    # 5 Pmt Type
                    # 6 Position
                    # 7 Total FCT
                    # 8 Rate
                    # 9 No of Spots
                    # 10 Net Cost

                    program = row[1]
                    program_time = row[2]
                    title = row[3]
                    duration = row[4]
                    position = row[6]
                    rate = row[8]
                    total_spots_given = row[9]
                    net_cost = row[10]

                    # ----------------------------------
                    # If Program contains time
                    # Example: RODP (08.00 - 10.00)
                    # ----------------------------------
                    if "(" in program and ")" in program:
                        m = re.search(r"(.*?)\((.*?)\)", program)

                        if m:
                            prog_name = clean(m.group(1))
                            time_val = clean(m.group(2))

                            if prog_name:
                                program = prog_name

                            if program_time == "":
                                program_time = time_val

                    # Blank row skip
                    if title == "" and program == "":
                        continue

                    # ----------------------------------
                    # Spot dates
                    # ----------------------------------
                    spot_dates, total_spots_calc = build_spot_dates(row)

                    if total_spots_given.isdigit():
                        total_spots = int(total_spots_given)
                    else:
                        total_spots = total_spots_calc

                    # ----------------------------------
                    # FCT
                    # ----------------------------------
                    try:
                        fct = int(total_spots) * int(duration)
                    except:
                        fct = ""

                    final_rows.append({
                        "Client Name": header["Client Name"],
                        "Bill Date": header["Bill Date"],
                        "Activity Month": header["Activity Month"],
                        "Brand": header["Brand"],
                        "Estimate No": header["Estimate No"],
                        "PAN Number": header["PAN Number"],
                        "Campaign Name": header["Campaign Name"],
                        "Channel Name": channel_name,
                        "Program": program,
                        "Program Time": program_time,
                        "Title": title,
                        "Position": position,
                        "Duration": duration,
                        "Gross Rate (Spot)": rate,
                        "Spot Date (#Spots)": spot_dates,
                        "Total Spots": total_spots,
                        "FCT": fct,
                        "Net Cost": net_cost,
                        "Page No": page_no,
                        "Table No": table_no
                    })

    return final_rows


# ==========================================
# MAIN
# ==========================================
def main():
    all_rows = []

    files = [
        f for f in os.listdir(INPUT_FOLDER)
        if f.lower().endswith(".pdf")
    ]

    if not files:
        print("No PDF found inside input_pdfs folder.")
        return

    for file in files:
        print("Processing:", file)

        path = os.path.join(INPUT_FOLDER, file)

        rows = process_pdf(path)

        all_rows.extend(rows)

    if all_rows:
        df = pd.DataFrame(all_rows)

        # remove duplicate rows
        df.drop_duplicates(inplace=True)

        df.to_excel(OUTPUT_FILE, index=False)

        print("Done ->", OUTPUT_FILE)

    else:
        print("No Data Extracted")


if __name__ == "__main__":
    main()