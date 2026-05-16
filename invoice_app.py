import os
import re
import pdfplumber
import pandas as pd
from glob import glob

# ==========================================
# PATHS
# ==========================================
INPUT_FOLDER = "invoice_pdfs"
OUTPUT_FOLDER = "output"
OUTPUT_FILE = os.path.join(OUTPUT_FOLDER, "invoice_extracted.xlsx")

os.makedirs(OUTPUT_FOLDER, exist_ok=True)



# ==========================================
# HELPERS
# ==========================================
def clean(x):
    """Clean text value"""
    if x is None:
        return ""
    return str(x).replace("\n", " ").strip()


def fix_zee_text(text):
    """Collapse spaced-out text caused by PDF font rendering.
    Handles: 'H I N D I' -> 'HINDI', 'H I NDI' -> 'HINDI', 
    'E N TERTAINMENT' -> 'ENTERTAINMENT'"""
    if not text:
        return text
    # Match sequences of single uppercase letters (possibly mixed with short fragments)
    # separated by spaces, where at least 2 single letters appear
    def collapse_match(m):
        return m.group(0).replace(' ', '')
    # First: collapse pure single-letter sequences: "H I N D I"
    text = re.sub(r'\b([A-Z] ){2,}[A-Z]\b', collapse_match, text)
    # Then: collapse single letter + fragment: "H I NDI", "E N TERTAINMENT"  
    # Pattern: single letter, space, single letter, space, fragment (or vice versa)
    for _ in range(5):
        new = re.sub(r'\b([A-Z]) ([A-Z]{2,})\b', r'\1\2', text)
        new = re.sub(r'\b([A-Z]{2,}) ([A-Z])\b', r'\1\2', new)
        new = re.sub(r'\b([A-Z]) ([A-Z])\b', r'\1\2', new)
        if new == text:
            break
        text = new
    return text


# Star India station code -> channel name mapping
STAR_STN_MAP = {
    'MAIN':    'Star Maa',
    'MAAIN':   'Star Maa HD',
    'MMAAHD':  'Star Maa HD',
    'MMIN':    'Star Maa Movies',
    'MMIN1':   'Star Maa Movies',
    'MMIN2':   'Star Maa Movies',
    'MMI':     'Star Maa Movies',
    'MGI':     'Star Maa Gold',
    'MGOLD':   'Star Maa Gold',
    'VIJAYIN': 'Star Vijay',
    'VIJAY':   'Star Vijay',
    'SBIN':    'Star Bharat',
    'SBHARAT': 'Star Bharat',
    'ASIANET': 'Asianet',
    'ASNET':   'Asianet',
    'SUVARNASTAR': 'Star Suvarna',
    'SUVARNA': 'Star Suvarna',
}

STAR_RELATION_KEYWORDS = [
    ('maa gold',    'Star Maa Gold'),
    ('maa hd',      'Star Maa HD'),
    ('maa movies',  'Star Maa Movies'),
    ('maa',         'Star Maa'),
    ('vijay',       'Star Vijay'),
    ('bharat',      'Star Bharat'),
    ('suvarna',     'Star Suvarna'),
    ('asianet',     'Asianet'),
]


def map_star_channel(stn_code, station_relation):
    """Return a human-readable channel name from STN code or Station Relation."""
    # Try STN code first
    ch = STAR_STN_MAP.get(stn_code.upper().strip())
    if ch:
        return ch
    # Fall back to Station Relation keyword scan
    sr_lower = station_relation.lower()
    for kw, name in STAR_RELATION_KEYWORDS:
        if kw in sr_lower:
            return name
    return station_relation  # last resort: return as-is


def _derive_day_from_date(date_str):
    """
    Derive day name (Mon, Tue, Wed, ...) from a date string.
    Supports: dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy, dd-Mon-yyyy, 
              dd-Mon-yy, dd/mm/yy, mm/dd/yyyy
    """
    from datetime import datetime
    if not date_str or not date_str.strip():
        return ""
    
    date_str = date_str.strip()
    
    # Try multiple date formats
    formats = [
        "%d/%m/%Y",      # 01/11/2023
        "%d-%m-%Y",      # 01-11-2023
        "%d.%m.%Y",      # 01.11.2023
        "%d-%b-%Y",      # 01-Nov-2023
        "%d-%b-%y",      # 01-Nov-23
        "%d/%m/%y",      # 01/11/23
        "%d.%m.%y",      # 01.11.23
        "%m/%d/%Y",      # 11/01/2023 (US format - Vasanth)
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%a")  # Mon, Tue, Wed, ...
        except ValueError:
            continue
    
    return ""


# ==========================================
# STAR INDIA PARSER
# ==========================================
def extract_star_header(full_text):
    import re
    header = {}
    lines = full_text.split("\n")

    broadcaster_match = re.search(
        r'^(.*?Pvt\.?\s*Ltd\.?)',
        full_text,
        re.IGNORECASE | re.MULTILINE
    )
    header["Broadcaster Name"] = clean(broadcaster_match.group(1)) if broadcaster_match else ""

    header["Advertiser Name"] = ""
    header["Agency Name"] = ""
    header["Invoice Number"] = ""
    header["Channel Name"] = ""
    header["Station Relation"] = ""
    header["Brand"] = ""

    for i, line in enumerate(lines):
        if "Advertiser" in line and "Agency" in line and "Invoice" in line:
            if i + 1 >= len(lines):
                break

            data_line = lines[i + 1]
            cont_line = lines[i + 2] if i + 2 < len(lines) else ""

            inv_no = ""
            inv_start_pos = -1

            for match in re.finditer(r'\b([A-Z0-9]{8,})\b', data_line):
                candidate = match.group(1)
                if re.search(r'\d{4,}', candidate):
                    inv_no = candidate
                    inv_start_pos = match.start()
                    break

            header["Invoice Number"] = inv_no

            if inv_start_pos > 0:
                left_part = data_line[:inv_start_pos].strip()
                right_part = data_line[inv_start_pos + len(inv_no):].strip()

                adv_text = ""
                agency_text = ""

                if cont_line:
                    cont_stripped = cont_line.strip()
                    first_word = cont_stripped.split()[0] if cont_stripped.split() else ""
                    adv_suffix = ""

                    if first_word.upper() in ("LIMITED", "LTD", "LTD.", "PRIVATE", "PVT", "INC", "CORP"):
                        adv_suffix = first_word

                    mid_split = re.search(
                        r'(.*?(?:PRODUCTS|INDUSTRIES|FOODS|BEVERAGES|ENTERPRISES|CHEMICALS|HOLDINGS|PVT\.?\s*LTD\.?|PRIVATE\s+LIMITED))\s+(.*)',
                        left_part, re.IGNORECASE
                    )
                    if mid_split and len(mid_split.group(2).strip()) > 5:
                        adv_text = mid_split.group(1).strip()
                        agency_text = mid_split.group(2).strip()
                    else:
                        parts = re.split(r'\s{2,}', left_part, maxsplit=1)
                        if len(parts) == 2 and len(parts[1].strip()) > 5:
                            adv_text = parts[0].strip()
                            agency_text = parts[1].strip()
                        else:
                            adv_text = left_part

                    if adv_suffix:
                        adv_text += " " + adv_suffix
                else:
                    adv_text = left_part

                header["Advertiser Name"] = clean(adv_text)
                header["Agency Name"] = clean(agency_text)

                header["Station Relation"] = clean(right_part)

            break

    header["Channel Name"] = map_star_channel("", header.get("Station Relation", ""))

    header["Invoice Date"] = ""
    for i, line in enumerate(lines):
        if "Invoice Date" in line and "Invoice No" not in line:
            d = re.search(r'(\d{2}/\d{2}/\d{4})', line)
            if d:
                header["Invoice Date"] = d.group(1)
            elif i + 1 < len(lines):
                d = re.search(r'(\d{2}/\d{2}/\d{4})', lines[i + 1])
                if d:
                    header["Invoice Date"] = d.group(1)
            break

    header["Billing Period"] = ""
    for i, line in enumerate(lines):
        if "Billing Period" in line:
            bp = re.search(r'(\d{2}/\d{2}/\d{4})\s*[-\u2013]\s*(\d{2}/\d{2}/\d{4})', line)
            if bp:
                header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}"
            elif i + 1 < len(lines):
                combined = line + " " + lines[i + 1]
                bp = re.search(r'(\d{2}/\d{2}/\d{4})\s*[-\u2013]\s*(\d{2}/\d{2}/\d{4})', combined)
                if bp:
                    header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}"
            break

    header["PO Number"] = ""
    for i, line in enumerate(lines):
        if "PO Number" in line:
            # PO value might be on same line after "PO Number"
            po = re.search(r'PO\s+Number\s+(\S+)', line)
            if po:
                po_val = clean(po.group(1))
                # Check if next line has continuation (trailing fragment like "00")
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    parts = next_line.split()
                    if parts and len(parts[-1]) <= 3 and parts[-1].isalnum():
                        po_val = po_val + parts[-1]
                header["PO Number"] = po_val
            else:
                # PO Number is just a label, value is at end of NEXT line
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    # Take the last whitespace-separated token(s) that look like a PO
                    # e.g. "BANGALORE, KARNATAKA, 560001 01/11/2023- 15/11/2023 NOV2023/TVBRO/04754/"
                    po2 = re.search(r'(\w+\d{4}/\w+/\d+/?(?:\d+)?)\s*$', next_line)
                    if po2:
                        po_val = clean(po2.group(1))
                        # Check if next line has trailing fragment (e.g. "INDIA INDIA 00")
                        if i + 2 < len(lines):
                            next2 = lines[i + 2].strip()
                            parts = next2.split()
                            if parts and len(parts[-1]) <= 3 and parts[-1].isalnum():
                                po_val = po_val + parts[-1]
                        header["PO Number"] = po_val
            break

    return header

def extract_star_rows(full_text):
    rows = []
    lines = full_text.split("\n")

    broadcast_start = -1
    # Try multiple anchors for the data table
    anchor_patterns = [
        r'ORDER_I',             # ORDER_ID column header
        r'Plan\s*No',           # Plan No column header
        r'Broadcast\s*Date',    # Broadcast Date
        r'Telecast\s*Date',     # Telecast Date
        r'Air\s*Date',          # Air Date
    ]
    for i, line in enumerate(lines):
        for pat in anchor_patterns:
            if re.search(pat, line, re.IGNORECASE):
                broadcast_start = i + 1
                # Skip one more line if it looks like a sub-header
                if broadcast_start < len(lines):
                    nxt = lines[broadcast_start].strip()
                    if re.search(r'(Date|Day|Time|Duration|Rate|Period)', nxt, re.IGNORECASE) and not re.match(r'^\d{7}', nxt):
                        broadcast_start += 1
                break
        if broadcast_start != -1:
            break

    # Last resort: find first line that starts with a 7-digit order ID
    if broadcast_start == -1:
        for i, line in enumerate(lines):
            if re.match(r'^\d{7}\s', line.strip()):
                broadcast_start = i
                break

    if broadcast_start == -1:
        return rows

    broadcast_lines = []
    for i in range(broadcast_start, len(lines)):
        line = lines[i]
        if "Total Spots" in line or "Total Taxable" in line or "This file is signed" in line:
            break
        broadcast_lines.append(line)

    entries = []
    current_main = None
    current_cont = []

    for line in broadcast_lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("Plan"):
            continue

        if re.match(r'^\d{7}', stripped):
            if current_main is not None:
                entries.append((current_main, current_cont))
            current_main = stripped
            current_cont = []
        else:
            if current_main is not None:
                current_cont.append(stripped)

    if current_main is not None:
        entries.append((current_main, current_cont))

    for main_line, cont_lines in entries:
        row = parse_star_entry(main_line, cont_lines)
        if row:
            rows.append(row)

    return rows


def parse_star_entry(main_line, cont_lines):
    import re
    result = {}
    cont_text = " ".join(cont_lines)

    date_match = re.search(r'(\d{2}/\d{2}/\d{4})', main_line)
    if not date_match:
        return None

    result["Date"] = date_match.group(1)
    before_date = main_line[:date_match.start()].strip()
    after_date = main_line[date_match.end():].strip()

    day_match = re.match(r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b', after_date, re.IGNORECASE)
    if day_match:
        result["Day"] = day_match.group(1)
        after_day = after_date[day_match.end():].strip()
    else:
        result["Day"] = ""
        after_day = after_date

    airtime_match = re.search(r'(\d{1,2}:\d{2}:\d{2}:\d{2})', after_day)
    if airtime_match:
        result["Air Time"] = airtime_match.group(1)
        after_airtime = after_day[airtime_match.end():].strip()
    else:
        airtime_match2 = re.search(r'(\d{1,2}:\d{2}:\d{2})', after_day)
        if airtime_match2:
            result["Air Time"] = airtime_match2.group(1)
            after_airtime = after_day[airtime_match2.end():].strip()
        else:
            result["Air Time"] = ""
            after_airtime = after_day

    all_amounts = re.findall(r'(\d{1,3}(?:,\d{3})*\.\d{2})', main_line)
    result["Rate (INR)"] = all_amounts[-1] if all_amounts else ""

    len_match = re.search(r'^(\d{1,3})\s', after_airtime)
    if len_match:
        result["LEN"] = len_match.group(1)
        after_len = after_airtime[len_match.end():].strip()
    else:
        result["LEN"] = ""
        after_len = after_airtime

    rate_str = result["Rate (INR)"]
    if rate_str and after_len:
        rate_escaped = re.escape(rate_str)
        brand_end_match = re.search(
            r'([A-Z][A-Z]+(?:\s+[A-Z][A-Z]+)*)\s+' + rate_escaped + r'\s*$',
            after_len
        )
        result["Brand"] = clean(brand_end_match.group(1)) if brand_end_match else ""
    else:
        result["Brand"] = ""

    spot_main = ""
    if result["Brand"] and after_len:
        brand_pos = after_len.rfind(result["Brand"])
        if brand_pos > 0:
            spot_main = clean(after_len[:brand_pos])
    elif after_len and rate_str:
        rate_pos = after_len.rfind(rate_str)
        if rate_pos > 0:
            spot_main = clean(after_len[:rate_pos])

    spot_cont_parts = []
    for cl in cont_lines:
        cl = cl.strip()
        if not cl or re.match(r'^[\(\-]', cl):
            continue
        no_bracket = re.sub(r'\([^)]*\)-?', '', cl).strip()
        if no_bracket:
            spot_cont_parts.append(no_bracket)
    spot_cont = " ".join(spot_cont_parts)

    if spot_main and spot_cont:
        result["Spot Copy"] = f"{spot_main} {spot_cont}".strip()
    elif spot_main:
        result["Spot Copy"] = spot_main
    elif spot_cont:
        result["Spot Copy"] = spot_cont
    else:
        result["Spot Copy"] = ""

    tp_match = re.search(r'\d{7}\s+\d+\s+(Spot\s+Buys|GOL\w*)', before_date, re.IGNORECASE)
    result["TP"] = clean(tp_match.group(1)) if tp_match else ""

    combined_all = before_date + " " + cont_text
    su_matches = re.findall(r'\([^)]+\)-?', combined_all)
    result["Time Range/Sales Unit"] = "".join(su_matches)

    prog_name = ""
    # Method 1: Find program name from combined text in brackets
    first_bracket = re.search(r'\(([A-Z][A-Z\s]+?)\)[)-]', combined_all)
    if first_bracket:
        candidate = clean(first_bracket.group(1))
        if re.search(re.escape(candidate[:6]), before_date, re.IGNORECASE):
            prog_name = candidate
    
    # Method 2: Find program text between last ")- " and date in before_date
    # E.g. "(RODP)-(MON-FRI)- EARLY EVENING 14/11/2023" -> "EARLY EVENING"
    if not prog_name:
        # Look for text after last ")- " pattern
        parts = re.split(r'\)-\s*', before_date)
        if len(parts) > 1:
            last_part = parts[-1].strip()
            # Remove the date if present
            last_part = re.sub(r'\d{2}/\d{2}/\d{4}.*', '', last_part).strip()
            # Remove any remaining bracket patterns
            last_part = re.sub(r'\([^)]*\)', '', last_part).strip()
            if last_part and len(last_part) > 2:
                prog_name = last_part
    
    # Method 3: Find program name as plain text right before the date
    if not prog_name:
        prog_before_date = re.search(r'(?:MAIN\s+)?(?:\([^)]*\)\s*)?([A-Za-z][A-Za-z\s\-\']+?)\s+\d{2}/\d{2}/\d{4}', before_date)
        if prog_before_date:
            candidate = clean(prog_before_date.group(1)).strip()
            candidate = re.sub(r'\([^)]*\)?', '', candidate).strip()
            if len(candidate) > 2:
                prog_name = candidate
    
    # Method 4: Try continuation lines - program appears after ")- " 
    if not prog_name:
        for cl in cont_lines:
            prog_m = re.search(r'\)\s*[-]?\s*([A-Z][A-Za-z\s]+?)(?:\s+\d{2}/|\s+CAN\s|\s*$)', cl)
            if prog_m:
                prog_name = clean(prog_m.group(1))
                break

    prog_name = re.sub(r'\[.*?\]', '', prog_name).strip()
    prog_name = re.sub(r'\s{2,}', ' ', prog_name)
    result["Program"] = prog_name.title() if prog_name else ""

    return result

# ==========================================
# ZEE ENTERTAINMENT PARSER
# ==========================================
def extract_zee_header(full_text):
    header = {}
    
    # Static Broadcaster
    header["Broadcaster Name"] = "Zee Entertainment Enterprises Limited"

    # Extract single fields
    inv_to = re.search(r'Invoice To\s*:(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(inv_to.group(1)) if inv_to else ""

    adv = re.search(r'Advertiser\s*:(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""

    chan = re.search(r'Channel\s*:(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""

    gst = re.search(r'GST Inv\.? No\s*:(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(gst.group(1)) if gst else ""

    date_match = re.search(r'Date\s*:(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(date_match.group(1)) if date_match else ""

    brand = re.search(r'Brand\s*:(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""

    # PO / Cust Ref (Might span two lines)
    lines = full_text.split('\n')
    po_full = ""
    for i, line in enumerate(lines):
        if "Cust. Ref." in line or "Cust Ref" in line:
            po_full = line.split(":", 1)[1].strip() if ":" in line else line
            # check next line for spillover (no colon, doesn't start with next label)
            if i + 1 < len(lines):
                next_line = lines[i+1].strip()
                if ":" not in next_line and not next_line.startswith("Reverse Charge") and not next_line.startswith("IRN"):
                    po_full += " " + next_line
            break
    header["PO Number"] = clean(po_full)

    # Empty fields to match output structure
    header["Station Relation"] = ""
    header["Billing Period"] = ""

    return header


def extract_zee_rows(full_text):
    import re
    rows = []
    lines = full_text.split('\n')

    start_idx = 0  # scan everything
    for i, line in enumerate(lines):
        if re.search(r'Item\s+Caption\s+Program|No\s+Time|S\.?No.*Date.*Time|Sl\.?No.*Program', line, re.IGNORECASE):
            start_idx = i
            break

    # Strict pattern: SNo ... dd.mm.yyyy DAY HH:MM:SS LEN Rate
    PAT_STRICT = re.compile(
        r'^(\d+)\s+(.*?)\s+(\d{2}\.\d{2}\.\d{4})\s+([A-Z]{3})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,\.]+)$'
    )
    # Loose pattern: allows missing day abbreviation
    PAT_LOOSE = re.compile(
        r'^(\d+)\s+(.*?)\s+(\d{2}[./]\d{2}[./]\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,\.]+)$'
    )
    # Compact Zee format: ItemNo Caption Program Date Day Time Duration Amount
    # Example: '1 C/TETLEY-EVERYBO K A N MAHARSHI VANI 12.06.2023 MON 08:27:09 25 4,500'
    PAT_COMPACT = re.compile(
        r'^(\d+)\s+(C/\S+)\s+(.*?)\s+(\d{2}\.\d{2}\.\d{4})\s+([A-Z]{3})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,]+)$'
    )
    # Compact without day: '1 C/TETLEY-EVERYBO PROGRAMME 12.06.2023 08:27:09 25 4,500'
    PAT_COMPACT2 = re.compile(
        r'^(\d+)\s+(C/\S+)\s+(.*?)\s+(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,]+)$'
    )

    for i in range(start_idx, len(lines)):
        line = lines[i].strip()
        if not line or re.search(r'Total Value|RUPEE|Taxable|Add :|IRN|Beneficiary|Bank Account|RTGS', line, re.IGNORECASE):
            continue

        # Try compact format first (most specific)
        match_c = PAT_COMPACT.search(line)
        if match_c:
            rows.append({
                "Date": match_c.group(4), "Day": match_c.group(5),
                "Air Time": match_c.group(6), "LEN": match_c.group(7),
                "Rate (INR)": match_c.group(8),
                "Spot Copy": fix_zee_text(match_c.group(2)).title(),
                "Program": fix_zee_text(match_c.group(3)).title(),
                "TP": "", "Time Range/Sales Unit": ""
            })
            continue

        match_c2 = PAT_COMPACT2.search(line)
        if match_c2:
            rows.append({
                "Date": match_c2.group(4), "Day": "",
                "Air Time": match_c2.group(5), "LEN": match_c2.group(6),
                "Rate (INR)": match_c2.group(7),
                "Spot Copy": fix_zee_text(match_c2.group(2)).title(),
                "Program": fix_zee_text(match_c2.group(3)).title(),
                "TP": "", "Time Range/Sales Unit": ""
            })
            continue

        match = PAT_STRICT.search(line)
        if match:
            middle_text = clean(match.group(2))
            parts = re.split(r'\s{2,}', middle_text)
            caption = parts[0] if parts else middle_text
            program = " ".join(parts[1:]) if len(parts) > 1 else ""
            if not program:
                words = middle_text.split()
                caption = " ".join(words[:max(1, len(words)//2)])
                program = " ".join(words[len(words)//2:])
            rows.append({
                "Date": match.group(3), "Day": match.group(4),
                "Air Time": match.group(5), "LEN": match.group(6),
                "Rate (INR)": match.group(7),
                "Spot Copy": fix_zee_text(caption).title(),
                "Program": fix_zee_text(program).title(),
                "TP": "", "Time Range/Sales Unit": ""
            })
            continue

        match2 = PAT_LOOSE.search(line)
        if match2:
            middle_text = clean(match2.group(2))
            parts = re.split(r'\s{2,}', middle_text)
            caption = parts[0] if parts else middle_text
            program = " ".join(parts[1:]) if len(parts) > 1 else ""
            rows.append({
                "Date": match2.group(3), "Day": "",
                "Air Time": match2.group(4), "LEN": match2.group(5),
                "Rate (INR)": match2.group(6),
                "Spot Copy": fix_zee_text(caption).title(),
                "Program": fix_zee_text(program).title(),
                "TP": "", "Time Range/Sales Unit": ""
            })

    return rows





# ==========================================
# SUN TV NETWORK PARSER
# ==========================================
def extract_sun_header(full_text):
    header = {"Broadcaster Name": "SUN TV NETWORK LIMITED"}
    # Invoice No
    inv = re.search(r'Invoice No\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    # Invoice Date
    inv_date = re.search(r'Invoice Date\s*:\s*(\d{2}[./\-]\d{2}[./\-]\d{4})', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    # PO/RO - Sun uses "RO Name :"
    ro = re.search(r'RO\s*(?:Name|No|Number|Id)?\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    # Agency - Sun uses "Advertising Agency:"
    agency = re.search(r'Advertising Agency\s*:\s*(.*?)(?:Invoice|\n)', full_text, re.IGNORECASE)
    if not agency:
        agency = re.search(r'Agency\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    # Advertiser
    adv = re.search(r'Advertiser\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    # Channel - Sun has no explicit Channel field, detect from Program names
    chan_map = [
        (r'Gemini Comedy', 'Gemini Comedy'), (r'Gemini Movies', 'Gemini Movies'),
        (r'Gemini Music', 'Gemini Music'), (r'Gemini Life', 'Gemini Life'),
        (r'Gemini TV', 'Gemini TV'), (r'Gemini', 'Gemini TV'),
        (r'Sun Life', 'Sun Life'), (r'Sun Music', 'Sun Music'),
        (r'Sun News', 'Sun News'), (r'Sun TV', 'Sun TV'),
        (r'KTV', 'KTV'), (r'Surya TV', 'Surya TV'), (r'Udaya TV', 'Udaya TV'),
    ]
    header["Channel Name"] = ""
    for pattern, name in chan_map:
        if re.search(pattern, full_text, re.IGNORECASE):
            header["Channel Name"] = name
            break
    # Brand
    brand = re.search(r'Brand\s*:\s*(.*?)(?:HSN|\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)).strip() if brand else ""
    header["Billing Period"] = ""
    header["Station Relation"] = ""
    return header

def extract_sun_rows(full_text):
    rows = []
    lines = full_text.split('\n')
    start_idx = -1
    for i, line in enumerate(lines):
        if "S.No" in line and "Program" in line:
            # Skip the sub-header line (Date Time etc.)
            start_idx = i + 2
            break
    if start_idx == -1: return rows
    
    current_prog = ""
    # Known caption prefixes (derived from brand/ad names)
    caption_keywords = [
        "CG LEAF", "TATA TEA", "TATA SALT", "CHK CARE", "C/TATA", "C/",
        "CGC ", "CHAKRA", "LEAF AP", "GOLD CARE", "TEA GOLD"
    ]
    
    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped:
            continue
        if re.search(r'(IRN|FOR SUN TV|Total|Grand Total|Taxable|computer generated)', stripped, re.IGNORECASE):
            continue
        if re.search(r'^\d+ of \d+$', stripped):
            continue
        if 'S.No' in stripped and 'Program' in stripped:
            continue
        if 'Date' in stripped and 'Time' in stripped and 'Sec' in stripped:
            continue
        if re.search(r'^(SUN TV|Maran Towers|Tel:|State |CIN |Tax Invoice|To$)', stripped, re.IGNORECASE):
            continue
        if re.search(r'^(GROUP M|29,|MG ROAD|560|INDIA$|Advertising Agency|Invoice Date|Agency Pan|RO |Agency GST|Advertiser|kirioskar|Brand)', stripped, re.IGNORECASE):
            continue
        
        def split_prog_caption(text):
            """Split 'Program Caption' using known caption keywords."""
            for kw in caption_keywords:
                idx = text.find(kw)
                if idx > 0:
                    return text[:idx].strip(), text[idx:].strip()
            # Fallback: try double-space split
            parts = re.split(r'\s{2,}', text)
            if len(parts) >= 2:
                return parts[0].strip(), " ".join(parts[1:]).strip()
            return text, ""
        
        # Main data line: SNo Program+Caption Date Time Duration Amount
        m = re.search(r'^(\d+)\s+(.*?)\s+(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,\.]+)$', stripped)
        if m:
            left = m.group(2)
            prog, cap = split_prog_caption(left)
            current_prog = prog
            rows.append({
                "Date": m.group(3), "Air Time": m.group(4),
                "LEN": m.group(5), "Rate (INR)": m.group(6),
                "Program": current_prog, "Spot Copy": cap
            })
            continue
        
        # Continuation data line: Caption Date Time Duration Amount (same program)
        m2 = re.search(r'^(.*?)\s+(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,\.]+)$', stripped)
        if m2:
            cap = m2.group(1).strip()
            rows.append({
                "Date": m2.group(2), "Air Time": m2.group(3),
                "LEN": m2.group(4), "Rate (INR)": m2.group(5),
                "Program": current_prog, "Spot Copy": cap
            })
            continue
        
        # Text continuation line (like "TV SUCCESS SEP'23 - 45 SECS")
        # Check if it starts with "TV" which is part of program name like "Gemini TV"
        if rows and stripped.startswith("TV ") and current_prog and not current_prog.endswith("TV"):
            current_prog = current_prog + " TV"
            # Don't append the rest as it's caption repetition
            
    return rows

# ==========================================
# SONY NETWORK PARSER
# ==========================================
def extract_sony_header(full_text):
    import re
    header = {"Broadcaster Name": "Culver Max Entertainment Pvt. Ltd."}
    
    lines = full_text.split('\n')
    client_idx = -1
    for i, line in enumerate(lines):
        if line.startswith("Client :"):
            client_idx = i
            break
            
    if client_idx != -1 and client_idx + 2 < len(lines):
        adv_raw = clean(lines[client_idx + 1])
        if " Channel Ref" in adv_raw:
            adv_raw = adv_raw.split(" Channel Ref")[0]
        header["Advertiser Name"] = adv_raw
        
        agency_raw = clean(lines[client_idx + 2])
        if " GST Inv" in agency_raw:
            agency_raw = agency_raw.split(" GST Inv")[0]
        if " KIRLOSKAR" in agency_raw:
            agency_raw = agency_raw.split(" KIRLOSKAR")[0]
        if " P a g e" in agency_raw:
            agency_raw = agency_raw.split(" P a g e")[0]
        header["Agency Name"] = clean(agency_raw).rstrip(",")
    else:
        header["Advertiser Name"] = ""
        header["Agency Name"] = ""

    inv = re.search(r'GST Inv\. No\.(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Date:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    ro = re.search(r'R\.O\.No\s*:\s*(\S+)', full_text, re.IGNORECASE)
    if not ro:
        for line in lines:
            if "R.O.No" in line:
                ro = re.search(r'R\.O\.No\s*:\s*(.*)', line, re.IGNORECASE)
                break
    
    if ro:
        ro_val = clean(ro.group(1))
        if ":" in ro_val:
            ro_val = ro_val.split(":")[-1]
        header["PO Number"] = ro_val
    else:
        header["PO Number"] = ""

    chan = re.search(r'Channel Ref\.\s*(.*?)\s+\d+', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    # Brand from Product field
    brand = re.search(r'Product\s*:\s*(.*?)(?:R\.O\.|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)).strip() if brand else ""
    header["Billing Period"] = ""
    header["Station Relation"] = ""
    return header

def extract_sony_rows(full_text):
    rows = []
    lines = full_text.split('\n')
    start_idx = -1

    # Try multiple header anchors
    for i, line in enumerate(lines):
        if "Program description" in line and "S.No" in line:
            start_idx = i + 2  # Skip sub-header line
            break
        if re.search(r'S\.?\s*No.*Program.*Date', line, re.IGNORECASE):
            start_idx = i + 1
            if i + 1 < len(lines) and re.search(r'(in secs|Duration)', lines[i+1], re.IGNORECASE):
                start_idx = i + 2
            break
        if re.search(r'Program.*Date.*Time.*Rate', line, re.IGNORECASE):
            start_idx = i + 1
            break

    if start_idx == -1:
        # Last resort: scan all lines
        start_idx = 0
    
    current_prog = ""
    current_caption = ""
    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped:
            continue
        if re.search(r'^(Amount:|HSN/SAC|Sale of|Payable|Add:|Net Amount|CGST|SGST|Amount Due)', stripped, re.IGNORECASE):
            break
        
        # Data line: Program SNo Date Day Time Rate Duration Amount
        # Example: CID 000010 23/10 Mon 21:25:16 2,19,800.00 25 5,49,500.00
        match = re.search(r'^(.*?)\s+(\d{6})\s+(\d{2}/\d{2})\s+([A-Za-z]{3})\s+(\d{2}:\d{2}:\d{2})\s+([\d,\.]+)\s+(\d+)\s+([\d,\.]+)$', stripped)
        if match:
            prog = match.group(1).strip()
            if prog:
                current_prog = prog
            
            rows.append({
                "Date": match.group(3),
                "Day": match.group(4),
                "Air Time": match.group(5),
                "LEN": match.group(7),
                "Rate (INR)": match.group(6),
                "Program": current_prog,
                "Spot Copy": current_caption
            })
            continue

        # Looser pattern: SNo Date Day Time Rate Duration Amount (no program prefix)
        match2 = re.search(r'^(\d{4,6})\s+(\d{2}/\d{2})\s+([A-Za-z]{3})\s+(\d{2}:\d{2}:\d{2})\s+([\d,\.]+)\s+(\d+)\s+([\d,\.]+)$', stripped)
        if match2:
            rows.append({
                "Date": match2.group(2),
                "Day": match2.group(3),
                "Air Time": match2.group(4),
                "LEN": match2.group(6),
                "Rate (INR)": match2.group(5),
                "Program": current_prog,
                "Spot Copy": current_caption
            })
            continue

        # Even looser: Date Day Time Duration Rate (Sony Pal variant)
        match3 = re.search(r'^(\d{2}/\d{2}(?:/\d{2,4})?)\s+([A-Za-z]{3})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,\.]+)$', stripped)
        if match3:
            rows.append({
                "Date": match3.group(1),
                "Day": match3.group(2),
                "Air Time": match3.group(3),
                "LEN": match3.group(4),
                "Rate (INR)": match3.group(5),
                "Program": current_prog,
                "Spot Copy": current_caption
            })
            continue
        
        # Caption/product description line (before or between data rows)
        if not re.search(r'\d{2}/\d{2}', stripped) and not re.search(r'^LEAF$', stripped):
            if re.search(r'[A-Z]{3,}', stripped):
                current_caption = stripped
    
    return rows

# ==========================================
# EENADU / ETV PARSER
# ==========================================
def extract_eenadu_header(full_text):
    import re
    header = {"Broadcaster Name": "EENADU TELEVISION PRIVATE LIMITED"}
    chan = re.search(r'Channel Name:\s*(.*?)\s+(?:Invoice No|Reference Invoice No)', full_text, re.IGNORECASE)
    if not chan:
        chan = re.search(r'Channel Name\s*:\s*(.*?)\s*\n', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    inv = re.search(r'Invoice No\.:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice Date:\s*(.*?)\n', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    
    lines = full_text.split('\n')
    ro_val = ""
    for line in lines:
        if "RO No." in line or "Ref.RO.No" in line:
            ro_m = re.search(r'(?:RO No\.|Ref\.RO\.No)\s*:\s*(\S+)', line, re.IGNORECASE)
            if ro_m:
                ro_val = clean(ro_m.group(1))
                break
    header["PO Number"] = ro_val
    
    adv = re.search(r'ADVERTISER NAME\n(.*?Karnataka|.*?\d{6})', full_text, re.IGNORECASE | re.DOTALL)
    if adv:
        adv_name = adv.group(1).replace("\n", " ")
        # Strip 6-digit pin codes and state names
        adv_name = re.sub(r'\b\d{6}\b', '', adv_name)
        adv_name = re.sub(r'(Karnataka|Andhra Pradesh|Telangana|Tamil Nadu|Maharashtra|Delhi|Mumbai|Bangalore|Kolkata)', '', adv_name, flags=re.IGNORECASE)
        adv_name = clean(adv_name).rstrip("-")
        header["Advertiser Name"] = adv_name.strip()
    else:
        header["Advertiser Name"] = ""

    agency_match = re.search(r'ORIGINAL INVOICE TO\n(.*?(?:PVT LTD|PRIVATE LIMITED).*?)(?:\n|$)', full_text, re.IGNORECASE | re.DOTALL)
    if agency_match:
        ag = agency_match.group(1)
        ag = re.sub(r'(RO No\.:.*?\n|ID No\.:.*?\n)', '', ag)
        header["Agency Name"] = clean(ag)
    else:
        header["Agency Name"] = ""

    bp = re.search(r'(\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+to\s+(\d{2}\s+[A-Za-z]{3}\s+\d{4})', full_text)
    header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}" if bp else ""

    return header

def extract_eenadu_rows(full_text):
    import re
    rows = []
    lines = full_text.split('\n')
    start_idx = -1

    # Try multiple header anchors
    for i, line in enumerate(lines):
        if "Srl.No" in line and ("Brand" in line or "Duration" in line or "Amount" in line):
            start_idx = i + 1
            # Skip sub-header line like "Date Time"
            if start_idx < len(lines) and re.match(r'^\s*(Date|Telecast)\s+(Time|Date)\s*$', lines[start_idx].strip(), re.IGNORECASE):
                start_idx += 1
            break
        if re.search(r'SrNo.*Brand.*Caption', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if re.search(r'Sr\.?\s*No.*Programme.*TelecastDate', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if re.search(r'Sr\.?\s*No.*Date.*Time.*Duration', line, re.IGNORECASE):
            start_idx = i + 1
            break

    if start_idx == -1:
        start_idx = 0

    SKIP_RE = re.compile(r'(Total|Grand\s+Total|Taxable|RUPEE|Net\s+Amount|CGST|SGST|HSN|'
                          r'Invoice|Agency|Advertiser|Channel|Page\s+\d|IRN|Beneficiary|'
                          r'^Date\s+Time$|TELECAST\s+CERTIFICATE|Business\s+Region|'
                          r'certify\s+that|activity\s+period|Telecast\s+Telecast|'
                          r'EENADU\s+TELEVISION|III\s+Floor|Email:|CIN\s+No|PAN\s+No|'
                          r'Phone\s+No|GSTN\s+No|Service\s+Description|Region\s+Name)', re.IGNORECASE)
    # Date: accept 2 or 4 digit year
    DT = r'\d{2}/\d{2}/\d{2,4}'

    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue
        # Skip continuation lines (don't start with a digit = SrNo)
        if not re.match(r'^\d', stripped):
            continue

        # Pattern E: ETV Cinema - SNo ... RODate(4yr) ... Programme TelecastDate(2yr) Time Duration Amount
        # Example: '1 TETLEY TETLEY EVERYBODYCAN RO/ECS/2023-06- JUN2023/TVBRO/022 06/06/2023 MATINE SHOW MOVIE 09/06/23 13:30:29 45 2736'
        me = re.search(r'^(\d+)\s+(.*?)\s+(\d{2}/\d{2}/\d{4})\s+(.*?)\s+(\d{2}/\d{2}/\d{2,4})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,\.]+)\s*$', stripped)
        if me:
            brand_caption = me.group(2).strip()
            programme = me.group(4).strip()
            words = brand_caption.split()
            brand = " ".join(words[:2]) if len(words) > 2 else brand_caption
            caption = " ".join(words[2:]) if len(words) > 2 else ""
            rows.append({
                "Date": me.group(5), "Air Time": me.group(6),
                "LEN": me.group(7), "Rate (INR)": me.group(8),
                "Program": programme.title(), "Spot Copy": caption, "Brand": brand
            })
            continue

        # Pattern A: Standard ETV format - SNo Brand+Caption+Program Date Time Duration Amount
        match = re.search(r'^\d+\s+(.*?)\s+(' + DT + r')\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\d,\.]+)$', stripped)
        if match:
            left = match.group(1).strip()
            brand, caption, prog = _split_eenadu_left(left)
            rows.append({
                "Date": match.group(2), "Air Time": match.group(3),
                "LEN": match.group(4), "Rate (INR)": match.group(5),
                "Program": prog.title(), "Spot Copy": caption, "Brand": brand
            })
            continue

        # Pattern C: Date Time Program Duration Amount (no SNo)
        mc = re.search(r'^(' + DT + r')\s+(\d{2}:\d{2}:\d{2})\s+(.*?)\s+(\d+)\s+([\d,\.]+)$', stripped)
        if mc:
            rows.append({
                "Date": mc.group(1), "Air Time": mc.group(2),
                "Program": mc.group(3).strip(), "LEN": mc.group(4), "Rate (INR)": mc.group(5),
            })
            continue

        # Pattern D: SNo Content Date Time Duration (no amount - telecast cert)
        md = re.search(r'^\d+\s+(.*?)\s+(' + DT + r')\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s*$', stripped)
        if md:
            rows.append({
                "Date": md.group(2), "Air Time": md.group(3),
                "Program": md.group(1).strip(), "LEN": md.group(4), "Rate (INR)": "",
            })
            continue

    return rows


def _split_eenadu_left(left):
    """Split combined Brand+Caption+Program text from Eenadu rows."""
    brand = ""
    caption = ""
    prog = ""
    known_brands = ["CHAKRA GOLD PREMIUM LEAF", "CHAKRA GOLD", "TATA TEA PREMIUM",
                     "TATA TEA GOLD", "TATA SALT", "TATA CONSUMER", "CHK CARE"]
    for kb in known_brands:
        if left.upper().startswith(kb):
            brand = kb
            left = left[len(kb):].strip()
            break
    if not brand:
        words = left.split()
        brand = " ".join(words[:3])
        left = " ".join(words[3:])
    if " THE CHOICE OF " in left:
        parts = left.split(" THE CHOICE OF ")
        caption = parts[0] + " THE CHOICE OF"
        prog = parts[1].strip()
    else:
        words = left.split()
        if len(words) >= 4:
            caption = " ".join(words[:-2])
            prog = " ".join(words[-2:])
        else:
            caption = left
            prog = ""
    return brand, caption, prog

# ==========================================
# B4U PARSER
# ==========================================
def extract_b4u_header(full_text):
    import re
    header = {"Broadcaster Name": "B4U BROADBAND (INDIA) PRIVATE LIMITED"}
    inv = re.search(r'Invoice Number:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice Date:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    chan = re.search(r'Channel:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    ro = re.search(r'AGENCY REF#:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    
    # B4U structure usually has "BILLED TO ADVERTISER: INVOICE DETAILS"
    # followed by "GROUP M MEDIA INDIA PVT LTD - TATA CONSUMER Invoice Number:"
    billed_match = re.search(r'BILLED TO\s+ADVERTISER:.*?\n(.*?)\n', full_text, re.IGNORECASE)
    if billed_match:
        names = billed_match.group(1)
        if " Invoice Number" in names:
            names = names.split(" Invoice Number")[0]
        if " - " in names:
            parts = names.split(" - ", 1)
            header["Agency Name"] = clean(parts[0])
            header["Advertiser Name"] = clean(parts[1])
        else:
            header["Agency Name"] = clean(names)
            header["Advertiser Name"] = clean(names)
    else:
        header["Advertiser Name"] = ""
        header["Agency Name"] = ""

    bp = re.search(r'Invoice Period:\s*(.*?)to\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    if bp:
        header["Billing Period"] = f"{clean(bp.group(1))} to {clean(bp.group(2))}"
    else:
        header["Billing Period"] = ""
        
    return header

def extract_b4u_rows(full_text):
    rows = []
    lines = full_text.split('\n')
    start_idx = -1
    for i, line in enumerate(lines):
        if "Date" in line and "Brand" in line and "Caption" in line:
            # Header is 2 lines: "Date Brand Caption Rate Duration Count Total Taxable Value"
            #                     "                                      Duration"
            start_idx = i + 2
            break
    if start_idx == -1: return rows
    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped:
            continue
        if re.search(r'^(Total|Value of|Add:|In words)', stripped, re.IGNORECASE):
            break
        match = re.search(r'^(\d{2}-\d{2}-\d{4})\s+(.*?)\s+INR\.([\d,\.]+)\s+(\d+)\s+(\d+)\s+(\d+)\s+INR\.([\d,\.]+)$', stripped)
        if match:
            spot_copy_raw = match.group(2)
            brand = ""
            caption = spot_copy_raw
            # Split brand and caption: "TATA TEA GOLD TTG GOLD LEAF HSM JUN 23 - 25 SEC"
            # Brand is before the caption descriptor
            if " - " in spot_copy_raw:
                # Find last " - " which separates caption from duration text
                last_dash = spot_copy_raw.rfind(" - ")
                # Check if text after dash is duration-like ("25 SEC")
                after_dash = spot_copy_raw[last_dash+3:].strip()
                if re.match(r'^\d+\s*SEC', after_dash, re.IGNORECASE):
                    spot_copy_raw = spot_copy_raw[:last_dash].strip()
            
            # Now split Brand from Caption
            parts = spot_copy_raw.split()
            if len(parts) > 3:
                brand = " ".join(parts[:3])
                caption = " ".join(parts[3:])
            else:
                brand = spot_copy_raw
                caption = spot_copy_raw
                
            rows.append({
                "Date": match.group(1),
                "Air Time": "",
                "Spot Copy": caption,
                "Brand": brand,
                "Rate (INR)": match.group(3),
                "LEN": match.group(6),
                "Program": "",
            })
    return rows

# ==========================================
# JAYA NETWORK PARSER
# ==========================================
def extract_jaya_header(full_text):
    header = {}
    lines = full_text.split('\n')

    # Broadcaster detection
    if re.search(r'J\s*MOVIE|MAVIS\s+SATCOM', full_text, re.IGNORECASE):
        header["Broadcaster Name"] = "MAVIS SATCOM LIMITED (J MOVIES)"
    elif re.search(r'JAYA\s+MOVIES', full_text, re.IGNORECASE):
        header["Broadcaster Name"] = "JAYA TV NETWORK (JAYA MOVIES)"
    else:
        header["Broadcaster Name"] = "JAYA TV NETWORK"

    # Invoice Number - try multiple patterns
    inv = re.search(r'Invoice\s*(?:No|Number)\s*\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""

    # Invoice Date
    inv_date = re.search(r'Invoice\s*Date\s*:\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{4})', full_text, re.IGNORECASE)
    if not inv_date:
        inv_date = re.search(r'Date\s*:\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{4})', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""

    # PO / RO Number
    ro = re.search(r'(?:RO|R\.O\.?)\s*No\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""

    # Agency - handle "Agency :GROUP M MEDIA..." format
    agency = re.search(r'Agency\s*(?:Name)?\s*:\s*(.*?)(?:\s+Advertiser|\s+Invoice|\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""

    # Advertiser
    adv = re.search(r'(?:Advertiser|Client)\s*(?:Name)?\s*:\s*(.*?)(?:\s+Invoice\s+Number|\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""

    # Channel
    chan = re.search(r'Channel\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""

    brand = re.search(r'Brand\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    header["Billing Period"] = ""
    header["Station Relation"] = ""
    return header

def extract_jaya_rows(full_text):
    rows = []
    lines = full_text.split('\n')

    # Find table header - try multiple anchor phrases
    start_idx = 0  # default: scan all lines
    for i, line in enumerate(lines):
        if re.search(r'(Date|Dt).*Time.*(Program|Caption|Spot|Programme)', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if re.search(r'S\.?\s*No.*Date|Sr\.?\s*No.*Date', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if re.search(r'Date\s+Day\s+Time', line, re.IGNORECASE):
            start_idx = i + 1
            break

    SKIP_RE = re.compile(r'(Total|Grand\s+Total|Taxable|RUPEE|CGST|SGST|Net\s+Amount|'
                          r'Invoice\s+(No|Date)|Agency|Advertiser|Channel|Brand|Page\s+\d|'
                          r'IRN\s+NO|ACK\s+No|State\s*Code|State:-|Branch\s*:|Due\s+Date|'
                          r'^\s*\d+\s+of\s+\d+\s*$|TAX\s+INVOICE|TELECAST\s+CERTIFICATE|'
                          r'MAVIS\s+SATCOM|J\s+MOVIE|INVOICE\s+CUM)', re.IGNORECASE)
    DF = r'(?:\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}|\d{2}[- ][A-Za-z]{3}[- ]\d{4})'
    AMT = r'[\d,]+(?:\.\d{1,2})?'

    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue

        # P5: Date Day Time Type Duration Programme Rate Value  (J Movies/Jaya Movies format)
        # Example: '20/08/2023 Sun 19:15:19 Paid 15 SOLAI KUYIL 165.00 247.50'
        m5 = re.search(r'^(\d{2}/\d{2}/\d{4})\s+([A-Za-z]{3})\s+(\d{2}:\d{2}:\d{2})\s+\w+\s+(\d{1,3})\s+(.*?)\s+(' + AMT + r')\s+(' + AMT + r')$', stripped)
        if m5:
            rows.append({"Date": m5.group(1), "Day": m5.group(2), "Air Time": m5.group(3),
                         "Program": m5.group(5).strip(), "LEN": m5.group(4), "Rate (INR)": m5.group(7)})
            continue

        # P6: Date Day Time Type Duration Programme (NO amounts - telecast cert only)
        # Example: '20/08/2023 Sun 19:15:19 Paid 15 SOLAI KUYIL'
        m6 = re.search(r'^(\d{2}/\d{2}/\d{4})\s+([A-Za-z]{3})\s+(\d{2}:\d{2}:\d{2})\s+\w+\s+(\d{1,3})\s+([A-Za-z].*?)\s*$', stripped)
        if m6 and not re.search(r'^\d', m6.group(5)):
            rows.append({"Date": m6.group(1), "Day": m6.group(2), "Air Time": m6.group(3),
                         "Program": m6.group(5).strip(), "LEN": m6.group(4), "Rate (INR)": ""})
            continue

        # P1: Date HH:MM:SS Program/Caption Duration Rate
        m = re.search(r'^(' + DF + r')\s+(\d{2}:\d{2}:\d{2})\s+(.*?)\s+(\d{1,4})\s+(' + AMT + r')$', stripped)
        if m:
            rows.append({"Date": m.group(1), "Air Time": m.group(2),
                         "Program": m.group(3).strip(), "LEN": m.group(4), "Rate (INR)": m.group(5)})
            continue

        # P2: SNo Date Time Program Duration Rate
        m2 = re.search(r'^\d+\s+(' + DF + r')\s+(\d{2}:\d{2}:\d{2})\s+(.*?)\s+(\d{1,4})\s+(' + AMT + r')$', stripped)
        if m2:
            rows.append({"Date": m2.group(1), "Air Time": m2.group(2),
                         "Program": m2.group(3).strip(), "LEN": m2.group(4), "Rate (INR)": m2.group(5)})
            continue

        # P3: Date Time Duration Rate (no program text)
        m3 = re.search(r'^(' + DF + r')\s+(\d{2}:\d{2}:\d{2})\s+(\d{1,4})\s+(' + AMT + r')$', stripped)
        if m3:
            rows.append({"Date": m3.group(1), "Air Time": m3.group(2),
                         "Program": "", "LEN": m3.group(3), "Rate (INR)": m3.group(4)})
            continue

        # P4: Program Date Time Duration Rate (program text first)
        m4 = re.search(r'^(.+?)\s+(' + DF + r')\s+(\d{2}:\d{2}:\d{2})\s+(\d{1,4})\s+(' + AMT + r')$', stripped)
        if m4:
            rows.append({"Date": m4.group(2), "Air Time": m4.group(3),
                         "Program": m4.group(1).strip(), "LEN": m4.group(4), "Rate (INR)": m4.group(5)})
            continue

        # P7: Date Time Duration (minimal - no amount, no program)
        m7 = re.search(r'^(' + DF + r')\s+(\d{2}:\d{2}:\d{2})\s+(\d{1,4})\s*$', stripped)
        if m7:
            rows.append({"Date": m7.group(1), "Air Time": m7.group(2),
                         "Program": "", "LEN": m7.group(3), "Rate (INR)": ""})
            continue

    return rows


# ==========================================
# POLIMER NETWORK PARSER
# ==========================================

def extract_polimer_header(full_text):
    header = {"Broadcaster Name": "POLIMER MEDIA PVT. LTD."}
    inv = re.search(r'Invoice No\.\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice Date:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    ro = re.search(r'R\.O\. Number:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    
    agency = re.search(r'Agency Name\s*:\s*(.*?)(?:\s+Invoice No|\n)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    
    adv = re.search(r'Client Name:\s*(.*?)(?:\s+Invoice Period|\n)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    
    # Channel Name
    chan = re.search(r'Channel\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    if not chan:
        # Try known Polimer channel names
        if "POLIMER" in full_text.upper():
            header["Channel Name"] = "Polimer TV"
        else:
            header["Channel Name"] = ""
    else:
        header["Channel Name"] = clean(chan.group(1))
    
    brand = re.search(r'Brand:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    
    bp = re.search(r'Invoice Period:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Billing Period"] = clean(bp.group(1)) if bp else ""
    header["Station Relation"] = ""
    
    return header

def extract_polimer_rows(full_text):
    rows = []
    lines = full_text.split('\n')
    start_idx = -1
    
    # Find TELECAST CERTIFICATE section first (page 2+), then find data table header
    telecast_start = -1
    for i, line in enumerate(lines):
        if "TELECAST CERTIFICATE" in line.upper():
            telecast_start = i
            break
    
    search_from = telecast_start if telecast_start >= 0 else 0
    for i in range(search_from, len(lines)):
        line = lines[i]
        if re.search(r'Sr\.?\s*No', line, re.IGNORECASE) and re.search(r'(CAPTION|PROGRAMME|Program)', line, re.IGNORECASE):
            # Skip sub-header line (Type DATE Time)
            if i + 1 < len(lines) and re.search(r'(Type|DATE|Time)', lines[i+1], re.IGNORECASE):
                start_idx = i + 2
            else:
                start_idx = i + 1
            break
            
    if start_idx == -1: return rows
    for i in range(start_idx, len(lines)):
        line = lines[i].strip()
        if not line:
            continue
        if re.search(r'(Page \d+|NO\.\d+|BALAJI NAGAR)', line, re.IGNORECASE):
            continue
        if re.search(r'^PM/', line):
            continue
        
        # Pattern: SNo CAPTION PROGRAMME DATE TIME DUR
        # Example: 1 CGC INCREASED GRAMMAGE PROMO MOVIE PD 01-Aug-2023 19:59:30 15
        m = re.search(r'^(\d+)\s+(.*?)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(\d+)$', line)
        if m:
            left = m.group(2)
            # Split caption and program: Program is last 1-2 words + "PD"
            # E.g. "CGC INCREASED GRAMMAGE PROMO MOVIE PD" 
            #   -> Caption: "CGC INCREASED GRAMMAGE PROMO", Program: "MOVIE PD"
            prog = ""
            caption = left
            # Greedy match: capture as much as possible for caption, minimal for program
            pd_match = re.search(r'^(.+)\s+(\w+\s+PD)\s*$', left, re.IGNORECASE)
            if pd_match:
                caption = pd_match.group(1).strip()
                prog = pd_match.group(2).strip()
            
            rows.append({
                "Date": m.group(3), "Air Time": m.group(4),
                "LEN": m.group(5), "Rate (INR)": "",
                "Program": prog, "Spot Copy": caption
            })
            continue
        
        # Skip continuation lines like "15SECS EDIT"
    
    return rows

# ==========================================
# RAJ TV PARSER
# ==========================================
def extract_raj_header(full_text):
    header = {"Broadcaster Name": "RAJ TELEVISION NETWORK LIMITED"}
    inv = re.search(r'Invoice No\.?\s*:?\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    # Invoice Date: "Invoice Date Nov 30, 2023" or "Invoice Date : dd/mm/yyyy"
    inv_date = re.search(r'Invoice Date\s*:?\s*(\w+\s+\d{1,2},?\s+\d{4}|\d{2}[/\-\.]\d{2}[/\-\.]\d{4})', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    # RO No
    ro = re.search(r'RO\s+No\s+(\S+)', full_text, re.IGNORECASE)
    if not ro:
        ro = re.search(r'RO\s+Ref\.?No\s+(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    # Agency
    agency = re.search(r'(?:Agency/Client|Agency)\s*\n\s*(.*?)(?:\s+RAJ/|\n)', full_text, re.IGNORECASE)
    if not agency:
        agency = re.search(r'Agency\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    # Advertiser
    adv = re.search(r'ADVERTISER\s*\n\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    if not adv:
        adv = re.search(r'Advertiser\s*:?\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    # Channel
    chan = re.search(r'CHANNEL\s+(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    if not chan:
        chan = re.search(r'Channel\s*:?\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    # Brand
    brand = re.search(r'Brand\s+(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    # Billing Period from Activity Period
    bp = re.search(r'Activity Period\s+([\d/]+)\s*[-–]\s*([\d/]+)', full_text, re.IGNORECASE)
    if bp:
        header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}"
    else:
        bp2 = re.search(r'period\s+from\s+([\d/]+)\s*[-–]\s*([\d/]+)', full_text, re.IGNORECASE)
        header["Billing Period"] = f"{bp2.group(1)} to {bp2.group(2)}" if bp2 else ""
    header["Station Relation"] = ""
    return header

def extract_raj_rows(full_text):
    """
    Raj TV telecast certificate format:
    dd-mm-yyyy PROGRAM_NAME CAPTION ( HH:MM:SS TYPE DUR RATE AMOUNT
    with possible continuation lines like '29TH JUNE )' or 'DOCUMENTRY 29TH JUNE )'
    """
    rows = []
    lines = full_text.split('\n')
    start_idx = -1

    # Find telecast certificate table header
    for i, line in enumerate(lines):
        if re.search(r'DATE\s+OF.*PROGRAM\s+NAME.*PRODUCT.*TIME\s+OF', line, re.IGNORECASE):
            start_idx = i + 1
            # Skip sub-header line like "TELECAST SECS"
            if start_idx < len(lines) and re.search(r'TELECAST\s+SECS', lines[start_idx], re.IGNORECASE):
                start_idx += 1
            break
        if "Date" in line and "Time" in line and "Program" in line:
            start_idx = i + 1
            break

    if start_idx == -1:
        return rows

    SKIP_RE = re.compile(r'(Total|Grand|Taxable|RUPEE|Net\s+Amount|CGST|SGST|FOR RAJ|TELECAST CERTIFICATE|Agency|Advertiser|Channel|State|GST|BANGALURU|M\.G\.ROAD|FLOOR|CHAMBERS|Invoice|RO\s+No|RO\s+Ref|Brand|Activity|Nature|Whether|Realese|certify|Spots\s+Type|IRN|ACK)', re.IGNORECASE)

    # Pattern: dd-mm-yyyy PROGRAM CAPTION ( HH:MM:SS TYPE DUR RATE AMOUNT
    # Example: '01-11-2023 TFF-PARISAM POTTACHU CG TN MURAI - KOLAM EDIT - TAMIL ( 11:08:37 PAID 20 70 140'
    PAT = re.compile(
        r'^(\d{2}-\d{2}-\d{4})\s+'    # Date
        r'(.*?)\s+'                     # Program + Caption
        r'\(\s*(\d{1,2}:\d{2}:\d{2})\s+'  # ( Time
        r'(\w+)\s+'                     # Type (PAID/FREE)
        r'(\d+)\s+'                     # Duration
        r'([\d,]+)\s+'                  # Rate
        r'([\d,]+)\s*$'                # Amount
    )

    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue
        # Skip continuation lines (don't start with date)
        if not re.match(r'^\d{2}-\d{2}-\d{4}', stripped):
            continue

        m = PAT.search(stripped)
        if m:
            left = m.group(2).strip()
            # Split program and caption - caption usually starts with "CG "
            program = left
            caption = ""
            cg_match = re.search(r'\s+(CG\s+.*)$', left)
            if cg_match:
                program = left[:cg_match.start()].strip()
                caption = cg_match.group(1).strip()

            rows.append({
                "Date": m.group(1),
                "Air Time": m.group(3),
                "Program": program,
                "Spot Copy": caption,
                "LEN": m.group(5),
                "Rate (INR)": m.group(6),
            })

    return rows

# ==========================================
# VENDHAR TV PARSER
# ==========================================
def extract_vendhar_header(full_text):
    header = {"Broadcaster Name": "SRN MEDIA VISION PVT. LTD."}
    inv = re.search(r'Invoice No\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice Date\s*:\s*(\S+)', full_text, re.IGNORECASE)
    if not inv_date:
        inv_date = re.search(r'Date\s*:\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{4})', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    ro = re.search(r'(?:RO|R\.O\.?)\s*No\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    agency = re.search(r'Agency\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    adv = re.search(r'(?:Advertiser|Client)\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    chan = re.search(r'Channel\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    brand = re.search(r'Brand\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    header["Billing Period"] = ""
    header["Station Relation"] = ""
    return header

def extract_vendhar_rows(full_text):
    rows = []
    lines = full_text.split('\n')
    start_idx = -1
    for i, line in enumerate(lines):
        if re.search(r'S\.?\s*No|Sr\.?\s*No|Date.*Time.*Program|Date.*Time.*Caption', line, re.IGNORECASE):
            start_idx = i + 1
            break
    if start_idx == -1:
        return rows
    for line in lines[start_idx:]:
        line = line.strip()
        if not line or "Total" in line or "Grand Total" in line or "Taxable" in line:
            continue
        # Try: Date Time Program Duration Rate
        m = re.search(r'(\d{2}[/\-\.]\d{2}[/\-\.]\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(.*?)\s+(\d+)\s+([\d,\.]+)$', line)
        if m:
            rows.append({
                "Date": m.group(1), "Air Time": m.group(2),
                "Program": m.group(3).strip(), "LEN": m.group(4),
                "Rate (INR)": m.group(5),
            })
            continue
        # Try: SNo Date Time Program Duration Rate  
        m2 = re.search(r'^\d+\s+(\d{2}[/\-\.]\d{2}[/\-\.]\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(.*?)\s+(\d+)\s+([\d,\.]+)$', line)
        if m2:
            rows.append({
                "Date": m2.group(1), "Air Time": m2.group(2),
                "Program": m2.group(3).strip(), "LEN": m2.group(4),
                "Rate (INR)": m2.group(5),
            })
    return rows

# ==========================================
# VANITHA TV PARSER
# ==========================================
def extract_vanitha_header(full_text):
    header = {"Broadcaster Name": "RACHANA TELEVISION PVT. LTD."}
    inv = re.search(r'Invoice No\.\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice Date\s*:\s*(\S+)', full_text, re.IGNORECASE)
    if not inv_date:
        inv_date = re.search(r'Date\s*:\s*(\d{2}[/\-\.]\d{2}[/\-\.]\d{4})', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    ro = re.search(r'(?:RO|R\.O\.?)\s*No\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    agency = re.search(r'Agency\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    adv = re.search(r'(?:Advertiser|Client)\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    chan = re.search(r'Channel\s*(?:Name)?\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    brand = re.search(r'Brand\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    header["Billing Period"] = ""
    header["Station Relation"] = ""
    return header

def extract_vanitha_rows(full_text):
    rows = []
    lines = full_text.split('\n')

    SKIP_RE = re.compile(r'(Total|Grand\s+Total|Taxable|RUPEE|Net\s+Amount|CGST|SGST|IGST|'
                          r'RACHANA\s+TELEVISION|Plot\s+#|Jubilee|Ph:|Fax:|Original|'
                          r'TAX\s+INVOICE|Spot\s+Release|Sl\s+#\s+Particulars|'
                          r'Beneficiary|Account\s+Number|Interest|payments|complaint|'
                          r'disputes|terms\s+and|Authorised|INR\s+)', re.IGNORECASE)

    # Invoice page pattern: SNo dd-Mon-yy HH:MM:SS - HH:MM:SS Caption Rate Dur Amount
    # Example: '1 16-Jun-23 08:00:00 - 10:00:00 Every Body Can-Bus Garage-AP-25 Sec 165.00 25 412.50'
    PAT_INV = re.compile(
        r'^(\d+)\s+'
        r'(\d{2}-[A-Za-z]{3}-\d{2,4})\s+'
        r'(\d{1,2}:\d{2}:\d{2})\s*-\s*(\d{1,2}:\d{2}:\d{2})\s+'
        r'(.*?)\s+'
        r'([\d,]+\.\d{2})\s+'
        r'(\d+)\s+'
        r'([\d,]+\.\d{2})\s*$'
    )

    # Telecast cert pattern: SNo dd-Mon-yy Program - Caption H:MM:SS H:MM:SS Dur
    # Example: '1 16-Jun-23 Amani - Every Body Can-Bus Garage-AP-25 Sec 9:25:10 9:25:35 25'
    PAT_TC = re.compile(
        r'^(\d+)\s+'
        r'(\d{2}-[A-Za-z]{3}-\d{2,4})\s+'
        r'(.*?)\s+'
        r'(\d{1,2}:\d{2}:\d{2})\s+'
        r'(\d{1,2}:\d{2}:\d{2})\s+'
        r'(\d+)\s*$'
    )

    for line in lines:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue

        # Try invoice page pattern first
        m = PAT_INV.search(stripped)
        if m:
            time_range = f"{m.group(3)} - {m.group(4)}"
            caption = m.group(5).strip()
            # Remove "Sec" suffix from caption
            caption = re.sub(r'\s*\d+\s*Sec\s*$', '', caption, flags=re.IGNORECASE).strip()
            rows.append({
                "Date": m.group(2), "Air Time": m.group(3),
                "Time Range/Sales Unit": time_range,
                "Spot Copy": caption,
                "Rate (INR)": m.group(6),
                "LEN": m.group(7),
                "Program": "",
            })
            continue

        # Try telecast cert pattern
        m2 = PAT_TC.search(stripped)
        if m2:
            content = m2.group(3).strip()
            # Split Program - Caption
            program = content
            caption = ""
            dash_split = re.split(r'\s+-\s+', content, maxsplit=1)
            if len(dash_split) == 2:
                program = dash_split[0].strip()
                caption = dash_split[1].strip()
                # Remove "Sec" suffix
                caption = re.sub(r'\s*\d+\s*Sec\s*$', '', caption, flags=re.IGNORECASE).strip()

            rows.append({
                "Date": m2.group(2), "Air Time": m2.group(4),
                "Program": program, "Spot Copy": caption,
                "LEN": m2.group(6), "Rate (INR)": "",
            })
            continue

    return rows

# ==========================================
# GENERIC / FALLBACK PARSER
# (Works for Goldmines, News Nation, Enter10,
#  Shemaroo, Matrix, Vasanth, and similar)
# ==========================================
def extract_generic_header(full_text, format_name="Unknown"):
    header = {}
    lines = full_text.split('\n')
    
    # Try to find Broadcaster name from first few lines
    broadcaster = ""
    for line in lines[:10]:
        line = line.strip()
        if re.search(r'(pvt\.?\s*ltd|private\s+limited|limited|network|entertainment|media|broadcast)', line, re.IGNORECASE):
            broadcaster = clean(line)
            break
    header["Broadcaster Name"] = broadcaster if broadcaster else format_name

    # Generic field extraction patterns
    inv = re.search(r'(?:Invoice|Inv\.?|GST Inv\.?)\s*No\.?\s*:?\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    
    inv_date = re.search(r'(?:Invoice\s+)?Date\s*:?\s*(\d{2}[/\-\.]\w{3,9}[/\-\.]\d{2,4}|\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4})', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    
    ro = re.search(r'(?:RO|R\.O\.?|PO|P\.O\.?)\s*(?:No|Number|Ref)?\.?\s*:?\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""

    agency = re.search(r'Agency\s*(?:Name)?\s*:?\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    
    adv = re.search(r'(?:Advertiser|Client)\s*(?:Name)?\s*:?\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    
    chan = re.search(r'Channel\s*(?:Name)?\s*:?\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    
    brand = re.search(r'Brand\s*:?\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    
    header["Billing Period"] = ""
    header["Station Relation"] = ""
    return header

def extract_generic_rows(full_text):
    """
    Universal row extractor: tries multiple date/time patterns, no mandatory anchor.
    Handles: dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy, dd-Mon-yyyy, dd/mm/yy.
    Accepts amounts with or without decimals.
    """
    rows = []
    lines = full_text.split('\n')

    # Try to find table start, but fall back to scanning all lines
    start_idx = 0
    for i, line in enumerate(lines):
        if re.search(r'(S\.?\s*No|Sr\.?\s*No|Sl\.?\s*No)', line, re.IGNORECASE) and \
           re.search(r'(Date|Time|Program|Caption|Rate|Amount)', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if re.search(r'Date\s+Time\s+(?:Program|Caption|Duration)', line, re.IGNORECASE):
            start_idx = i + 1
            break

    SKIP_RE = re.compile(
        r'(^Total|Grand\s+Total|Taxable|RUPEE|Sub\s+Total|Net\s+Amount|CGST|SGST|HSN|'
        r'Invoice\s+(No|Date|Period)|Agency|Advertiser|Channel\s+Name|'
        r'^\d+\s+of\s+\d+$|Page\s+\d+)',
        re.IGNORECASE
    )
    DF = r'(?:\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}|\d{2}[- ][A-Za-z]{3}[- ]\d{4})'
    AMT = r'[\d,]+(?:\.\d{1,2})?'
    # Time pattern: HH:MM:SS or HH:MM:SS:FF (with frame counter)
    TM = r'\d{1,2}:\d{2}:\d{2}(?::\d{2})?'

    for line in lines[start_idx:]:
        line = line.strip()
        if not line or SKIP_RE.search(line):
            continue

        # P1: SNo Date Time Content Duration Rate
        m = re.search(r'^\d+\s+(' + DF + r')\s+(' + TM + r')\s+(.*?)\s+(\d{1,4})\s+(' + AMT + r')$', line)
        if m:
            rows.append({"Date": m.group(1), "Air Time": m.group(2),
                         "Program": m.group(3).strip(), "LEN": m.group(4), "Rate (INR)": m.group(5)})
            continue

        # P2: Date Time Content Duration Rate (no SNo)
        m2 = re.search(r'^(' + DF + r')\s+(' + TM + r')\s+(.*?)\s+(\d{1,4})\s+(' + AMT + r')$', line)
        if m2:
            rows.append({"Date": m2.group(1), "Air Time": m2.group(2),
                         "Program": m2.group(3).strip(), "LEN": m2.group(4), "Rate (INR)": m2.group(5)})
            continue

        # P3: SNo Content Date Day Time Duration Rate (Zee-style)
        m3 = re.search(r'^\d+\s+(.*?)\s+(\d{2}\.\d{2}\.\d{4})\s+([A-Z]{3})\s+(\d{2}:\d{2}:\d{2})\s+(\d{1,4})\s+(' + AMT + r')$', line)
        if m3:
            rows.append({"Date": m3.group(2), "Day": m3.group(3), "Air Time": m3.group(4),
                         "Program": m3.group(1).strip(), "LEN": m3.group(5), "Rate (INR)": m3.group(6)})
            continue

        # P4: Date(dd-Mon-yyyy) Time Content Duration Rate
        m4 = re.search(r'^(\d{2}-[A-Za-z]{3}-\d{4})\s+(' + TM + r')\s+(.*?)\s+(\d{1,4})\s+(' + AMT + r')$', line)
        if m4:
            rows.append({"Date": m4.group(1), "Air Time": m4.group(2),
                         "Program": m4.group(3).strip(), "LEN": m4.group(4), "Rate (INR)": m4.group(5)})
            continue

        # P5: SNo Date Time Duration Amount (program text missing)
        m5 = re.search(r'^\d+\s+(' + DF + r')\s+(' + TM + r')\s+(\d{1,4})\s+(' + AMT + r')$', line)
        if m5:
            rows.append({"Date": m5.group(1), "Air Time": m5.group(2),
                         "Program": "", "LEN": m5.group(3), "Rate (INR)": m5.group(4)})
            continue

        # P6: Content Date Time Duration Rate (content first – Matrix/Shemaroo)
        m6 = re.search(r'^(.+?)\s+(' + DF + r')\s+(' + TM + r')\s+(\d{1,4})\s+(' + AMT + r')$', line)
        if m6:
            rows.append({"Date": m6.group(2), "Air Time": m6.group(3),
                         "Program": m6.group(1).strip(), "LEN": m6.group(4), "Rate (INR)": m6.group(5)})
            continue

        # P7: SNo Content Date Time Duration (NO amount - telecast cert)
        m7 = re.search(r'^\d+\s+(.*?)\s+(' + DF + r')\s+(' + TM + r')\s+(\d{1,4})\s*$', line)
        if m7:
            rows.append({"Date": m7.group(2), "Air Time": m7.group(3),
                         "Program": m7.group(1).strip(), "LEN": m7.group(4), "Rate (INR)": ""})
            continue

        # P8: Date Time Content Duration (no amount, no SNo)
        m8 = re.search(r'^(' + DF + r')\s+(' + TM + r')\s+(.*?)\s+(\d{1,4})\s*$', line)
        if m8:
            rows.append({"Date": m8.group(1), "Air Time": m8.group(2),
                         "Program": m8.group(3).strip(), "LEN": m8.group(4), "Rate (INR)": ""})
            continue

    return rows


# ==========================================
# SHEMAROO PARSER
# ==========================================

def extract_shemaroo_header(full_text):
    header = {"Broadcaster Name": "SHEMAROO ENTERTAINMENT LIMITED"}
    inv = re.search(r'Invoice\s+No\.?\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice\s+Date\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    chan = re.search(r'Channel\s+Name\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    agency = re.search(r'Agency\s*:\s*(.*?)(?:\s+Invoice|\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    adv = re.search(r'Advertiser\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    ro = re.search(r'(?:Reference\s+Number|Agency\s+RO\s+Number)\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    bp = re.search(r'Invoice\s+Period\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    if not bp:
        bp = re.search(r'Invoicing\s+Period\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Billing Period"] = clean(bp.group(1)) if bp else ""
    header["Brand"] = ""
    header["Station Relation"] = ""
    return header


def extract_shemaroo_rows(full_text):
    """
    Shemaroo telecast cert format:
    SrNo dd-Mon-yyyy ContentName HH:MM:SS Caption Brand Duration
    Example: '1 13-Nov-2023 SHRAVANI 13:12:13 TT AGNI 20 SEC GURJIT BRAND RANGE TATA TEA AGNI 20'
    """
    rows = []
    lines = full_text.split('\n')
    start_idx = 0

    # Find telecast certificate header
    for i, line in enumerate(lines):
        if re.search(r'Sr\.?\s*No.*Telecast\s+Date|Telecast\s+Date.*Content\s+Name', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if re.search(r'TELECAST\s+CERTIFICATE', line, re.IGNORECASE):
            # Look for the table header after this
            for j in range(i+1, min(i+10, len(lines))):
                if re.search(r'Sr\.?\s*No', lines[j], re.IGNORECASE):
                    start_idx = j + 1
                    break

    SKIP_RE = re.compile(r'(Total|Grand|Taxable|Net\s+Amount|CGST|SGST|HSN|computer\s+generated|'
                          r'SHEMAROO\s+ENTERTAINMENT|Page\s+\d|Invoice\s+No|Invoice\s+Date|'
                          r'Channel\s+Name|Agency|Advertiser|Contract|Executive|Invoice\s+Period|'
                          r'Reference|Invoice\s+Type|PAN|State|Place\s+of|Invoicing)', re.IGNORECASE)

    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue
        # Skip continuation lines (don't start with digit)
        if not re.match(r'^\d', stripped):
            continue

        # Pattern: SrNo Date(dd-Mon-yyyy) ContentName Time ... Duration(last number)
        m = re.search(r'^(\d+)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+(.*?)\s+(\d{2}:\d{2}:\d{2})\s+(.*?)\s+(\d{1,3})\s*$', stripped)
        if m:
            content = m.group(3).strip()
            caption_brand = m.group(5).strip()
            # Try to extract brand from caption_brand (last known brand name)
            brand = ""
            caption = caption_brand
            brand_m = re.search(r'(TATA\s+TEA\s+\w+|TATA\s+SALT|TETLEY|CHAKRA\s+GOLD)', caption_brand, re.IGNORECASE)
            if brand_m:
                # Everything after the brand match is part of brand name
                brand_start = caption_brand.rfind(brand_m.group(0)[:4])
                if brand_start > 0:
                    caption = caption_brand[:brand_start].strip()
                    brand = caption_brand[brand_start:].strip()
                else:
                    brand = brand_m.group(0)

            rows.append({
                "Date": m.group(2), "Air Time": m.group(4),
                "Program": content, "LEN": m.group(6),
                "Spot Copy": caption, "Brand": brand,
                "Rate (INR)": "",
            })
            continue

    # Also try to extract rate info from invoice pages (pages 1-2 have rate data)
    # Look for lines like: '1 TATA TEA AGNI TT AGNI 20 SEC ... 1000 20 2000 3 60 6000.00'
    if not rows:
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            m2 = re.search(r'^(\d+)\s+(.*?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d,\.]+)\s*$', stripped)
            if m2:
                rows.append({
                    "Program": "", "LEN": m2.group(4),
                    "Rate (INR)": m2.group(8), "Spot Copy": m2.group(2).strip(),
                    "Date": "", "Air Time": "",
                })

    return rows


# ==========================================
# VASANTH TV PARSER
# ==========================================
def extract_vasanth_header(full_text):
    header = {"Broadcaster Name": "VASANTH & CO MEDIA NETWORK (P) LTD"}
    inv = re.search(r'INVOICE\s+No\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Date\s*:\s*(\d{2}[\-\.]\d{2}[\-\.]\d{4})', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    ro = re.search(r'R\.O\.No\.\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    agency = re.search(r'M/s\.\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    adv = re.search(r'Client\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    header["Channel Name"] = "Vasanth TV"
    brand = re.search(r'BRAND/SLOT\s*\n\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    if not brand:
        brand = re.search(r'Brand\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    bp = re.search(r'TC\s+Date\s*:\s*([\d\.]+)\s+TO\s+([\d\.]+)', full_text, re.IGNORECASE)
    header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}" if bp else ""
    header["Station Relation"] = ""
    return header

def extract_vasanth_rows(full_text):
    """
    Vasanth TV telecast certificate format:
    Description line (e.g. 'TAMIL CULTURE/LADY DOING')
    mm/dd/yyyy HH:MM:SS AM/PM HH:MM:SS AM/PM Rate LEN Amount BRAND
    Continuation line (e.g. 'RANGOLI/FOOD')
    """
    rows = []
    lines = full_text.split('\n')
    start_idx = -1

    # Find TC header line
    for i, line in enumerate(lines):
        if re.search(r'Date\s+Start\s+Time\s+End\s+Time\s+Rate', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if "TELECAST CERTIFICATE" in line.upper():
            start_idx = i + 1

    if start_idx == -1:
        return rows

    SKIP_RE = re.compile(r'(Total|Grand|VASANTH|No\.27|Phone|GSTIN|PAN|Railway|Kavery|Saidapet|'
                          r'Chennai|Authorised|Payment|HSN|INVOICE|DATE:|accounts@)', re.IGNORECASE)

    # Pattern: mm/dd/yyyy HH:MM:SS AM/PM HH:MM:SS AM/PM Rate LEN Amount BRAND
    # Example: '11/01/2023 08:06:13 AM 08:06:33 AM 80.00 20 160.00 TATA CHAKRA GOLD'
    PAT = re.compile(
        r'^(\d{2}/\d{2}/\d{4})\s+'
        r'(\d{1,2}:\d{2}:\d{2})\s+(AM|PM)\s+'
        r'(\d{1,2}:\d{2}:\d{2})\s+(AM|PM)\s+'
        r'([\d,]+\.\d{2})\s+'
        r'(\d+)\s+'
        r'([\d,]+\.\d{2})\s+'
        r'(.*?)\s*$'
    )

    prev_description = ""
    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue

        m = PAT.search(stripped)
        if m:
            brand = m.group(9).strip()
            rows.append({
                "Date": m.group(1),
                "Air Time": f"{m.group(2)} {m.group(3)}",
                "Rate (INR)": m.group(6),
                "LEN": m.group(7),
                "Spot Copy": prev_description,
                "Brand": brand,
                "Program": "",
            })
            prev_description = ""  # Reset
            continue

        # Non-data line = description for next row
        if not re.match(r'^\d', stripped) and len(stripped) > 3:
            prev_description = stripped

    return rows


# ==========================================
# GOLDMINES PARSER
# ==========================================
def extract_goldmines_header(full_text):
    header = {"Broadcaster Name": "GOLDMINES TELEFILMS PRIVATE LIMITED"}
    inv = re.search(r'Invoice\s+No\.\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice\s+Date\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    agency = re.search(r'Agency\s+Name\s*:\s*(.*?)(?:\s+Invoice|\n)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    adv = re.search(r'Advertiser\s+Name\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    if not adv:
        adv = re.search(r'Client\s+Name\s*:\s*(.*?)(?:\s+Invoice|\n)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    chan = re.search(r'Channel\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    brand = re.search(r'Brand\s+Name\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    ro = re.search(r'RO\s+No\.\s*&?\s*Date\s*:\s*(\S+)', full_text, re.IGNORECASE)
    if not ro:
        ro = re.search(r'Client\s+Ref\s+No\.\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    bp = re.search(r'Period\s*:\s*(\S+)\s+To\s+(\S+)', full_text, re.IGNORECASE)
    header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}" if bp else ""
    header["Station Relation"] = ""
    return header

def extract_goldmines_rows(full_text):
    """
    Goldmines telecast certificate format:
    SrNo TAPE_CAPTION TxDate TelecastTime PROGRAMME DUR
    Example: '1 TTG_GOLD_LEAF_HSM_JUN_23_HINDI_25_SEC 16-Oct-2023 14:17:22 CHHOTI BAHOO 25'
    """
    rows = []
    lines = full_text.split('\n')
    start_idx = -1

    for i, line in enumerate(lines):
        if re.search(r'Sr\s+No\s+TAPE\s+CAPTION.*Tx\s+DATE', line, re.IGNORECASE):
            start_idx = i + 1
            break
        if re.search(r'Sr\s*No.*Caption.*Date.*Time.*Program', line, re.IGNORECASE):
            start_idx = i + 1
            break

    if start_idx == -1:
        return rows

    SKIP_RE = re.compile(r'(Total|Grand|Corporate|Page\s+\d|GB-\d|Malad|Road|Mumbai|Palm\s+Court)', re.IGNORECASE)

    # Pattern: SrNo Caption Date Time Program Dur
    PAT = re.compile(
        r'^(\d+)\s+'
        r'(\S+)\s+'
        r'(\d{2}-[A-Za-z]{3}-\d{4})\s+'
        r'(\d{1,2}:\d{2}:\d{2})\s+'
        r'(.*?)\s+'
        r'(\d+)\s*$'
    )

    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue
        if not re.match(r'^\d+\s', stripped):
            continue

        m = PAT.search(stripped)
        if m:
            rows.append({
                "Date": m.group(3),
                "Air Time": m.group(4),
                "Spot Copy": m.group(2).replace('_', ' '),
                "Program": m.group(5).strip(),
                "LEN": m.group(6),
                "Rate (INR)": "",
            })

    return rows


# ==========================================
# ENTER10 / DANGAL PARSER
# ==========================================
def extract_enter10_header(full_text):
    header = {"Broadcaster Name": "ENTER 10 TELEVISION PVT. LTD."}
    inv = re.search(r'Invoice\s+No\.\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice\s+Date\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    agency = re.search(r'(?:Buyer\s+Name|Agency)\s*:\s*(.*?)(?:\s+Invoice|\n)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    adv = re.search(r'Client\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    chan = re.search(r'Channel\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    brand = re.search(r'Brand\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    ro = re.search(r'RO\s+NO\s*&?\s*Date\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    bp = re.search(r'Period\s+From\s+(\S+)\s+To\s+(\S+)', full_text, re.IGNORECASE)
    header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}" if bp else ""
    header["Station Relation"] = ""
    return header

def extract_enter10_rows(full_text):
    """
    Enter10 telecast certificate format:
    SrNo Caption TxDate TelecastTime Program Dur
    Example: '1 TT AGNI 20 SEC GURJIT BRAND RANGE 13.11.2023 16:52:57 JANAM JANAM KA SAATH 20'
    With continuation: 'ENDSLA RELAUNCH CLEN'
    """
    rows = []
    lines = full_text.split('\n')
    start_idx = -1

    for i, line in enumerate(lines):
        if re.search(r'Sr\.?\s*No\s+Caption\s+Tx\s+Date', line, re.IGNORECASE):
            start_idx = i + 1
            # Skip sub-header like 'Time'
            if start_idx < len(lines) and re.match(r'^\s*Time\s*$', lines[start_idx].strip()):
                start_idx += 1
            break

    if start_idx == -1:
        return rows

    SKIP_RE = re.compile(r'(Total|certify|telecasted|Enter\s+10|Authorized|Page)', re.IGNORECASE)

    # Pattern: SrNo Caption Date Time Program Dur
    PAT = re.compile(
        r'^(\d+)\s+'
        r'(.*?)\s+'
        r'(\d{2}\.\d{2}\.\d{4})\s+'
        r'(\d{2}:\d{2}:\d{2})\s+'
        r'(.*?)\s+'
        r'(\d+)\s*$'
    )

    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue
        if not re.match(r'^\d+\s', stripped):
            continue

        m = PAT.search(stripped)
        if m:
            rows.append({
                "Date": m.group(3),
                "Air Time": m.group(4),
                "Spot Copy": m.group(2).strip(),
                "Program": m.group(5).strip(),
                "LEN": m.group(6),
                "Rate (INR)": "",
            })

    return rows


# ==========================================
# NEWS NATION PARSER
# ==========================================
def extract_newsnation_header(full_text):
    header = {"Broadcaster Name": "NEWS NATION NETWORK PVT LTD."}
    inv = re.search(r'Invoice\s+No\s+(\S+)', full_text, re.IGNORECASE)
    header["Invoice Number"] = clean(inv.group(1)) if inv else ""
    inv_date = re.search(r'Invoice\s+Date\s+(\S+)', full_text, re.IGNORECASE)
    header["Invoice Date"] = clean(inv_date.group(1)) if inv_date else ""
    agency = re.search(r'Agency\s+Name\s*:\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Agency Name"] = clean(agency.group(1)) if agency else ""
    adv = re.search(r'Client\s+(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Advertiser Name"] = clean(adv.group(1)) if adv else ""
    chan = re.search(r'Channel\s*:?\-?\s*(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Channel Name"] = clean(chan.group(1)) if chan else ""
    brand = re.search(r'Brand\s+Name\s+(.*?)(?:\n|$)', full_text, re.IGNORECASE)
    header["Brand"] = clean(brand.group(1)) if brand else ""
    ro = re.search(r'RO\s+No\.\s*:\s*(\S+)', full_text, re.IGNORECASE)
    header["PO Number"] = clean(ro.group(1)) if ro else ""
    bp = re.search(r'Period\s+From\s+(\S+)\s+To\s+(\S+)', full_text, re.IGNORECASE)
    header["Billing Period"] = f"{bp.group(1)} to {bp.group(2)}" if bp else ""
    header["Station Relation"] = ""
    return header

def extract_newsnation_rows(full_text):
    """
    News Nation telecast certificate format:
    SrNo CAPTION PROGRAMME SPOT_TYPE CATEGORY DATE TIME DUR
    Example: '1 TT AGNI VANDANA BRAND RANGE ENDSLATE 20S NECEWS STATE PD RODP 12/11/2023 21:55:24 20'
    """
    rows = []
    lines = full_text.split('\n')
    start_idx = -1

    for i, line in enumerate(lines):
        if re.search(r'Sr\.?\s*No\.?\s+CAPTION\s+PROGRAMME', line, re.IGNORECASE):
            start_idx = i + 1
            # Skip sub-headers like 'NAME TYPE (SEC)' and 'DATE TIME'
            while start_idx < len(lines) and not re.match(r'^\s*\d+\s', lines[start_idx]):
                start_idx += 1
            break

    if start_idx == -1:
        return rows

    SKIP_RE = re.compile(r'(Total|Grand|News\s+Nation|CLIENT\s+COPY|TELECAST\s+CERT|Client\s+:|'
                          r'Address|Agency|Brand|Period|T\.O\.|Sr\.?\s*No|NAME\s+TYPE|DATE\s+TIME)', re.IGNORECASE)

    # Pattern: SrNo Caption+Programme+SpotType+Category Date Time Dur
    PAT = re.compile(
        r'^(\d+)\s+'
        r'(.*?)\s+'
        r'(\d{2}/\d{2}/\d{4})\s+'
        r'(\d{2}:\d{2}:\d{2})\s+'
        r'(\d+)\s*$'
    )

    for line in lines[start_idx:]:
        stripped = line.strip()
        if not stripped or SKIP_RE.search(stripped):
            continue
        if not re.match(r'^\d+\s', stripped):
            continue

        m = PAT.search(stripped)
        if m:
            content = m.group(2).strip()
            # Split: Caption is before PD/RODP, Program is after first program-like word
            # Example: 'TT AGNI VANDANA BRAND RANGE ENDSLATE 20S NECEWS STATE PD RODP'
            caption = content
            program = ""
            # Try to find program name (usually after ENDSLATE xxS/SEC)
            prog_m = re.search(r'(?:ENDSLATE?\s+\d+S(?:EC)?)\s+(.*?)(?:\s+PD\s+|\s+FREE\s+)', content, re.IGNORECASE)
            if prog_m:
                program = prog_m.group(1).strip()
                caption_end = content.find(prog_m.group(0))
                if caption_end > 0:
                    caption = content[:caption_end + len("ENDSLATE")].strip()
                    caption = re.sub(r'\s+\d+S(EC)?\s*$', '', caption, flags=re.IGNORECASE).strip()

            rows.append({
                "Date": m.group(3),
                "Air Time": m.group(4),
                "Spot Copy": caption,
                "Program": program,
                "LEN": m.group(5),
                "Rate (INR)": "",
            })

    return rows


# ==========================================

# PDFPLUMBER TABLE-BASED FALLBACK EXTRACTOR
# ==========================================

def extract_rows_from_tables(pdf_path_or_stream, header):
    """
    Use pdfplumber's built-in table detection as a universal fallback.
    Kicks in when text-based parser returns < 3 rows.
    """
    rows = []
    DATE_RE = re.compile(r'\d{2}[/.\-]\d{2}[/.\-]\d{2,4}|\d{2}[- ][A-Za-z]{3}[- ]\d{4}')
    TIME_RE = re.compile(r'\d{1,2}:\d{2}:\d{2}')
    AMT_RE  = re.compile(r'[\d,]+\.\d{2}')

    try:
        if hasattr(pdf_path_or_stream, 'seek'):
            pdf_path_or_stream.seek(0)
        with pdfplumber.open(pdf_path_or_stream) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if not tables:
                    continue
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    # Identify columns from header row
                    hdr = [str(c or '').lower().strip() for c in table[0]]
                    col = {}
                    for idx, h in enumerate(hdr):
                        if 'date' in h and 'date' not in col:         col['date'] = idx
                        if any(k in h for k in ['time','air']) and 'time' not in col: col['time'] = idx
                        if any(k in h for k in ['prog','programme']) and 'program' not in col: col['program'] = idx
                        if any(k in h for k in ['caption','spot','copy','brand','description']) and 'caption' not in col: col['caption'] = idx
                        if any(k in h for k in ['dur','sec','len']) and 'len' not in col: col['len'] = idx
                        if any(k in h for k in ['rate','amount','value','cost']) and 'rate' not in col: col['rate'] = idx

                    data_start = 1
                    # If no column mapping, auto-detect from first data row
                    if not col:
                        for ri, row in enumerate(table):
                            vals = [str(c or '') for c in row]
                            if any(DATE_RE.search(v) for v in vals):
                                data_start = ri
                                for ci, v in enumerate(vals):
                                    v = v.strip()
                                    if DATE_RE.match(v) and 'date' not in col: col['date'] = ci
                                    elif TIME_RE.match(v) and 'time' not in col: col['time'] = ci
                                    elif AMT_RE.match(v) and 'rate' not in col: col['rate'] = ci
                                    elif v.isdigit() and int(v) < 300 and 'len' not in col: col['len'] = ci
                                break

                    for row in table[data_start:]:
                        if not row: continue
                        vals = [str(c or '').strip() for c in row]
                        row_str = ' '.join(vals)
                        if re.search(r'(total|grand total|taxable|rupee|sub total|net amount|cgst|sgst|hsn)', row_str, re.IGNORECASE): continue
                        if not DATE_RE.search(row_str): continue

                        def _get(key):
                            idx = col.get(key)
                            return vals[idx] if idx is not None and idx < len(vals) else ''

                        date_val = _get('date')
                        time_val = _get('time')
                        rate_val = _get('rate')
                        dm = DATE_RE.search(row_str)
                        tm = TIME_RE.search(row_str)
                        am = AMT_RE.findall(row_str)
                        if not date_val and dm: date_val = dm.group(0)
                        if not time_val and tm: time_val = tm.group(0)
                        if not rate_val and am: rate_val = am[-1]
                        if not date_val: continue

                        rows.append({
                            'Date': date_val, 'Air Time': time_val,
                            'Program': _get('program'), 'Spot Copy': _get('caption'),
                            'LEN': _get('len'), 'Rate (INR)': rate_val,
                        })
    except Exception as e:
        print(f"  [TABLE-FALLBACK] Error: {e}")
    return rows


# ==========================================
# PROCESS SINGLE PDF
# ==========================================

def process_pdf_stream(pdf_stream, filename="Uploaded PDF"):
    print(f"[*] Processing: {filename}")
    all_rows = []

    with pdfplumber.open(pdf_stream) as pdf:
        full_text = ""
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                full_text += txt + "\n"

    if not full_text.strip():
        print("  [!] No text extracted from PDF")
        return all_rows

    # Determine Format — scan FULL text so nothing is missed
    ft = full_text  # use entire text, not just first 1500 chars
    format_detected = "UNKNOWN"

    if "Zee Entertainment" in ft or "ZEE ENTERTAINMENT" in ft:
        format_detected = "Zee Entertainment"
    elif re.search(r'Star\s+India|STAR\s+INDIA|Star\s+Maa|STAR\s+MAA|Star\s+Vijay|Star\s+Bharat|Star\s+Gold|Star\s+Utsav|Star\s+Suvarna|Asianet', ft):
        format_detected = "Star India"
    elif re.search(r'SUN\s+TV|Sun\s+TV|Gemini\s+TV|GEMINI|Sun\s+Network|SUN\s+NETWORK|Sun\s+Life|Surya\s+TV|KTV', ft):
        format_detected = "Sun TV Network"
    elif re.search(r'Sony\s+Pictures|Culver\s+Max|SONY\s+PICTURES|SONY\s+ENTERTAINMENT|Set\s+India|SPE\s+Networks', ft, re.IGNORECASE):
        format_detected = "Sony Network"
    elif re.search(r'Eenadu|EENADU|ETV\s+Network|ETV\s+Telugu|ETV\s+Cinema', ft, re.IGNORECASE):
        format_detected = "Eenadu / ETV"
    elif "B4U" in ft:
        format_detected = "B4U"
    elif re.search(r'Jaya\s+TV|JAYA\s+TV|J\s+Movies?|Jaya\s+Network|MAVIS\s+SATCOM|Jaya\s+Movies', ft, re.IGNORECASE):
        format_detected = "Jaya Network"
    elif re.search(r'Polimer|POLIMER', ft):
        format_detected = "Polimer Network"
    elif re.search(r'Raj\s+Television|RAJ\s+TELEVISION|Raj\s+TV|RAJ\s+TV\s+NETWORK', ft):
        format_detected = "Raj TV"
    elif re.search(r'Vendhar|VENDHAR|SRN\s+Media|SRN\s+MEDIA', ft):
        format_detected = "Vendhar TV"
    elif re.search(r'Vanitha|VANITHA|Rachana\s+Television', ft):
        format_detected = "Vanitha TV"
    elif re.search(r'Goldmines|GOLDMINES', ft):
        format_detected = "Goldmines"
    elif re.search(r'News\s+Nation|NEWS\s+NATION|IBN\s+Lokmat', ft):
        format_detected = "News Nation"
    elif re.search(r'Enter10|ENTER10|Dangal\s+TV|DANGAL\s+TV|Enterr10', ft, re.IGNORECASE):
        format_detected = "Enter10"
    elif re.search(r'Shemaroo|SHEMAROO', ft):
        format_detected = "Shemaroo"
    elif re.search(r'Matrix\s+Broadcast|MATRIX\s+BROADCAST|Matrix\s+Media|MATRIX\s+MEDIA', ft):
        format_detected = "Matrix"
    elif re.search(r'Vasanth|VASANTH', ft):
        format_detected = "Vasanth"
    else:
        format_detected = "Generic"

    print(f"  [INFO] Format detected: {format_detected}")
    
    if format_detected == "Zee Entertainment":
        header = extract_zee_header(full_text)
        broadcast_rows = extract_zee_rows(full_text)
    elif format_detected == "Star India":
        header = extract_star_header(full_text)
        broadcast_rows = extract_star_rows(full_text)
    elif format_detected == "Sun TV Network":
        header = extract_sun_header(full_text)
        broadcast_rows = extract_sun_rows(full_text)
    elif format_detected == "Sony Network":
        header = extract_sony_header(full_text)
        broadcast_rows = extract_sony_rows(full_text)
    elif format_detected == "Eenadu / ETV":
        header = extract_eenadu_header(full_text)
        broadcast_rows = extract_eenadu_rows(full_text)
    elif format_detected == "B4U":
        header = extract_b4u_header(full_text)
        broadcast_rows = extract_b4u_rows(full_text)
    elif format_detected == "Jaya Network":
        header = extract_jaya_header(full_text)
        broadcast_rows = extract_jaya_rows(full_text)
    elif format_detected == "Polimer Network":
        header = extract_polimer_header(full_text)
        broadcast_rows = extract_polimer_rows(full_text)
    elif format_detected == "Raj TV":
        header = extract_raj_header(full_text)
        broadcast_rows = extract_raj_rows(full_text)
    elif format_detected == "Vendhar TV":
        header = extract_vendhar_header(full_text)
        broadcast_rows = extract_vendhar_rows(full_text)
    elif format_detected == "Vanitha TV":
        header = extract_vanitha_header(full_text)
        broadcast_rows = extract_vanitha_rows(full_text)
    elif format_detected == "Shemaroo":
        header = extract_shemaroo_header(full_text)
        broadcast_rows = extract_shemaroo_rows(full_text)
    elif format_detected == "Vasanth":
        header = extract_vasanth_header(full_text)
        broadcast_rows = extract_vasanth_rows(full_text)
    elif format_detected == "Goldmines":
        header = extract_goldmines_header(full_text)
        broadcast_rows = extract_goldmines_rows(full_text)
    elif format_detected == "Enter10":
        header = extract_enter10_header(full_text)
        broadcast_rows = extract_enter10_rows(full_text)
    elif format_detected == "News Nation":
        header = extract_newsnation_header(full_text)
        broadcast_rows = extract_newsnation_rows(full_text)
    else:
        # Generic parser for Matrix and any unknown formats
        header = extract_generic_header(full_text, format_detected)
        broadcast_rows = extract_generic_rows(full_text)

    # ── FALLBACK: if text-based parser returned very few rows, try pdfplumber table extraction ──
    if len(broadcast_rows) < 3:
        table_rows = extract_rows_from_tables(pdf_stream, header)
        if len(table_rows) > len(broadcast_rows):
            print(f"  [FALLBACK] Table extractor found {len(table_rows)} rows (text got {len(broadcast_rows)})")
            broadcast_rows = table_rows

    print(f"  [OK] Broadcaster:    {header.get('Broadcaster Name', 'N/A')}")
    print(f"  [OK] Advertiser:     {header.get('Advertiser Name', 'N/A')}")
    print(f"  [OK] Agency:         {header.get('Agency Name', 'N/A')}")
    print(f"  [OK] Channel:        {header.get('Channel Name', 'N/A')}")
    print(f"  [OK] Invoice No:     {header.get('Invoice Number', 'N/A')}")
    print(f"  [OK] Invoice Date:   {header.get('Invoice Date', 'N/A')}")
    print(f"  [OK] PO Number:      {header.get('PO Number', 'N/A')}")
    print(f"  [OK] Broadcast Rows: {len(broadcast_rows)}")

    # Combine header + each broadcast row
    for brow in broadcast_rows:
        
        # Calculate new column: (INR Rate * Duration) / 10
        calculated_amount = ""
        try:
            amt = float(str(brow.get("Rate (INR)", "0")).replace(",", ""))
            dur = float(str(brow.get("LEN", "0")).replace(",", ""))
            if dur > 0:
                calc = (amt * dur) / 10
                calculated_amount = f"{calc:,.2f}"
        except ValueError:
            pass

        # If brand is at header level (Zee), use it. If at row level (Star), use row level.
        brand = brow.get("Brand", "")
        if not brand and header.get("Brand"):
            brand = header.get("Brand")

        # Auto-derive Day from Date if missing
        day_val = brow.get("Day", "")
        if not day_val:
            day_val = _derive_day_from_date(brow.get("Date", ""))

        row = {
            "Broadcaster Name": header.get("Broadcaster Name", ""),
            "Agency Name": header.get("Agency Name", ""),
            "Advertiser Name": header.get("Advertiser Name", ""),
            "Channel Name": header.get("Channel Name", ""),
            "Billing Period": header.get("Billing Period", ""),
            "PO Number": header.get("PO Number", ""),
            "Invoice Number": header.get("Invoice Number", ""),
            "Invoice Date": header.get("Invoice Date", ""),
            "TP": brow.get("TP", ""),
            "Program": brow.get("Program", ""),
            "Date": brow.get("Date", ""),
            "Day": day_val,
            "Air Time": brow.get("Air Time", ""),
            "LEN (Duration Sec)": brow.get("LEN", ""),
            "Spot Copy (Caption)": brow.get("Spot Copy", ""),
            "Brand": brand,
            "Rate (INR)": brow.get("Rate (INR)", ""),
            "Calculated Amount (INR)": calculated_amount,
            "Time Range/Sales Unit": brow.get("Time Range/Sales Unit", ""),
        }
        all_rows.append(row)

    return all_rows


# ==========================================
# MAIN - PROCESS ALL PDFs
# ==========================================
def process_all_invoices():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    pdf_files = glob(os.path.join(INPUT_FOLDER, "*.pdf"))

    if not pdf_files:
        print(f"[X] No PDFs found in '{INPUT_FOLDER}/' folder")
        return

    print(f"[*] Found {len(pdf_files)} PDF file(s)\n")
    print("=" * 60)

    all_data = []
    for pdf_path in pdf_files:
        with open(pdf_path, 'rb') as f:
            rows = process_pdf_stream(f, os.path.basename(pdf_path))
        all_data.extend(rows)
        print()

    if not all_data:
        print("[X] No data extracted from any PDF")
        return

    column_order = [
        "Broadcaster Name",
        "Agency Name",
        "Advertiser Name",
        "Channel Name",
        "Billing Period",
        "PO Number",
        "Invoice Number",
        "Invoice Date",
        "Brand",
        "Spot Copy (Caption)",
        "Program",
        "TP",
        "Time Range/Sales Unit",
        "Date",
        "Day",
        "Air Time",
        "LEN (Duration Sec)",
        "Rate (INR)",
        "Calculated Amount (INR)",
    ]

    df = pd.DataFrame(all_data)

    for col in column_order:
        if col not in df.columns:
            df[col] = ""

    df = df[column_order]

    import time
    try:
        df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
        final_file = OUTPUT_FILE
    except PermissionError:
        ts = int(time.time())
        alt_file = OUTPUT_FILE.replace(".xlsx", f"_{ts}.xlsx")
        df.to_excel(alt_file, index=False, engine="openpyxl")
        print(f"[!] Original file locked, saved to: {alt_file}")
        final_file = alt_file

    print("=" * 60)
    print(f"[OK] Total rows extracted: {len(df)}")
    print(f"[OK] Data saved to: {final_file}")
    print("=" * 60)


if __name__ == "__main__":
    process_all_invoices()