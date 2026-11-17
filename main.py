#Objetive of the program: to create a program to filter a data base for two process 1. Mailchimp, 2. Mercately,  and export the results to a csv file. 
#Logic General: 1. Read the data base 2. Get "Asegurado" and split the names in first name and last name
#Logic Mailchimp: 3. Add email 4. Add Company column and fill with "KS" 5. Add "Publico" column and fill with "Cliente" 6. Re order the columns into "Email Address",	"First Name",	"Last Name",	"Compañía",	"Público" 7. Drop duplicates by "Email Address" 8. Export to csv the rows with duplicates emails 9. Split the data base into 1000 rows per file and export to csv.
#Logic Mercately: 3. Add email 4. Add "Phone" column and validate that every one has 10 digits, replace the first charater (verify if is "0") with "593", the cases where the phone number has less that 10 digits, does not have number or start with a character different to "0" export it to a csv file 5. Add "Tags" column and fill with "FormSatisfaccionCliente" 6. Drop duplicates by "Email" and "Phone" 7. Export to csv the rows with duplicates emails or phone numbers 8. Export to csv.
import pandas as pd
import re
import math
import glob
from pathlib import Path

# ============================================================
# Configuration
# ============================================================
INPUT_GLOB = "reports/*.xlsx"   # adjust to your folder
MAILCHIMP_CHUNK_SIZE = 1000

# ============================================================
# Helpers
# ============================================================

def clean_poliza(value) -> str:
    """Normalize contract numbers (remove spaces, dashes, leading zeros)."""
    if pd.isna(value):
        return ""
    s = str(value).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s.lstrip("0")

def is_valid_email(email: str) -> bool:
    """Validate email addresses, including multi-level domains, and reject empty values."""
    if not email or pd.isna(email):
        return False
    email = str(email).strip()
    if email == "":
        return False
    pattern = (
        r'^(?!.*\.\.)'                                # no consecutive dots
        r'[a-zA-Z0-9._%+-]+'                          # local part
        r'@'
        r'(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+'  # domain labels
        r'[a-zA-Z]{2,}$'                              # final TLD
    )
    return re.match(pattern, email) is not None

def split_name(full_name):
    """
    Robust split for ASEGURADO:
    - If the value contains a comma, assume "Last, First" and split on first comma.
    - Else split on first whitespace: first token -> First Name, remainder -> Last Name.
    Returns (first_name, last_name).
    """
    s = "" if pd.isna(full_name) else str(full_name).strip()
    if s == "":
        return "", ""
    # If comma present, treat as "Last, First"
    if "," in s:
        parts = [p.strip() for p in s.split(",", 1)]
        last = parts[0].upper()
        first = parts[1].upper() if len(parts) > 1 else ""
        return first, last
    # Else split on first space: first token -> First Name, remainder -> Last Name
    parts = re.sub(r"\s+", " ", s).split(" ", 1)
    if len(parts) == 1:
        return parts[0].upper(), ""
    return parts[0].upper(), parts[1].upper()

def lower_then_title_series(s: pd.Series) -> pd.Series:
    """Lowercase then Title Case each word; handles NaN safely."""
    return s.fillna("").astype(str).str.lower().str.title()

def lower_email_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.lower()

def export_chunks(df: pd.DataFrame, filename_prefix: str, chunk_size: int = 1000):
    """Split dataframe into chunks and export to CSV files with numeric suffix."""
    if df.empty:
        # create an empty file to indicate no data
        Path(f"{filename_prefix}_1.csv").write_text("")
        return
    num_chunks = math.ceil(len(df) / chunk_size)
    for i in range(num_chunks):
        chunk = df.iloc[i*chunk_size:(i+1)*chunk_size]
        chunk.to_csv(f"{filename_prefix}_{i+1}.csv", index=False)

def clean_phone(phone) -> str:
    """Remove everything except digits from the phone number."""
    if pd.isna(phone):
        return ""
    return re.sub(r'[^0-9]', '', str(phone))

def validate_phone_raw(phone_raw: str):
    """
    Validate and normalize phone numbers according to rules:
    - If cleaned has 9 digits and starts with '9' -> prepend '593' (valid)
    - If cleaned has 10 digits and starts with '0' -> replace leading '0' with '593' (valid)
    - If cleaned has 10 digits and does not start with '0' -> invalid (return cleaned, False)
    - Otherwise invalid (return cleaned, False)
    Returns (normalized_phone_or_cleaned, is_valid_bool)
    """
    cleaned = clean_phone(phone_raw)
    if cleaned == "":
        return "", False
    if len(cleaned) == 9 and cleaned[0] == "9":
        return "593" + cleaned, True
    if len(cleaned) == 10:
        if cleaned[0] == "0":
            return "593" + cleaned[1:], True
        else:
            # 10 digits but doesn't start with 0 -> invalid per spec
            return cleaned, False
    # anything else invalid
    return cleaned, False

def find_phone_column(df: pd.DataFrame):
    """Return the first likely phone column name from common variants, or None."""
    candidates = ["Phone", "PHONE", "Teléfono", "TELÉFONO", "TELEFONO", "TEL", "CELULAR", "CEL"]
    for c in candidates:
        if c in df.columns:
            return c
    # fallback: any column with 'phone' in name (case-insensitive)
    for c in df.columns:
        if "phone" in c.lower() or "tel" in c.lower() or "cel" in c.lower():
            return c
    return None

# ============================================================
# Main program
# ============================================================

# 1) Read all Excel files matching glob
excel_files = glob.glob(INPUT_GLOB)
if not excel_files:
    raise SystemExit(f"No Excel files found with pattern: {INPUT_GLOB}")

dfs = []
for file in excel_files:
    try:
        df = pd.read_excel(file, engine="openpyxl")
        df = df.rename(columns=lambda x: str(x).strip())
        dfs.append(df)
    except Exception as e:
        print(f"Warning: could not read {file}: {e}")

if not dfs:
    raise SystemExit("No readable Excel files found.")

merged_df = pd.concat(dfs, ignore_index=True, sort=False)

# Normalize column names to consistent keys (uppercase trimmed)
merged_df.columns = [str(c).strip() for c in merged_df.columns]

# 2) Split Asegurado into First Name / Last Name
# Accept multiple possible column names for Asegurado
asegurado_col = None
for c in ["ASEGURADO", "Asegurado", "Asegurado "]:
    if c in merged_df.columns:
        asegurado_col = c
        break

if asegurado_col is None:
    # create empty name columns if not present
    merged_df["First Name"] = ""
    merged_df["Last Name"] = ""
else:
    names = merged_df[asegurado_col].apply(lambda x: pd.Series(split_name(x)))
    names.columns = ["First Name", "Last Name"]
    merged_df[["First Name", "Last Name"]] = names

# Normalize only the name columns to Title Case (first letter of each word)
merged_df["First Name"] = lower_then_title_series(merged_df["First Name"])
merged_df["Last Name"] = lower_then_title_series(merged_df["Last Name"])

# ============================================================
# MAILCHIMP PROCESS
# ============================================================

mailchimp = merged_df.copy()

# Find email column (common variants)
email_col = None
for c in ["E-MAIL", "E_MAIL", "EMAIL", "Email", "E-mail", "E mail"]:
    if c in mailchimp.columns:
        email_col = c
        break

if email_col is None:
    mailchimp["Email Address"] = ""
else:
    mailchimp["Email Address"] = mailchimp[email_col].astype(str).str.strip()

# Validate emails (use lowercase for validation consistency)
mailchimp["Email Address"] = lower_email_series(mailchimp["Email Address"])
mailchimp["EmailValidFlag"] = mailchimp["Email Address"].apply(is_valid_email)

# Export invalid emails
invalid_emails_mailchimp = mailchimp[~mailchimp["EmailValidFlag"]]
invalid_emails_mailchimp.to_csv("mailchimp_invalid_emails.csv", index=False)

# Keep only valid emails for Mailchimp flow
mailchimp = mailchimp[mailchimp["EmailValidFlag"]].copy()

# Add required columns
mailchimp["Compañía"] = "KS"
mailchimp["Público"] = "Cliente"

# Title-case names again to ensure proper casing (already done on merged_df)
mailchimp["First Name"] = lower_then_title_series(mailchimp["First Name"])
mailchimp["Last Name"] = lower_then_title_series(mailchimp["Last Name"])

# Reorder columns for Mailchimp
mailchimp = mailchimp.reindex(columns=["Email Address", "First Name", "Last Name", "Compañía", "Público"])

# Duplicates: export rows that have duplicate emails (keep=False to get all duplicates)
duplicates_mailchimp = mailchimp[mailchimp.duplicated(subset=["Email Address"], keep=False)]
duplicates_mailchimp.to_csv("mailchimp_duplicates.csv", index=False)

# Drop duplicates keeping first occurrence
mailchimp = mailchimp.drop_duplicates(subset=["Email Address"], keep="first")

# Export Mailchimp in chunks of 1000 rows
export_chunks(mailchimp, "mailchimp_cleaned", chunk_size=MAILCHIMP_CHUNK_SIZE)

# ============================================================
# MERCATELY PROCESS
# ============================================================

mercately = merged_df.copy()

# Email handling for Mercately: find same email column
if email_col is None:
    mercately["Email"] = ""
else:
    mercately["Email"] = lower_email_series(mercately[email_col].astype(str).str.strip())

# Validate emails and export invalids
mercately["EmailValidFlag"] = mercately["Email"].apply(is_valid_email)
invalid_emails_mercately = mercately[~mercately["EmailValidFlag"]]
invalid_emails_mercately.to_csv("mercately_invalid_emails.csv", index=False)

# Keep only rows with valid emails for Mercately (spec says validate email first)
mercately = mercately[mercately["EmailValidFlag"]].copy()

# Phone handling: detect phone column
phone_col = find_phone_column(mercately)
if phone_col is None:
    mercately["Phone_raw"] = ""
else:
    mercately["Phone_raw"] = mercately[phone_col]

# Validate and normalize phones
validated = mercately["Phone_raw"].apply(lambda x: pd.Series(validate_phone_raw(x)))
validated.columns = ["Phone_normalized", "PhoneValidFlag"]
mercately[["Phone", "PhoneValidFlag"]] = validated

# Export invalid phones
invalid_phones = mercately[~mercately["PhoneValidFlag"]].copy()
invalid_phones.to_csv("mercately_invalid_phones.csv", index=False)

# Keep only valid phones
mercately = mercately[mercately["PhoneValidFlag"]].copy()
mercately["Phone"] = mercately["Phone"].astype(str)

# Add Tags column
mercately["Tags"] = "FormSatisfaccionCliente"

# Title-case only the name columns
mercately["First Name"] = lower_then_title_series(mercately["First Name"])
mercately["Last Name"] = lower_then_title_series(mercately["Last Name"])

# Duplicates: find rows duplicated by Email OR Phone (export them)
dup_by_email = mercately[mercately.duplicated(subset=["Email"], keep=False)]
dup_by_phone = mercately[mercately.duplicated(subset=["Phone"], keep=False)]
duplicates_mercately = pd.concat([dup_by_email, dup_by_phone]).drop_duplicates()
duplicates_mercately.to_csv("mercately_duplicates.csv", index=False)

# Drop duplicates by both Email and Phone (keep first)
mercately = mercately.drop_duplicates(subset=["Email", "Phone"], keep="first")

# Final Mercately export columns (choose order required by Mercately)
final_mercately_cols = ["Email", "First Name", "Last Name", "Phone", "Tags"]
for c in final_mercately_cols:
    if c not in mercately.columns:
        mercately[c] = ""
mercately[final_mercately_cols].to_csv("mercately_cleaned.csv", index=False)

# ============================================================
# Final logs
# ============================================================
print("Done.")
print(f"Mailchimp: exported {len(mailchimp)} unique valid rows (chunks saved).")
print(f"Mailchimp invalid emails: {len(invalid_emails_mailchimp)} rows saved to mailchimp_invalid_emails.csv")
print(f"Mailchimp duplicates exported: {len(duplicates_mailchimp)} rows saved to mailchimp_duplicates.csv")
print(f"Mercately: exported {len(mercately)} unique valid rows to mercately_cleaned.csv")
print(f"Mercately invalid emails: {len(invalid_emails_mercately)} rows saved to mercately_invalid_emails.csv")
print(f"Mercately invalid phones: {len(invalid_phones)} rows saved to mercately_invalid_phones.csv")
print(f"Mercately duplicates exported: {len(duplicates_mercately)} rows saved to mercately_duplicates.csv")
