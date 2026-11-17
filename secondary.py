import pandas as pd
import re
from pathlib import Path

# ============================================================
# Configuration
# ============================================================
INPUT_FILE = "Compilado Leads-03-03-4_15.xlsx"   # adjust to your file path

# ============================================================
# Helpers (unchanged except export_chunks removed)
# ============================================================

def clean_poliza(value) -> str:
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
    candidates = ["Phone", "PHONE", "Teléfono", "TELÉFONO", "TELEFONO", "TEL", "CELULAR", "CEL", "phone"]
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

# 1) Read Excel file directly
try:
    merged_df = pd.read_excel(INPUT_FILE, engine="openpyxl")
    merged_df = merged_df.rename(columns=lambda x: str(x).strip())
except Exception as e:
    raise SystemExit(f"Could not read {INPUT_FILE}: {e}")

# Normalize column names
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

# Drop duplicates
duplicates_mailchimp = mailchimp[mailchimp.duplicated(subset=["Email Address"], keep=False)]
duplicates_mailchimp.to_csv("mailchimp_duplicates.csv", index=False)
mailchimp = mailchimp.drop_duplicates(subset=["Email Address"], keep="first")

# Export Mailchimp as a single CSV (no chunking)
mailchimp.to_csv("mailchimp_cleaned.csv", index=False)

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
print(f"Mailchimp: exported {len(mailchimp)} unique valid rows to mailchimp_cleaned.csv")
print(f"Mailchimp invalid emails: {len(invalid_emails_mailchimp)} rows saved to mailchimp_invalid_emails.csv")
print(f"Mailchimp duplicates exported: {len(duplicates_mailchimp)} rows saved to mailchimp_duplicates.csv")
print(f"Mercately: exported {len(mercately)} unique valid rows to mercately_cleaned.csv")
print(f"Mercately invalid emails: {len(invalid_emails_mercately)} rows saved to mercately_invalid_emails.csv")
print(f"Mercately invalid phones: {len(invalid_phones)} rows saved to mercately_invalid_phones.csv")
print(f"Mercately duplicates exported: {len(duplicates_mercately)} rows saved to mercately_duplicates.csv")
