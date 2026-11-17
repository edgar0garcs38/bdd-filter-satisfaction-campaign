import pandas as pd
import re
from pathlib import Path

# ============================================================
# Configuration
# ============================================================
INPUT_FILE = "Compilado Leads-03-03-4_15.xlsx"   # adjust to your file path

# ============================================================
# Helpers
# ============================================================

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
        return "+593" + cleaned, True
    if len(cleaned) == 12 and cleaned.startswith("593"):
        return "+" + cleaned, True
    if len(cleaned) == 10:
        if cleaned[0] == "0":
            return "+593" + cleaned[1:], True
        else:
            return cleaned, False
    return cleaned, False

def split_name_custom(full_name: str):
    """
    New logic:
    - 1 word: First Name only
    - 2 words: First word = First Name, second word = Last Name
    - >2 words: First two words = First Name, remainder = Last Name
    """
    if pd.isna(full_name) or str(full_name).strip() == "":
        return "", ""
    
    parts = str(full_name).strip().split()
    
    if len(parts) == 1:
        return parts[0].title(), ""
    elif len(parts) == 2:
        return parts[0].title(), parts[1].title()
    else:
        return " ".join(parts[:2]).title(), " ".join(parts[2:]).title()

# ============================================================
# Main program
# ============================================================

# 1) Read Excel file directly
try:
    df = pd.read_excel(INPUT_FILE, engine="openpyxl")
    df = df.rename(columns=lambda x: str(x).strip())
except Exception as e:
    raise SystemExit(f"Could not read {INPUT_FILE}: {e}")

# Keep only required columns
required_cols = ["full_name", "phone", "email"]
for c in required_cols:
    if c not in df.columns:
        df[c] = ""
df = df[required_cols]

# ============================================================
# MERCATELY PROCESS
# ============================================================

mercately = df.copy()

# Email handling
mercately["email"] = mercately["email"].astype(str).str.strip().str.lower()
mercately["EmailValidFlag"] = mercately["email"].apply(is_valid_email)

invalid_emails = mercately[~mercately["EmailValidFlag"]]
invalid_emails.to_csv("mercately_invalid_emails.csv", index=False)

mercately = mercately[mercately["EmailValidFlag"]].copy()

# Phone handling
validated = mercately["phone"].apply(lambda x: pd.Series(validate_phone_raw(x)))
validated.columns = ["Phone_normalized", "PhoneValidFlag"]
mercately[["phone", "PhoneValidFlag"]] = validated

invalid_phones = mercately[~mercately["PhoneValidFlag"]].copy()
invalid_phones.to_csv("mercately_invalid_phones.csv", index=False)

mercately = mercately[mercately["PhoneValidFlag"]].copy()
mercately["phone"] = mercately["phone"].astype(str)

# Name splitting
mercately["First Name"], mercately["Last Name"] = zip(*mercately["full_name"].apply(split_name_custom))

# Duplicates: by email or phone
dup_by_email = mercately[mercately.duplicated(subset=["email"], keep=False)]
dup_by_phone = mercately[mercately.duplicated(subset=["phone"], keep=False)]
duplicates = pd.concat([dup_by_email, dup_by_phone]).drop_duplicates()
duplicates.to_csv("mercately_duplicates.csv", index=False)

mercately = mercately.drop_duplicates(subset=["email", "phone"], keep="first")

# Final export
final_cols = ["email", "First Name", "Last Name", "phone"]
mercately[final_cols].to_excel("mercately_cleaned.xlsx", index=False)

# ============================================================
# Final logs
# ============================================================
print("Done.")
print(f"Mercately: exported {len(mercately)} unique valid rows to mercately_cleaned.csv")
print(f"Mercately invalid emails: {len(invalid_emails)} rows saved to mercately_invalid_emails.csv")
print(f"Mercately invalid phones: {len(invalid_phones)} rows saved to mercately_invalid_phones.csv")
print(f"Mercately duplicates exported: {len(duplicates)} rows saved to mercately_duplicates.csv")
