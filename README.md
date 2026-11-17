# Customer Database Cleaner for Marketing Campaigns

Prepares an insurance customer database for a satisfaction campaign, producing ready-to-import files for **Mailchimp** (email) and **Mercately** (WhatsApp).

## What it does
1. Reads and merges every Excel export matching a pattern
2. Splits the insured's full name into first name / last name (handles `"Last, First"` and multi-word names)
3. Validates emails (including multi-level domains) and rejects empties
4. Normalizes Ecuadorian mobile numbers to international format (`593…`): keeps only digits and accepts `09XXXXXXXX` or `9XXXXXXXX`
5. Detects duplicates by email or phone
6. Exports:
   - Mailchimp CSVs, split into 1,000-row chunks to respect import limits
   - A cleaned Mercately Excel file
   - Separate reports of invalid emails, invalid phones and duplicates for manual review

## Scripts
| File | Purpose |
|---|---|
| `main.py` | Full pipeline: Mailchimp + Mercately outputs |
| `secondary.py` | Same cleaning for a single leads file, without chunked export |
| `onlymercately.py` | Mercately-only (WhatsApp) cleaning |

## Tech stack
Python · pandas · regex

## Run it
```bash
pip install -r requirements.txt
python main.py
```

> **Note:** The input files contain confidential client and company data, so they are not included in this repository. Only the source code is published.
