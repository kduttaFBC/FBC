"""
Run this from your backend directory: python3 repro_script.py
Uses your REAL uploaded servicer file and SOUP file to trace excel_rows
through the actual pipeline, no mocking.
"""
import sys
import file_readers, checks, corrections

SERVICER_XLSX_PATH = sys.argv[1]  # path to a real servicer file
SOUP_XLSX_PATH = sys.argv[2]      # path to the real SOUP file
FILENAME = sys.argv[3] if len(sys.argv) > 3 else None  # e.g. "SoupKey||Display__original.xlsx"

with open(SERVICER_XLSX_PATH, "rb") as f:
    servicer_bytes = f.read()
with open(SOUP_XLSX_PATH, "rb") as f:
    soup_bytes = f.read()

soup_df = file_readers.read_soup_file(soup_bytes)
soup_key, display_label, loan_terms_df, servicer_inputs_df, deal = file_readers.read_servicer_file(
    servicer_bytes, filename=FILENAME
)

print("servicer_inputs_df columns:", list(servicer_inputs_df.columns))
print("has _excel_row column:", "_excel_row" in servicer_inputs_df.columns)
if "_excel_row" in servicer_inputs_df.columns:
    print(servicer_inputs_df[["SBA GP Number", "_excel_row", "_sheet_name", "_header_excel_row"]].head(10))

soup_servicer_df = file_readers.filter_soup_by_servicer(soup_df, soup_key)
stat01 = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name=display_label)
print("\ncorrection_data (first 3):")
for c in stat01["correction_data"][:3]:
    print(c)

eligible = corrections.build_eligible_corrections(stat01["correction_data"])
print("\neligible (first 3):")
for e in eligible[:3]:
    print(e)
