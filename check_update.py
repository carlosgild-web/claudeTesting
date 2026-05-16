"""
Check whether the HPIAdmins tab in PowerBIData.xlsx has changed since the last run.

Exit 0 → data has changed (or no prior hash exists)
Exit 1 → data is unchanged

Prints a single JSON line to stdout with the result.
"""
import hashlib, json, os, sys
import pandas as pd

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH  = os.path.join(BASE_DIR, 'PowerBIData.xlsx')
STATE_DIR  = os.path.join(BASE_DIR, '.hpi_state')
HASH_FILE  = os.path.join(STATE_DIR, 'last_hash.txt')


def compute_hash():
    df = pd.read_excel(XLSX_PATH, sheet_name='HPIAdmins')
    raw = df.to_csv(index=False).encode('utf-8')
    return hashlib.md5(raw).hexdigest(), df


def read_stored_hash():
    if not os.path.exists(HASH_FILE):
        return None
    with open(HASH_FILE, 'r') as f:
        return f.read().strip()


def main():
    if not os.path.exists(XLSX_PATH):
        print(json.dumps({'changed': False, 'error': 'PowerBIData.xlsx not found'}))
        sys.exit(1)

    current_hash, df = compute_hash()
    stored_hash = read_stored_hash()

    df_valid = df[df.columns].copy()
    df_valid['Date'] = pd.to_datetime(df_valid['Date'], errors='coerce')
    valid_rows = df_valid[df_valid['SalesVolume'].fillna(0) > 0] if 'SalesVolume' in df_valid else df_valid
    latest_date = valid_rows['Date'].max().strftime('%B %Y') if not valid_rows.empty else 'unknown'

    if current_hash != stored_hash:
        result = {
            'changed':     True,
            'hash':        current_hash,
            'rows':        len(df),
            'latest_date': latest_date,
            'first_run':   stored_hash is None,
        }
        print(json.dumps(result))
        sys.exit(0)
    else:
        print(json.dumps({'changed': False, 'hash': current_hash}))
        sys.exit(1)


if __name__ == '__main__':
    main()
