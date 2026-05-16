"""
Orchestrates the HPI pipeline in sequence:
  1. explore.py
  2. validate.py
  3. clean.py                → HPIAdmins_v1.xlsx
  4. build_dashboards_all.py → 4 HTML dashboards

Sets HPI_OUTPUT_DATE env var to today's YYYYMMDD so build scripts
use a fresh dated folder instead of a hardcoded one.

On success, saves the new hash to .hpi_state/last_hash.txt.
Prints a JSON summary to stdout.
"""
import datetime, hashlib, json, os, subprocess, sys

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
STATE_DIR   = os.path.join(BASE_DIR, '.hpi_state')
HASH_FILE   = os.path.join(STATE_DIR, 'last_hash.txt')
XLSX_PATH   = os.path.join(BASE_DIR, 'PowerBIData.xlsx')

TODAY       = datetime.date.today().strftime('%Y%m%d')
OUTPUT_DIR  = os.path.join(BASE_DIR, TODAY)

STAGES = [
    ('Explore',      'explore.py'),
    ('Validate',     'validate.py'),
    ('Clean',        'clean.py'),
    ('Dashboards',   'build_dashboards_all.py'),
    ('Presentation', 'build_ppt_pdf.py'),
]


def run_stage(label, script):
    path = os.path.join(BASE_DIR, script)
    env  = os.environ.copy()
    env['HPI_OUTPUT_DATE'] = TODAY
    result = subprocess.run(
        [sys.executable, path],
        capture_output=True, text=True, env=env, cwd=BASE_DIR
    )
    return result


def save_hash():
    import pandas as pd
    df  = pd.read_excel(XLSX_PATH, sheet_name='HPIAdmins')
    raw = df.to_csv(index=False).encode('utf-8')
    h   = hashlib.md5(raw).hexdigest()
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(HASH_FILE, 'w') as f:
        f.write(h)
    return h


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    results = []

    for label, script in STAGES:
        print(f'[{label}] Running {script}...', flush=True)
        r = run_stage(label, script)
        ok = r.returncode == 0
        results.append({'stage': label, 'script': script, 'ok': ok,
                         'stdout': r.stdout[-500:] if r.stdout else '',
                         'stderr': r.stderr[-500:] if r.stderr else ''})
        if ok:
            print(f'[{label}] OK', flush=True)
        else:
            print(f'[{label}] FAILED', flush=True)
            print(r.stderr or r.stdout, flush=True)
            print(json.dumps({'success': False, 'failed_stage': label,
                               'stages': results}))
            sys.exit(1)

    new_hash = save_hash()
    print(json.dumps({
        'success':    True,
        'output_dir': OUTPUT_DIR,
        'date_label': TODAY,
        'hash':       new_hash,
        'stages':     [{'stage': r['stage'], 'ok': r['ok']} for r in results],
    }))


if __name__ == '__main__':
    main()
