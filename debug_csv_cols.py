"""Διαγνωστικό script — εμφανίζει στήλες col50-col62 από stat2_2/CSV_ αρχείο."""
import sys, os, glob, zipfile, io
import pandas as pd

# Ψάξε αρχείο
folders = []
docs = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks', 'downloads')
if os.path.isdir(docs):
    for d in sorted(os.listdir(docs), reverse=True):
        folders.append(os.path.join(docs, d))
folders.append(os.path.join(os.path.expanduser('~'), 'Downloads'))

path = None
for folder in folders:
    for pat in ('CSV_*.zip', 'stat2_2*.csv', 'stat2_2*.xlsx', 'CSV_*.csv'):
        m = [f for f in glob.glob(os.path.join(folder, pat))
             if not f.endswith(('.tmp', '.crdownload'))]
        if m:
            path = sorted(m)[-1]
            break
    if path:
        break

if not path:
    print("ΔΕΝ ΒΡΕΘΗΚΕ αρχείο CSV_/stat2_2.")
    sys.exit(1)

print(f"Αρχείο: {path}\n")

if path.lower().endswith('.zip'):
    with zipfile.ZipFile(path) as z:
        raw = z.read(z.namelist()[0])
    df = pd.read_csv(io.StringIO(raw.decode('cp1253')), sep=';', dtype=str)
elif path.lower().endswith('.xlsx'):
    df = pd.read_excel(path, dtype=str)
else:
    df = pd.read_csv(path, sep=';', dtype=str, encoding='cp1253')

print(f"Σύνολο στηλών: {len(df.columns)}\n")
print("=== Στήλες col50–col62 (header | 1η τιμή) ===")
for i in range(50, min(63, len(df.columns))):
    hdr = df.columns[i]
    val = df[hdr].dropna().iloc[0] if not df[hdr].dropna().empty else '(κενό)'
    print(f"  col{i:2d}: [{hdr}]  →  {val}")
