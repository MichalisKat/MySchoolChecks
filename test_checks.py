"""
test_checks.py - Δοκιμή φόρτωσης checks στο frozen exe
Τρέξε: dist\test_checks.exe
"""
import sys, os, importlib, traceback

print("=" * 50)
print("CHECK LOADER TEST")
print("=" * 50)
print(f"frozen : {getattr(sys, 'frozen', False)}")
print(f"__file__: {__file__}")
print(f"executable: {sys.executable}")

if getattr(sys, 'frozen', False):
    base = sys._MEIPASS
    print(f"_MEIPASS: {sys._MEIPASS}")
else:
    base = os.path.dirname(os.path.abspath(__file__))

print(f"base: {base}")

# Βεβαιωσου οτι το base ειναι στο path
if base not in sys.path:
    sys.path.insert(0, base)

checks_dir = os.path.join(base, 'checks')
print(f"checks_dir: {checks_dir}")
print(f"exists: {os.path.exists(checks_dir)}")

if os.path.exists(checks_dir):
    contents = os.listdir(checks_dir)
    print(f"contents: {contents}")

    for fname in sorted(contents):
        if fname.endswith('.py') and not fname.startswith('_'):
            mod_name = f'checks.{fname[:-3]}'
            print(f"\n>>> Importing {mod_name}...")
            try:
                mod = importlib.import_module(mod_name)
                title = getattr(mod, 'CHECK_TITLE', 'no title')
                print(f"    OK: {title}")
            except Exception:
                print(f"    ERROR:")
                traceback.print_exc()
else:
    print("ERROR: checks_dir δεν υπάρχει!")
    print(f"\nπεριεχόμενα του base ({base}):")
    try:
        for f in os.listdir(base):
            print(f"  {f}")
    except Exception as e:
        print(f"  Δεν μπόρεσα να διαβάσω: {e}")

print("\n" + "=" * 50)
input("Πάτα Enter για έξοδο...")
