"""
setup_credentials.py
====================
Πρώτη ρύθμιση credentials για νέο χρήστη του MySchool Checks.

Εκτελείται μια φορά (ή ξανά όταν χρειαστεί αλλαγή password).
Αποθηκεύει τα credentials στο Windows Credential Manager μέσω keyring.

Χρήση:
    python setup_credentials.py

    ή από το main app αυτόματα, αν δεν βρεθούν credentials.
"""

import sys
import getpass
import smtplib
import socket
import logging
from pathlib import Path

# ── Εισαγωγή encryption module ─────────────────────────────────────────────
try:
    from encryption import (
        store_credential,
        store_all_credentials,
        get_credential,
        credentials_exist,
        get_missing_credentials,
        print_credential_status,
        SENSITIVE_KEYS,
        APP_SERVICE,
    )
except ImportError:
    print("ΣΦΑΛΜΑ: Δεν βρέθηκε το encryption.py. Βεβαιώσου ότι βρίσκεται στον ίδιο φάκελο.")
    sys.exit(1)

logging.basicConfig(level=logging.WARNING)

# ─── Σταθερές για validation ──────────────────────────────────────────────────

MYSCHOOL_TEST_URL = "myschool.edu.gr"   # domain για connectivity check
SMTP_TEST_HOST    = "mail.sch.gr"
SMTP_TEST_PORT    = 587
SMTP_TIMEOUT_SEC  = 10


# ─── UI helpers ──────────────────────────────────────────────────────────────

def banner(title: str) -> None:
    print(f"\n{'═'*54}")
    print(f"  {title}")
    print(f"{'═'*54}")


def section(title: str) -> None:
    print(f"\n── {title} {'─'*(48 - len(title))}")


def success(msg: str) -> None:
    print(f"  ✓  {msg}")


def warn(msg: str) -> None:
    print(f"  ⚠  {msg}")


def error(msg: str) -> None:
    print(f"  ✗  {msg}")


def ask(prompt: str, default: str = "") -> str:
    """Ζητά εισαγωγή κειμένου. Αν ο χρήστης πατήσει Enter, επιστρέφει default."""
    suffix = f" [{default}]" if default else ""
    value = input(f"  → {prompt}{suffix}: ").strip()
    return value if value else default


def ask_password(prompt: str) -> str:
    """Ζητά password χωρίς να εμφανίζεται στην οθόνη."""
    while True:
        pwd = getpass.getpass(f"  → {prompt}: ")
        if pwd:
            return pwd
        warn("Το password δεν μπορεί να είναι κενό. Δοκίμασε ξανά.")


# ─── Validation ──────────────────────────────────────────────────────────────

def test_smtp_connection(host: str, port: int, user: str, password: str) -> bool:
    """
    Ελέγχει αν η σύνδεση SMTP λειτουργεί με τα δοθέντα credentials.

    Returns:
        True αν η σύνδεση και το login πέτυχαν, False διαφορετικά.
    """
    try:
        print(f"  Σύνδεση στο {host}:{port} ...", end="", flush=True)
        with smtplib.SMTP(host, port, timeout=SMTP_TIMEOUT_SEC) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(user, password)
        print(" OK")
        return True
    except smtplib.SMTPAuthenticationError:
        print(" ΑΠΟΤΥΧΙΑ")
        error("Λάθος username ή password για email.")
        return False
    except (smtplib.SMTPException, socket.timeout, OSError) as e:
        print(f" ΣΦΑΛΜΑ: {e}")
        warn("Δεν ήταν δυνατή η σύνδεση στον mail server. Ελέγξτε το δίκτυο.")
        return False


def test_internet_connectivity(host: str = "8.8.8.8", port: int = 53) -> bool:
    """Γρήγορος έλεγχος ότι υπάρχει σύνδεση στο internet."""
    try:
        socket.setdefaulttimeout(3)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except OSError:
        return False


# ─── Κύριες ρουτίνες setup ───────────────────────────────────────────────────

def setup_myschool_credentials() -> dict:
    """
    Συλλέγει τα credentials για το MySchool από τον χρήστη.

    Returns:
        dict με MYSCHOOL_USER και MYSCHOOL_PASS
    """
    section("MySchool Credentials")
    print("  Τα credentials για την πύλη e-myschool.gr")
    print("  (ΑΜ εκπαιδευτικού + κωδικός πρόσβασης)\n")

    # Έλεγχος αν υπάρχουν ήδη — προτείνουμε αλλαγή μόνο αν χρειάζεται
    existing_user = get_credential("MYSCHOOL_USER")
    if existing_user:
        warn(f"Βρέθηκε αποθηκευμένο username: {existing_user}")
        overwrite = ask("Θες να το αντικαταστήσεις; (ναι/όχι)", default="όχι")
        if overwrite.lower() not in ("ναι", "nai", "yes", "y"):
            return {}  # Ο χρήστης δεν θέλει αλλαγή

    username = ask("ΑΜ εκπαιδευτικού (username MySchool)")
    password = ask_password("Κωδικός MySchool")

    return {
        "MYSCHOOL_USER": username,
        "MYSCHOOL_PASS": password,
    }


def setup_email_credentials(smtp_host: str = SMTP_TEST_HOST) -> dict:
    """
    Συλλέγει τα credentials για αποστολή email.

    Returns:
        dict με FROM_PASSWORD (και προαιρετικά σχόλιο)
    """
    section("Email / SMTP Credentials")
    print(f"  Χρησιμοποιείται ο mail server: {smtp_host}")
    print("  (email + κωδικός sch.gr λογαριασμού)\n")

    existing = get_credential("FROM_PASSWORD")
    if existing:
        warn("Βρέθηκε αποθηκευμένο email password.")
        overwrite = ask("Θες να το αντικαταστήσεις; (ναι/όχι)", default="όχι")
        if overwrite.lower() not in ("ναι", "nai", "yes", "y"):
            return {}

    password = ask_password("Κωδικός email (FROM_PASSWORD)")

    return {
        "FROM_PASSWORD": password,
    }


def run_setup(validate: bool = True) -> bool:
    """
    Εκτελεί ολόκληρη τη διαδικασία πρώτης ρύθμισης.

    Args:
        validate: Αν True, κάνει test σύνδεσης μετά την αποθήκευση.

    Returns:
        True αν η ρύθμιση ολοκληρώθηκε επιτυχώς.
    """
    banner("MySchool Checks — Ρύθμιση Credentials")

    print("""
  Τα credentials σου θα αποθηκευτούν με ασφάλεια στο
  Windows Credential Manager (κρυπτογραφημένα με DPAPI).
  ΔΕΝ θα αποθηκευτούν σε κανένα αρχείο κειμένου.
    """)

    # ── Έλεγχος internet ────────────────────────────────────────────────────
    if not test_internet_connectivity():
        warn("Δεν φαίνεται να υπάρχει σύνδεση στο internet.")
        warn("Τα credentials θα αποθηκευτούν χωρίς validation.")
        validate = False

    # ── Συλλογή credentials ──────────────────────────────────────────────────
    all_new_credentials = {}

    myschool_creds = setup_myschool_credentials()
    all_new_credentials.update(myschool_creds)

    # Χρειαζόμαστε το FROM_EMAIL για validation SMTP
    section("Email αποστολής")
    from_email = ask("Email αποστολής (FROM_EMAIL)", default="itdipea@sch.gr")
    store_credential("FROM_EMAIL_DISPLAY", from_email)  # non-sensitive, αλλά βολεύει

    email_creds = setup_email_credentials()
    all_new_credentials.update(email_creds)

    # ── Αποθήκευση ──────────────────────────────────────────────────────────
    if all_new_credentials:
        section("Αποθήκευση")
        for key, value in all_new_credentials.items():
            store_credential(key, value)
            success(f"{key} αποθηκεύτηκε στο Credential Manager")
    else:
        warn("Δεν εισήχθησαν νέα credentials.")

    # ── Validation ──────────────────────────────────────────────────────────
    if validate and "FROM_PASSWORD" in all_new_credentials:
        section("Έλεγχος σύνδεσης Email")
        smtp_ok = test_smtp_connection(
            host=SMTP_TEST_HOST,
            port=SMTP_TEST_PORT,
            user=from_email,
            password=all_new_credentials["FROM_PASSWORD"],
        )
        if smtp_ok:
            success("Η σύνδεση email επαληθεύτηκε επιτυχώς.")
        else:
            warn("Το email password ενδέχεται να είναι λάθος.")
            warn("Μπορείς να τρέξεις ξανά το setup: python setup_credentials.py")

    # ── Τελική κατάσταση ────────────────────────────────────────────────────
    section("Ολοκλήρωση")
    print_credential_status()

    missing = get_missing_credentials()
    if missing:
        warn(f"Τα παρακάτω credentials λείπουν ακόμα: {missing}")
        return False

    success("Όλα τα credentials είναι αποθηκευμένα. Μπορείς να ξεκινήσεις το MySchool Checks!")
    return True


def run_migration(json_path: str) -> None:
    """
    Βοηθητική λειτουργία: Μεταφέρει credentials από παλιό JSON στο keyring.
    Χρήσιμο για χρήστες που αναβαθμίζουν από παλιά έκδοση.

    Args:
        json_path: Path προς το local_settings.json
    """
    from encryption import migrate_from_json

    banner("Migration από JSON → Credential Manager")
    print(f"\n  Αρχείο προέλευσης: {json_path}\n")

    confirm = ask("Θες να συνεχίσεις τη μεταφορά; (ναι/όχι)", default="ναι")
    if confirm.lower() not in ("ναι", "nai", "yes", "y"):
        print("  Ακύρωση.")
        return

    ok = migrate_from_json(json_path)
    if ok:
        success("Migration ολοκληρώθηκε.")
        success("Το local_settings.json καθαρίστηκε από passwords.")
        success("Ένα backup αποθηκεύτηκε ως local_settings.json.bak")
    else:
        error("Η migration απέτυχε. Δες τα logs για λεπτομέρειες.")


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="MySchool Checks — Ρύθμιση ασφαλών credentials"
    )
    parser.add_argument(
        "--migrate",
        metavar="JSON_PATH",
        help="Μεταφορά credentials από παλιό local_settings.json στο Credential Manager",
    )
    parser.add_argument(
        "--status",
        action="store_true",
        help="Εμφάνιση κατάστασης αποθηκευμένων credentials",
    )
    parser.add_argument(
        "--no-validate",
        action="store_true",
        help="Παράλειψη ελέγχου σύνδεσης μετά την αποθήκευση",
    )
    args = parser.parse_args()

    if args.status:
        print_credential_status()
    elif args.migrate:
        run_migration(args.migrate)
    else:
        success_flag = run_setup(validate=not args.no_validate)
        sys.exit(0 if success_flag else 1)
