"""
encryption.py
=============
Κεντρικό module διαχείρισης credentials για το MySchool Checks.

Χρησιμοποιεί το keyring library που προσπελαύνει το Windows Credential Manager
(στα Windows) ή το αντίστοιχο secure store σε macOS/Linux.

Τα credentials ΔΕΝ αποθηκεύονται σε αρχεία — αποθηκεύονται απευθείας στο
λειτουργικό σύστημα με κρυπτογράφηση.

Microsoft Store compliance:
- Δεν χρησιμοποιείται κρυπτογράφηση 3rd party
- Χρησιμοποιείται η native κρυπτογράφηση του Windows (DPAPI μέσω Credential Manager)
- Δεν υπάρχουν plain-text passwords σε κανένα αρχείο
"""

import keyring
import json
import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# ─── Σταθερές ───────────────────────────────────────────────────────────────

# Το "service name" που εμφανίζεται στο Windows Credential Manager
APP_SERVICE = "MySchoolChecks"

# Τα κλειδιά που θεωρούνται "ευαίσθητα" και αποθηκεύονται ΜΟΝΟ στο keyring
SENSITIVE_KEYS = {
    "MYSCHOOL_USER",   # username για το MySchool
    "MYSCHOOL_PASS",   # password για το MySchool
    "FROM_PASSWORD",   # password για SMTP / email αποστολή
}

# Τα κλειδιά που θεωρούνται "μη ευαίσθητα" και μένουν στο local_settings.json
NON_SENSITIVE_KEYS = {
    "FROM_NAME",
    "FROM_EMAIL",
    "SMTP_HOST",
    "SMTP_PORT",
    # ... προσθέτεις ό,τι άλλο υπάρχει στο JSON σου
}


# ─── Αποθήκευση ─────────────────────────────────────────────────────────────

def store_credential(key: str, value: str) -> None:
    """
    Αποθηκεύει ένα credential στο Windows Credential Manager.

    Args:
        key:   Το όνομα του credential (π.χ. "MYSCHOOL_USER")
        value: Η τιμή που θα αποθηκευτεί κρυπτογραφημένη
    """
    keyring.set_password(APP_SERVICE, key, value)
    logger.info(f"Credential '{key}' αποθηκεύτηκε στο Credential Manager.")


def store_all_credentials(credentials: dict) -> None:
    """
    Αποθηκεύει πολλαπλά credentials ταυτόχρονα.

    Args:
        credentials: dict με key→value ζεύγη ευαίσθητων δεδομένων
    """
    for key, value in credentials.items():
        if key in SENSITIVE_KEYS:
            store_credential(key, value)
        else:
            logger.warning(
                f"Το κλειδί '{key}' δεν είναι στη λίστα SENSITIVE_KEYS — "
                f"παραλείφθηκε από keyring."
            )


# ─── Ανάκτηση ────────────────────────────────────────────────────────────────

def get_credential(key: str) -> str | None:
    """
    Ανακτά ένα credential από το Windows Credential Manager.

    Args:
        key: Το όνομα του credential (π.χ. "MYSCHOOL_PASS")

    Returns:
        Η τιμή ως string, ή None αν δεν βρεθεί.
    """
    value = keyring.get_password(APP_SERVICE, key)
    if value is None:
        logger.warning(f"Credential '{key}' δεν βρέθηκε στο Credential Manager.")
    return value


def get_all_credentials() -> dict:
    """
    Ανακτά όλα τα ευαίσθητα credentials από το keyring.

    Returns:
        dict με τα credentials που βρέθηκαν (missing keys → None)
    """
    return {key: get_credential(key) for key in SENSITIVE_KEYS}


# ─── Διαγραφή ────────────────────────────────────────────────────────────────

def delete_credential(key: str) -> bool:
    """
    Διαγράφει ένα credential από το Windows Credential Manager.

    Args:
        key: Το όνομα του credential προς διαγραφή

    Returns:
        True αν διαγράφηκε, False αν δεν υπήρχε.
    """
    try:
        keyring.delete_password(APP_SERVICE, key)
        logger.info(f"Credential '{key}' διαγράφηκε.")
        return True
    except keyring.errors.PasswordDeleteError:
        logger.warning(f"Credential '{key}' δεν βρέθηκε — δεν έγινε διαγραφή.")
        return False


def delete_all_credentials() -> None:
    """Διαγράφει ΟΛΑ τα ευαίσθητα credentials από το Credential Manager."""
    for key in SENSITIVE_KEYS:
        delete_credential(key)
    logger.info("Όλα τα credentials διαγράφηκαν από το Credential Manager.")


# ─── Έλεγχος ─────────────────────────────────────────────────────────────────

def credentials_exist() -> bool:
    """
    Ελέγχει αν ΟΛΑ τα απαραίτητα credentials υπάρχουν στο keyring.

    Returns:
        True αν όλα τα SENSITIVE_KEYS έχουν τιμή, False διαφορετικά.
    """
    for key in SENSITIVE_KEYS:
        if get_credential(key) is None:
            return False
    return True


def get_missing_credentials() -> list[str]:
    """
    Επιστρέφει λίστα με τα credentials που λείπουν από το keyring.

    Returns:
        List με ονόματα κλειδιών που δεν βρέθηκαν.
    """
    return [key for key in SENSITIVE_KEYS if get_credential(key) is None]


# ─── Migration από JSON ───────────────────────────────────────────────────────

def migrate_from_json(json_path: str | Path) -> bool:
    """
    Μεταφέρει credentials από παλιό plain-text JSON αρχείο στο keyring,
    και αφαιρεί τα ευαίσθητα πεδία από το JSON.

    ΧΡΗΣΗ: Εκτελείται ΜΟΝΟ μια φορά κατά την αναβάθμιση από παλιά έκδοση.

    Args:
        json_path: Path προς το local_settings.json

    Returns:
        True αν η migration έγινε επιτυχώς, False αν υπήρξε σφάλμα.
    """
    json_path = Path(json_path)

    if not json_path.exists():
        logger.error(f"Αρχείο δεν βρέθηκε: {json_path}")
        return False

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            settings = json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        logger.error(f"Σφάλμα ανάγνωσης JSON: {e}")
        return False

    # Αποθήκευσε τα ευαίσθητα στο keyring
    migrated = []
    for key in SENSITIVE_KEYS:
        if key in settings and settings[key]:
            store_credential(key, str(settings[key]))
            migrated.append(key)

    if not migrated:
        logger.info("Δεν βρέθηκαν ευαίσθητα credentials στο JSON για migration.")
        return True

    # Αφαίρεσε τα ευαίσθητα από το JSON και αποθήκευσε
    cleaned_settings = {k: v for k, v in settings.items() if k not in SENSITIVE_KEYS}

    # Backup του παλιού αρχείου πριν το τροποποιήσεις
    backup_path = json_path.with_suffix(".json.bak")
    json_path.rename(backup_path)
    logger.info(f"Backup αρχείο δημιουργήθηκε: {backup_path}")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(cleaned_settings, f, indent=2, ensure_ascii=False)

    logger.info(
        f"Migration ολοκληρώθηκε. Μεταφέρθηκαν: {migrated}. "
        f"Το JSON καθαρίστηκε από ευαίσθητα δεδομένα."
    )
    return True


# ─── Βοηθητικές ──────────────────────────────────────────────────────────────

def print_credential_status() -> None:
    """Εκτυπώνει την κατάσταση των credentials (για debugging/setup)."""
    print(f"\n{'='*50}")
    print(f"  Κατάσταση Credentials — {APP_SERVICE}")
    print(f"{'='*50}")
    for key in SENSITIVE_KEYS:
        val = keyring.get_password(APP_SERVICE, key)
        status = "✓ Αποθηκευμένο" if val else "✗ Λείπει"
        print(f"  {key:<20} {status}")
    print(f"{'='*50}\n")
