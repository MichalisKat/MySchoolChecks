# Security & Credential Management — MySchool Checks

> This document is intended for Microsoft Store submission reviewers, security auditors, and developers contributing to the project.  
> Repository: [https://github.com/MichalisKat/myschool-checks](https://github.com/MichalisKat/myschool-checks)

---

## Overview

MySchool Checks authenticates to the Greek Ministry of Education's e-myschool.gr portal and sends automated email notifications via SMTP. This requires storing user credentials securely on the end user's machine.

**As of v2.0, the application stores NO credentials in plain text.** All sensitive data is protected using the operating system's native secure credential store.

---

## Credential Storage Architecture

### What is stored — and where

| Credential | Storage Location | Plain text? |
|---|---|---|
| MySchool username (ΑΜ) | Windows Credential Manager | ✗ Encrypted |
| MySchool password | Windows Credential Manager | ✗ Encrypted |
| Email (SMTP) password | Windows Credential Manager | ✗ Encrypted |
| FROM_EMAIL (display) | `local_settings.json` | ✓ Non-sensitive |
| FROM_NAME | `local_settings.json` | ✓ Non-sensitive |
| SMTP_HOST / SMTP_PORT | `local_settings.json` | ✓ Non-sensitive |

### Windows Credential Manager (DPAPI)

On Windows, the application uses the [keyring](https://pypi.org/project/keyring/) library, which stores credentials in the **Windows Credential Manager** via the **Data Protection API (DPAPI)**.

- Credentials are encrypted with a key derived from the user's Windows login credentials.
- No third-party encryption library is used — this is native OS functionality.
- Credentials are accessible only to the user account that stored them.
- Credentials survive application reinstallation (they live in the OS, not the app folder).

**Registry path (Windows):** `Control Panel\User Accounts\Credential Manager` → Windows Credentials → entry named `MySchoolChecksPlus`

### The `local_settings.json` file

The configuration file at `%LOCALAPPDATA%\MySchoolChecksPlus\data\local_settings.json` contains **only non-sensitive settings** such as display name, email address, and SMTP server hostname. No passwords or authentication tokens are stored here.

**Example of a clean (post-migration) `local_settings.json`:**
```json
{
  "FROM_NAME": "Σχολείο ΔΠΕ",
  "FROM_EMAIL": "itdipea@sch.gr",
  "SMTP_HOST": "mail.sch.gr",
  "SMTP_PORT": 587
}
```

---

## First-Time Setup Flow

When a user installs the application for the first time:

1. The application checks Windows Credential Manager for stored credentials via `encryption.credentials_exist()`.
2. If credentials are missing, the user is prompted to enter them via `setup_credentials.py` (or the built-in setup wizard).
3. Credentials are stored in Windows Credential Manager via `keyring.set_password()`.
4. The application never writes passwords to disk at any point in this flow.

For users **upgrading from an older version** that stored credentials in `local_settings.json`, the `setup_credentials.py --migrate` command performs a one-time migration:
- Reads credentials from the old JSON file.
- Stores them in Windows Credential Manager.
- Removes the sensitive fields from the JSON file.
- Creates a `local_settings.json.bak` backup.

---

## Microsoft Store Compliance

This application complies with Microsoft Store policies regarding credential and data handling:

| Policy requirement | How we comply |
|---|---|
| No plain-text passwords in files or registry | Credentials stored exclusively in Windows Credential Manager (DPAPI) |
| Use of OS-native security APIs | `keyring` library delegates to Windows Credential Manager, not a custom crypto implementation |
| User data isolation | DPAPI encrypts per-user; no other user or process can read the credentials |
| Transparent data handling | This document and in-app messages explain exactly what is stored and where |
| No credential exfiltration | Credentials are read only for MySchool portal login and SMTP auth; no telemetry or remote transmission |

---

## Relevant Source Files

| File | Purpose |
|---|---|
| [`encryption.py`](encryption.py) | Core credential store/retrieve/delete/migrate functions |
| [`setup_credentials.py`](setup_credentials.py) | First-time setup wizard; also handles migration from legacy JSON |
| [`main_credentials_integration.py`](main_credentials_integration.py) | Integration guide showing how `main.py` loads credentials from keyring |

---

## Dependencies

```
keyring>=24.0.0
```

The `keyring` library is a well-maintained, MIT-licensed Python package with 5M+ weekly downloads. On Windows it uses `pywin32` under the hood to interface with the native Credential Manager API.

**PyPI:** https://pypi.org/project/keyring/  
**Source:** https://github.com/jaraco/keyring

---

## Threat Model

| Threat | Mitigation |
|---|---|
| Attacker reads `local_settings.json` | File contains no credentials — only public configuration values |
| Attacker reads app source / decompiles executable | Source contains no credentials — credentials are read at runtime from OS keychain |
| Attacker with physical access to the machine | Same access as any app using Windows Credential Manager; mitigated by OS-level user login security |
| Credential leak via logs | `encryption.py` uses Python `logging` module; credential values are never logged |

---

## Contact

For security disclosures or questions, contact the maintainer via the [GitHub repository](https://github.com/MichalisKat/myschool-checks/issues).
