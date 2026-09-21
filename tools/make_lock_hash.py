#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
tools/make_lock_hash.py
=======================
Δημιουργεί νέο hash για τον κωδικό κλειδώματος των εργαλείων που είναι
αποκλειστικά για τη ΔΙΠΕ Αν. Θεσ/κης (βλ. _password_gate στο main.py).

Χρήση:
    python tools/make_lock_hash.py

Ζητάει τον νέο κωδικό (δεν φαίνεται στην οθόνη) και τυπώνει τις 3 γραμμές
που αντικαθιστούν τις _LOCK_SALT_HEX / _LOCK_ITERATIONS / _LOCK_HASH_HEX
στο MySchoolChecks/main.py. Ο ίδιος ο κωδικός δεν γράφεται πουθενά.

Προτείνεται κωδικός τουλάχιστον 10 χαρακτήρων με γράμματα ΚΑΙ αριθμούς:
ένας 4ψήφιος κωδικός σπάει με δοκιμές σε λίγα λεπτά, ακόμη και με hash.
"""
import getpass
import hashlib
import os

ITERATIONS = 600_000

pwd = getpass.getpass('Νέος κωδικός: ')
if pwd != getpass.getpass('Επανάληψη:   '):
    raise SystemExit('Οι κωδικοί δεν ταιριάζουν.')
if len(pwd) < 10:
    print('⚠ Προσοχή: κωδικός κάτω από 10 χαρακτήρες είναι εύκολο να σπάσει.')

salt = os.urandom(16)
digest = hashlib.pbkdf2_hmac('sha256', pwd.encode('utf-8'), salt, ITERATIONS)

print('\nΑντικατάστησε στο MySchoolChecks/main.py:\n')
print(f"_LOCK_SALT_HEX   = '{salt.hex()}'")
print(f"_LOCK_ITERATIONS = {ITERATIONS}")
print(f"_LOCK_HASH_HEX   = '{digest.hex()}'")
