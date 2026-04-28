#!/usr/bin/env python3
"""
gen_odigos.py
═════════════
Παράγει το MySchoolChecks_Odigos.pdf — Οδηγός Χρήστη

Χρήση:
    python gen_odigos.py

Απαιτήσεις:
    pip install reportlab
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, HRFlowable, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# ── Fonts ──────────────────────────────────────────────────────────────────────
_FDIR = r'C:\Windows\Fonts'
pdfmetrics.registerFont(TTFont('Arial',        os.path.join(_FDIR, 'arial.ttf')))
pdfmetrics.registerFont(TTFont('Arial-Bold',   os.path.join(_FDIR, 'arialbd.ttf')))
pdfmetrics.registerFont(TTFont('Arial-Italic', os.path.join(_FDIR, 'ariali.ttf')))

# ── Σταθερές ──────────────────────────────────────────────────────────────────
VERSION = '0.9.8'
AUTHOR  = 'Μιχάλης Κατσιρντάκης'
EMAIL   = 'itdipea@sch.gr'
TEL     = '2310 954145'
ORG     = 'Διεύθυνση Π.Ε. Ανατολικής Θεσσαλονίκης'
OUTPUT  = 'MySchoolChecks_Odigos.pdf'

# ── Χρώματα ───────────────────────────────────────────────────────────────────
HDR    = colors.HexColor('#1A237E')
ACCENT = colors.HexColor('#3949AB')
LIGHT  = colors.HexColor('#E8EAF6')
TIP    = colors.HexColor('#E8F5E9')
WARN   = colors.HexColor('#FFF3E0')
TEXT   = colors.HexColor('#212121')
DESC   = colors.HexColor('#757575')
BORDER = colors.HexColor('#C5CAE9')


# ── Styles ─────────────────────────────────────────────────────────────────────
def _s(**kw):
    kw.setdefault('fontName', 'Arial')
    kw.setdefault('textColor', TEXT)
    return ParagraphStyle('_', **kw)

S = {
    'h1':    _s(fontName='Arial-Bold', fontSize=18, leading=24, spaceAfter=10, textColor=HDR),
    'h2':    _s(fontName='Arial-Bold', fontSize=12, leading=17, spaceBefore=12, spaceAfter=5, textColor=HDR),
    'body':  _s(fontSize=10, leading=15, spaceAfter=6, alignment=TA_JUSTIFY),
    'note':  _s(fontName='Arial-Italic', fontSize=9, leading=13, textColor=DESC),
    'ctr':   _s(fontSize=10, leading=14, alignment=TA_CENTER),
    'toc':   _s(fontSize=10, leading=17),
    'toc2':  _s(fontSize=9.5, leading=15, leftIndent=16, textColor=DESC),
    'tip':   _s(fontSize=9.5, leading=14, textColor=colors.HexColor('#1B5E20')),
    'warn':  _s(fontSize=9.5, leading=14, textColor=colors.HexColor('#E65100')),
    'hdr_w': _s(fontName='Arial-Bold', fontSize=9, textColor=colors.white, leading=12),
    'cell':  _s(fontSize=9, leading=12),
    'small_hdr': _s(fontName='Arial-Bold', fontSize=8.5, textColor=HDR, leading=11),
    'step_n': _s(fontName='Arial-Bold', fontSize=14, textColor=colors.white,
                  alignment=TA_CENTER, leading=18),
    'step_t': _s(fontName='Arial-Bold', fontSize=10, leading=14),
    'step_d': _s(fontSize=10, leading=14),
}


# ── Header / Footer ────────────────────────────────────────────────────────────
def _hf(canvas, doc):
    canvas.saveState()
    if doc.page == 1:
        canvas.restoreState()
        return
    w, h = A4
    # Header
    canvas.setFont('Arial-Bold', 9); canvas.setFillColor(HDR)
    canvas.drawString(2*cm, h - 1.5*cm, 'MySchool Checks  |  Οδηγός Χρήστη')
    canvas.setFont('Arial', 8); canvas.setFillColor(DESC)
    canvas.drawString(2*cm, h - 1.9*cm, f'{ORG}  ·  {AUTHOR}  ·  {EMAIL}')
    canvas.setFillColor(HDR)
    canvas.rect(2*cm, h - 2.2*cm, w - 4*cm, 0.5, fill=1, stroke=0)
    # Footer
    canvas.setFillColor(HDR)
    canvas.rect(2*cm, 2*cm, w - 4*cm, 0.5, fill=1, stroke=0)
    canvas.setFont('Arial', 8); canvas.setFillColor(DESC)
    canvas.drawCentredString(w / 2, 1.5*cm, f'Σελίδα {doc.page - 1}')
    canvas.restoreState()


# ── Βοηθητικές ────────────────────────────────────────────────────────────────
def _ftable(rows, col1=4.5*cm, col2=10.5*cm):
    """Πίνακας label: value."""
    data = [
        [Paragraph(f'<b>{lbl}</b>', ParagraphStyle('_l', fontName='Arial-Bold',
          fontSize=9.5, leading=13, textColor=HDR)),
         Paragraph(val, ParagraphStyle('_v', fontName='Arial', fontSize=9.5, leading=13))]
        for lbl, val in rows
    ]
    t = Table(data, colWidths=[col1, col2])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), LIGHT),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('BOX', (0,0), (-1,-1), 0.5, BORDER),
        ('LINEBELOW', (0,0), (-1,-2), 0.3, BORDER),
    ]))
    return t


def _steps(steps):
    """Αριθμημένα βήματα."""
    rows = []
    for num, title, desc in steps:
        rows.append([
            Paragraph(str(num), S['step_n']),
            [Paragraph(title, S['step_t']), Paragraph(desc, S['step_d'])],
        ])
    t = Table(rows, colWidths=[1.2*cm, 13.8*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), ACCENT),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (1,0), (1,-1), 10),
        ('ROWBACKGROUNDS', (1,0), (1,-1), [colors.white, LIGHT]),
        ('BOX', (0,0), (-1,-1), 0.5, BORDER),
        ('LINEBELOW', (0,0), (-1,-2), 0.3, BORDER),
    ]))
    return t


def _tip(text, warn=False):
    bg = WARN if warn else TIP
    st = S['warn'] if warn else S['tip']
    pfx = '⚠  ' if warn else '💡  '
    t = Table([[Paragraph(pfx + text, st)]], colWidths=[15*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg),
        ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('BOX', (0,0), (-1,-1), 0.5, BORDER),
    ]))
    return t


def _sp(n=0.3): return Spacer(1, n*cm)
def _hr(): return HRFlowable(width='100%', color=ACCENT, thickness=1, spaceAfter=10)


# ── Κατασκευή PDF ─────────────────────────────────────────────────────────────
def build():
    doc = SimpleDocTemplate(OUTPUT, pagesize=A4,
                             rightMargin=2*cm, leftMargin=2*cm,
                             topMargin=2.8*cm, bottomMargin=2.5*cm)
    story = []

    # ── Εξώφυλλο ──────────────────────────────────────────────────────────────
    story += [
        _sp(3),
        Paragraph('MySchool Checks', ParagraphStyle('_ct', fontName='Arial-Bold',
            fontSize=32, leading=40, textColor=HDR, alignment=TA_CENTER)),
        _sp(0.3),
        Paragraph('Οδηγός Χρήστη', ParagraphStyle('_cs', fontName='Arial',
            fontSize=18, leading=24, textColor=ACCENT, alignment=TA_CENTER)),
        _sp(0.4),
        HRFlowable(width='60%', color=ACCENT, thickness=2, spaceAfter=14),
        Paragraph('Αυτοματοποιημένοι Έλεγχοι Δεδομένων MySchool',
            ParagraphStyle('_cd', fontName='Arial', fontSize=13, leading=18,
                           textColor=TEXT, alignment=TA_CENTER)),
        _sp(0.3),
        Paragraph(ORG, ParagraphStyle('_co', fontName='Arial-Bold', fontSize=11,
            leading=16, textColor=HDR, alignment=TA_CENTER)),
        _sp(4),
    ]

    info = Table([
        [Paragraph(f'<b>{AUTHOR}</b>', S['ctr'])],
        [Paragraph(f'{EMAIL}  ·  {TEL}', S['ctr'])],
        [Paragraph(f'Απρίλιος 2026  ·  v{VERSION}', ParagraphStyle('_cv',
            fontName='Arial', fontSize=9, textColor=DESC, alignment=TA_CENTER))],
    ], colWidths=[11*cm])
    info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), LIGHT),
        ('BOX', (0,0), (-1,-1), 1, ACCENT),
        ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ]))
    story.append(Table([[info]], colWidths=[15*cm],
                        style=[('ALIGN',(0,0),(-1,-1),'CENTER')]))
    story.append(PageBreak())

    # ── Πίνακας Περιεχομένων ──────────────────────────────────────────────────
    story += [Paragraph('Πίνακας Περιεχομένων', S['h1']), _hr()]
    toc = [
        ('1. Τι είναι το MySchool Checks', []),
        ('2. Εγκατάσταση', ['2.1 Κατέβασμα και εκτέλεση setup', '2.2 Πρώτη εκκίνηση']),
        ('3. Ρυθμίσεις', ['3.1 Καρτέλα Σύνδεση', '3.2 Καρτέλα Email', '3.3 Καρτέλα Αρχεία']),
        ('4. Αυτόματη Λήψη Στατιστικών', []),
        ('5. Εκπαιδευτικοί ανά Ειδικότητα', []),
        ('6. Στοιχεία Σχολικών Μονάδων', []),
        ('7. Εκτέλεση Ελέγχων', ['7.1 Επιλογή αποστολής email']),
        ('8. Οδηγός Ελέγχων — Τι κάνει ο καθένας', [
            '8.1 Έλεγχος Διοικητικού Έργου — Ανάλυση αποτελεσμάτων',
            '8.2 Υπόλοιπα Ωραρίου — Κατώφλι και Pivot αναφορά',
        ]),
        ('9. Πού βρίσκω τα αποτελέσματα;', []),
        ('10. Αντιμετώπιση Προβλημάτων', []),
    ]
    for title, subs in toc:
        story.append(Paragraph(title, S['toc']))
        for sub in subs:
            story.append(Paragraph(sub, S['toc2']))
    story.append(PageBreak())

    # ── 1. Τι είναι ───────────────────────────────────────────────────────────
    story += [
        Paragraph('1. Τι είναι το MySchool Checks', S['h1']),
        Paragraph(
            'Το MySchool Checks είναι εφαρμογή για Windows που αυτοματοποιεί ελέγχους δεδομένων '
            'εκπαιδευτικών στο MySchool. Συνδέεται στο Πληροφοριακό Σύστημα, κατεβάζει τα στατιστικά '
            'αρχεία, τα επεξεργάζεται και παράγει αναφορές Excel με δυνατότητα αυτόματης αποστολής email '
            'στα σχολεία.', S['body']),
        Paragraph('<b>Δε χρειάζεται καμία γνώση προγραμματισμού.</b>', S['body']),
        _sp(0.3),
        _ftable([
            ('Έλεγχοι',         '8 αυτοματοποιημένοι έλεγχοι δεδομένων MySchool'),
            ('Ειδικότητες',     'Εξαγωγή λίστας εκπαιδευτικών ανά ειδικότητα σε Excel'),
            ('Σχολικές Μονάδες','Εξαγωγή στοιχείων σχολείων ανά Δήμο σε Excel'),
            ('Στατιστικά',      'Αυτόματο κατέβασμα αρχείων από το MySchool'),
            ('Αποτελέσματα',    'Αρχεία Excel στον φάκελο Έγγραφα → MySchoolChecks'),
            ('Email',           'Αποστολή ανά σχολείο ή σε test mode'),
            ('Browser',         'Chrome ή Firefox — επιλογή από τις Ρυθμίσεις'),
        ]),
        _sp(0.5),
    ]

    # ── 2. Εγκατάσταση ────────────────────────────────────────────────────────
    _sec2 = [
        Paragraph('2. Εγκατάσταση', S['h1']),
        Paragraph('2.1 Κατέβασμα και εκτέλεση setup', S['h2']),
        _steps([
            ('1', 'Κατεβάστε το αρχείο εγκατάστασης',
             'Πηγαίνετε στη διεύθυνση github.com/MichalisKat/myschool-checks → Releases → '
             f'κατεβάστε την τελευταία έκδοση (myschool-checks-{VERSION}-setup.exe)'),
            ('2', 'Τρέξτε το setup',
             f'Διπλό κλικ στο myschool-checks-{VERSION}-setup.exe. Αν εμφανιστεί προειδοποίηση '
             'Windows (SmartScreen), κάντε κλικ "Περισσότερες πληροφορίες" → "Εκτέλεση ούτως ή άλλως".'),
            ('3', 'Ολοκλήρωση εγκατάστασης',
             'Πατήστε Next → Install → Finish. Δημιουργείται αυτόματα συντόμευση στην επιφάνεια εργασίας.'),
        ]),
        _sp(0.3),
        Paragraph('2.2 Πρώτη εκκίνηση', S['h2']),
        _steps([
            ('1', 'Ανοίξτε την εφαρμογή',
             'Διπλό κλικ στη συντόμευση MySchool Checks στην επιφάνεια εργασίας.'),
            ('2', 'Οθόνη εκκίνησης',
             'Εμφανίζεται μικρό παράθυρο "Εκκίνηση..." — περιμένετε λίγα δευτερόλεπτα.'),
            ('3', 'Κύριο παράθυρο',
             'Εδώ επιλέγετε ελέγχους, κατεβάζετε στατιστικά και ξεκινάτε τους ελέγχους.'),
        ]),
        _sp(0.3),
        _tip('Την πρώτη φορά που χρησιμοποιείτε αυτόματη λήψη, η εφαρμογή κατεβάζει αυτόματα '
             'τον κατάλληλο driver για τον browser σας. Απαιτείται σύνδεση στο internet.'),
    ]
    story.append(KeepTogether(_sec2))
    story.append(PageBreak())

    # ── 3. Ρυθμίσεις ──────────────────────────────────────────────────────────
    story += [
        Paragraph('3. Ρυθμίσεις', S['h1']),
        Paragraph('Πατήστε το κουμπί <b>Ρυθμίσεις</b> στο κύριο παράθυρο. Το παράθυρο έχει τρεις καρτέλες:', S['body']),
        Paragraph('3.1 Καρτέλα Σύνδεση', S['h2']),
        _ftable([
            ('Username',         'Ο 7ψήφιος κωδικός σύνδεσης στο MySchool'),
            ('Κωδικός MySchool', 'Το password σύνδεσης στο MySchool'),
            ('Κωδικός email',    'Ο κωδικός του email αποστολής (π.χ. itdipea@sch.gr) — '
                                 '<b>προαιρετικό</b>, συμπληρώστε μόνο αν επιθυμείτε αυτόματη αποστολή email'),
            ('Browser',          'Επιλέξτε Chrome ή Firefox — και οι δύο πρέπει να είναι εγκατεστημένοι'),
        ]),
        _sp(0.2),
        _tip('Τα passwords αποθηκεύονται κρυπτογραφημένα στο Windows Credential Manager. '
             'Δεν γράφονται σε κανένα αρχείο κειμένου.'),
        Paragraph('3.2 Καρτέλα Email', S['h2']),
        _ftable([
            ('Εμφανιζόμενο όνομα', 'Όνομα που εμφανίζεται ως αποστολέας'),
            ('Email αποστολής',    'Διεύθυνση από την οποία στέλνονται τα αποτελέσματα'),
            ('SMTP Host',          'Διακομιστής email — συνήθως mail.sch.gr'),
            ('Υπογραφή',           'Κείμενο που εμφανίζεται στο τέλος κάθε email'),
        ]),
        Paragraph('3.3 Καρτέλα Αρχεία', S['h2']),
        _ftable([
            ('Αρχείο Αδυνατούντων',
             'Αρχείο Excel/CSV με αδυνατούντες υπό έγκριση — χρειάζεται μόνο για τον έλεγχο "Υπόλοιπα Ωραρίου"'),
        ]),
        _sp(0.2),
        Paragraph('Πατήστε <b>Αποθήκευση</b> για να αποθηκεύσετε τις ρυθμίσεις.', S['body']),
        PageBreak(),
    ]

    # ── 4. Λήψη Στατιστικών ───────────────────────────────────────────────────
    story += [
        Paragraph('4. Αυτόματη Λήψη Στατιστικών', S['h1']),
        Paragraph(
            'Πατήστε <b>Λήψη Δεδομένων</b> στο κύριο παράθυρο. Τα checkboxes ξεκινούν '
            'απενεργοποιημένα — επιλέξτε τα αρχεία που θέλετε ή πατήστε <b>«Όλα»</b> για να '
            'τα επιλέξετε μαζί. Ο browser ανοίγει αυτόματα, συνδέεται στο MySchool και κατεβάζει '
            'τα επιλεγμένα αρχεία.', S['body']),
        _sp(0.2),
        _ftable([
            ('Τοποθετήσεις', 'Τοποθετήσεις'),
            ('2.1',  'Κατάλογος σχολικών μονάδων'),
            ('4.1',  'Οργανικές τοποθετήσεις εκπαιδευτικών'),
            ('4.2',  'Αποσπασμένοι εκπαιδευτικοί'),
            ('4.8',  'Ωράριο εκπαιδευτικών'),
            ('4.9',  'Παρόντες εκπαιδευτικοί'),
            ('4.11', 'Μείωση ωραρίου'),
            ('4.12', 'Συμπλήρωση ωραρίου'),
            ('4.16', 'Αιτιολόγηση απουσίας εκπαιδευτικών'),
            ('4.20', 'Άδειες Άνευ Αποδοχών'),
            ('4.21', 'Άδειες (πλην ΑΑ)'),
            ('8.2',  'Επιβεβαίωση δεδομένων σχολείων'),
            ('Αδυνατούντες', 'Αδυνατούντες ανά ειδικότητα — κατεβαίνει απευθείας'),
            ('2.2',  'Εκτεταμένα στοιχεία σχολικών μονάδων (επικοινωνία, Διευθυντής)'),
            ('3.1',  'Κατανομή μαθητών ανά τάξη, τμήμα και σχολείο'),
        ], col1=3.5*cm, col2=11.5*cm),
        _sp(0.3),
        _tip('Αρχεία που υπάρχουν ήδη από σήμερα παραλείπονται αυτόματα (εμφανίζονται με ✓). '
             'Αν δεν θέλετε αυτόματη λήψη, κάθε έλεγχος επιτρέπει χειροκίνητη επιλογή αρχείου.'),
        Paragraph('<b>Πού αποθηκεύονται:</b> Έγγραφα → MySchoolChecks → downloads → YYYYMMDD', S['body']),
        PageBreak(),
    ]

    # ── 5. Εκπαιδευτικοί ανά Ειδικότητα (ΝΕΟ) ────────────────────────────────
    story += [
        Paragraph('5. Εκπαιδευτικοί ανά Ειδικότητα', S['h1']),
        Paragraph(
            'Πατήστε <b>Εκπ/κοί ανά Ειδικότητα</b> στο κύριο παράθυρο για εξαγωγή λίστας '
            'εκπαιδευτικών συγκεκριμένης ειδικότητας σε αρχείο Excel.', S['body']),
        _sp(0.2),
        _steps([
            ('1', 'Επιλέξτε ειδικότητα',
             'Η λίστα συμπληρώνεται αυτόματα από το αρχείο Τοποθετήσεων.'),
            ('2', 'Επιλέξτε στήλες εξόδου',
             'Επιλέξτε αν θέλετε να περιλαμβάνονται στο Excel: Email ΠΣΔ / Email / Κινητό.'),
            ('3', 'Εξαγωγή ή Αποστολή',
             'Πατήστε «Μόνο Excel» για άμεση εξαγωγή, ή «Δημιουργία & Αποστολή» για '
             'αποστολή email σε σύμβουλο.'),
        ]),
        _sp(0.3),
        _ftable([
            ('Απόντες',      'Εμφανίζονται με κόκκινο χρώμα — περιλαμβάνουν αιτιολόγηση '
                             '& ημερομηνία επιστροφής'),
            ('Ταξινόμηση',   'Αλφαβητική κατά επώνυμο'),
            ('Αποτελέσματα', 'Αρχεία Excel στο φάκελο Έγγραφα → MySchoolChecks → results_YYYYMMDD'),
        ]),
        _sp(0.3),
        _tip('Απαιτούμενα αρχεία: Τοποθετήσεις + Κατάλογος σχολείων (2.1) + stat4_1 + stat4_2 + '
             'stat4_16. Κατεβάστε τα πρώτα από «Λήψη Δεδομένων».'),
        PageBreak(),
    ]

    # ── 6. Στοιχεία Σχολικών Μονάδων ─────────────────────────────────────────
    story += [
        Paragraph('6. Στοιχεία Σχολικών Μονάδων', S['h1']),
        Paragraph(
            'Πατήστε <b>Σχολικές Μονάδες</b> στο κύριο παράθυρο για εξαγωγή '
            'στοιχείων σχολικών μονάδων (Δημοτικά + Νηπιαγωγεία) ανά Δήμο σε αρχείο Excel.', S['body']),
        _sp(0.2),
        _steps([
            ('1', 'Επιλέξτε Δήμο',
             'Η λίστα συμπληρώνεται αυτόματα από το αρχείο 3.1 (κατανομή μαθητών).'),
            ('2', 'Επιλέξτε προβολή',
             '«Ανά Τάξη» — αναλυτικά δεδομένα ανά τάξη με μερικά αθροίσματα (μόνο Δημοτικά) '
             'και γενικό άθροισμα Δήμου.  '
             '«Ανά Μονάδα» — σύνολα τμημάτων, αγοριών, κοριτσιών ανά σχολείο με στοιχεία επικοινωνίας.'),
            ('3', 'Εξαγωγή ή Αποστολή',
             'Πατήστε «Μόνο Excel» για άμεση εξαγωγή, ή «Δημιουργία & Αποστολή» για '
             'αποστολή email σε σχετικό παραλήπτη.'),
        ]),
        _sp(0.3),
        _ftable([
            ('Δεδομένα',       'Τμήματα, Αγόρια, Κορίτσια, Σύνολο μαθητών ανά τάξη'),
            ('Επικοινωνία',    'Τηλέφωνο, Email, Διεύθυνση, Διευθυντής (από αρχείο 2.2)'),
            ('Φίλτρα',         'Εξαιρούνται Ιδιωτικά, Ξένα σχολεία και σχολεία με Αναστολή'),
            ('Ταξινόμηση',     'Πρώτα Νηπιαγωγεία, έπειτα Δημοτικά — αλφαβητικά'),
            ('Αποτελέσματα',   'Αρχεία Excel στο φάκελο Έγγραφα → MySchoolChecks → results_YYYYMMDD'),
        ]),
        _sp(0.3),
        _tip('Απαιτούμενα αρχεία: 3.1 (Κατανομή μαθητών) και 2.2 (Εκτεταμένα Στοιχεία). '
             'Κατεβάστε τα πρώτα από «Λήψη Δεδομένων».'),
        PageBreak(),
    ]

    # ── 7. Εκτέλεση Ελέγχων ───────────────────────────────────────────────────
    story += [
        Paragraph('7. Εκτέλεση Ελέγχων', S['h1']),
        _steps([
            ('1', 'Κατεβάστε τα στατιστικά αρχεία',
             'Πατήστε «Λήψη Δεδομένων» και κατεβάστε τα αρχεία που χρειάζονται οι έλεγχοι. '
             'Αν παραλείψετε αυτό το βήμα, η εφαρμογή σας ενημερώνει ποια αρχεία λείπουν.'),
            ('2', 'Επιλέξτε έναν ή περισσότερους ελέγχους',
             'Κάντε κλικ στα checkboxes δίπλα στους ελέγχους. Το κουμπί "Όλοι" επιλέγει όλους.'),
            ('3', 'Πατήστε Εκκίνηση ελέγχου',
             'Το εφαρμογή εκτελεί τους ελέγχους διαδοχικά.'),
            ('4', 'Επιλέξτε τρόπο αποστολής',
             'Κάθε έλεγχος ρωτά αν θέλετε να στείλετε email ή όχι (δείτε παρακάτω).'),
            ('5', 'Δείτε τα αποτελέσματα',
             'Στο τέλος εμφανίζεται παράθυρο με σύνοψη — πλοηγηθείτε με τα βέλη αριστερά/δεξιά.'),
        ]),
        _sp(0.4),
        Paragraph('7.1 Επιλογή αποστολής email', S['h2']),
        _ftable([
            ('Χωρίς αποστολή',   'Δημιουργεί μόνο το αρχείο Excel — δεν στέλνει τίποτα'),
            ('Test mode',         'Στέλνει ένα email στον δικό σας λογαριασμό για έλεγχο'),
            ('Κανονική αποστολή', 'Ένα αρχείο Excel + email για κάθε σχολείο με ευρήματα'),
        ]),
        _sp(0.3),
        _tip('Ξεκινήστε πάντα με Test mode για να ελέγξετε τα αποτελέσματα πριν κάνετε '
             'κανονική αποστολή στα σχολεία. Η «Κανονική αποστολή» εμφανίζεται μόνο αν έχετε '
             'ορίσει κωδικό email στις Ρυθμίσεις.'),
        PageBreak(),
    ]

    # ── 8. Οδηγός Ελέγχων ─────────────────────────────────────────────────────
    story.append(Paragraph('8. Οδηγός Ελέγχων — Τι κάνει ο καθένας', S['h1']))

    checks = [
        [Paragraph('Έλεγχος', S['hdr_w']), Paragraph('Τι κάνει', S['hdr_w']),
         Paragraph('Αρχεία', S['hdr_w']), Paragraph('Email', S['hdr_w'])],
        [Paragraph('1 · Επιβεβαίωση\nΔεδομένων', S['small_hdr']),
         Paragraph('Σχολεία χωρίς επιβεβαίωση δεδομένων. Η ημερομηνία cutoff υπολογίζεται αυτόματα (1η ή 15η του μήνα).', S['cell']),
         Paragraph('8.2', S['cell']), Paragraph('Ένα ανά σχολείο', S['cell'])],
        [Paragraph('2 · Διαφορές\nAK-AL', S['small_hdr']),
         Paragraph('Εκπαιδευτικοί όπου το υποχρεωτικό ωράριο (AK) διαφέρει από το άθροισμα ωρών (AL).', S['cell']),
         Paragraph('4.9', S['cell']), Paragraph('Test mode', S['cell'])],
        [Paragraph('3 · Αρνητικά\nΥπόλοιπα', S['small_hdr']),
         Paragraph('Εκπαιδευτικοί με αναθέσεις περισσότερες από το διδακτικό τους ωράριο.', S['cell']),
         Paragraph('4.8', S['cell']), Paragraph('Ένα ανά σχολείο', S['cell'])],
        [Paragraph('4 · Άδειες ΑΑ &\nΠαρόντες', S['small_hdr']),
         Paragraph('Εκπαιδευτικοί παρόντες ενώ βρίσκονται σε άδεια άνευ αποδοχών.', S['cell']),
         Paragraph('4.20 + 4.9', S['cell']), Paragraph('Test mode', S['cell'])],
        [Paragraph('5 · Άδειες &\nΠαρόντες', S['small_hdr']),
         Paragraph('Εκπαιδευτικοί παρόντες ενώ βρίσκονται σε μακροχρόνια άδεια.', S['cell']),
         Paragraph('4.21 + 4.9', S['cell']), Paragraph('Ένα ανά σχολείο', S['cell'])],
        [Paragraph('6 · Ελλιπή\nΑνάληψη', S['small_hdr']),
         Paragraph('Εκπαιδευτικοί χωρίς Ημερομηνία Ανάληψης σε ενεργή τοποθέτηση.', S['cell']),
         Paragraph('4.8', S['cell']), Paragraph('Ένα ανά σχολείο', S['cell'])],
        [Paragraph('7 · Διοικητικό\nΈργο', S['small_hdr']),
         Paragraph('Γραμματειακή Υποστήριξη στο 4.12 — σύγκριση με ΠΔΕ απόφαση και Αδυνατούντες.', S['cell']),
         Paragraph('4.12 + Αδυν.', S['cell']), Paragraph('Test mode', S['cell'])],
        [Paragraph('8 · Υπόλοιπα\nΩραρίου', S['small_hdr']),
         Paragraph('Υπόλοιπα ωραρίου εκπαιδευτικών. Παράγει συνολικό αρχείο + pivot αναφορά (5 φύλλα).', S['cell']),
         Paragraph('4.8+4.12\n+4.11+Αδυν.', S['cell']), Paragraph('Ένα ανά σχολείο', S['cell'])],
    ]
    ct = Table(checks, colWidths=[3*cm, 7*cm, 2.5*cm, 2.5*cm])
    ct.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HDR),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT]),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('BOX', (0,0), (-1,-1), 0.5, BORDER),
        ('INNERGRID', (0,0), (-1,-1), 0.3, BORDER),
    ]))
    story.append(ct)
    story.append(PageBreak())
    story += [
        Paragraph('8.1 Έλεγχος Διοικητικού Έργου — Ανάλυση αποτελεσμάτων', S['h2']),
        _ftable([
            ('Φύλλο 1\n«ΠΔΕ vs Αδυνατούντες»',
             'Περιλαμβάνει εκπαιδευτικούς από το 4.12 με μη-μηδενική Γραμματειακή Υποστήριξη '
             'και παρατηρήσεις της μορφής ΠΔΕ αριθμός/ημερομηνία (π.χ. ΠΔΕ 12345/15-03-2026). '
             'Ομαδοποιεί ανά ειδικότητα και συγκρίνει το πλήθος με το αρχείο Αδυνατούντων.'),
            ('Φύλλο 2\n«Χωρίς Απόφαση ΠΔΕ»',
             'Εγγραφές με μη-μηδενική Γραμματειακή Υποστήριξη αλλά χωρίς έγκυρη μορφή ΠΔΕ '
             'στις παρατηρήσεις. Γραμμές με πορτοκαλί: εκπαιδευτικοί με πλήρες διοικητικό '
             '(Γραμματειακή = Ώρες Φορέα).'),
        ], col1=3.5*cm, col2=11.5*cm),
        _sp(0.3),
        _tip('Κρίσιμο: Κατά την καταχώρηση Γραμματειακής Υποστήριξης στο 4.12, το πεδίο '
             'Παρατηρήσεις πρέπει να είναι συμπληρωμένο με τα στοιχεία έγκρισης της ΠΔΕ '
             'στη μορφή: ΠΔΕ αριθμός/ημερομηνία (π.χ. ΠΔΕ 12345/15-03-2026).', warn=True),
        _sp(0.5),
        Paragraph('8.2 Υπόλοιπα Ωραρίου — Κατώφλι και Pivot αναφορά', S['h2']),
        _ftable([
            ('Κατώφλι',
             'Κατά την εκτέλεση ορίζεις τον ελάχιστο αριθμό ωρών υπολοίπου για να εμφανιστεί '
             'εκπαιδευτικός (προεπιλογή: 8 ώρες). Εκπαιδευτικοί με υπόλοιπο κάτω από αυτό '
             'αγνοούνται. Αγνοούνται επίσης όσοι βρίσκονται στο αρχείο Αδυνατούντων υπό έγκριση.'),
        ], col1=3.5*cm, col2=11.5*cm),
        _sp(0.2),
        Paragraph('Η <b>Pivot αναφορά</b> περιέχει 5 φύλλα:', S['body']),
        _ftable([
            ('Φύλλο 1\nΑνά Ειδικότητα',     'Πλήθος εκπαιδευτικών και σύνολο ωρών υπολοίπου ανά ειδικότητα.'),
            ('Φύλλο 2\nΑνά Σχολείο',         'Πίνακας σχολείων × ειδικοτήτων — ώρες υπολοίπου ανά συνδυασμό.'),
            ('Φύλλο 3\nΔιευθυντές',          'Μόνο εκπαιδευτικοί σε θέση Διευθυντή Σχολείου.'),
            ('Φύλλο 4\nΥποδιευθυντές',       'Μόνο εκπαιδευτικοί σε θέση Υποδιευθυντή Σχολείου.'),
            ('Φύλλο 5\nΑναπληρωτές',         'Μόνο αναπληρωτές.'),
        ], col1=3.5*cm, col2=11.5*cm),
    ]
    story.append(PageBreak())

    # ── 9. Αποτελέσματα ───────────────────────────────────────────────────────
    story += [
        Paragraph('9. Πού βρίσκω τα αποτελέσματα;', S['h1']),
        Paragraph('Όλα τα αρχεία αποθηκεύονται στον φάκελο <b>Έγγραφα → MySchoolChecks</b>:', S['body']),
        _ftable([
            ('Αποτελέσματα\nελέγχων',
             'Έγγραφα → MySchoolChecks → results_YYYYMMDD → [όνομα ελέγχου]'),
            ('Εκπ/κοί ανά\nειδικότητα',
             'Έγγραφα → MySchoolChecks → results_YYYYMMDD → Εκπαιδευτικοί_{ειδικ}_{ημερ}.xlsx'),
            ('Σχολικές\nΜονάδες',
             'Έγγραφα → MySchoolChecks → results_YYYYMMDD → ΣχολικέςΜονάδες_{δήμος}_{ημερ}.xlsx'),
            ('Στατιστικά\nαρχεία',
             'Έγγραφα → MySchoolChecks → downloads → YYYYMMDD'),
        ]),
        _sp(0.3),
        Paragraph(
            'Το YYYYMMDD είναι η ημερομηνία (π.χ. 20260426 = 26 Απριλίου 2026). '
            'Τα αρχεία κάθε μέρας είναι σε ξεχωριστό φάκελο.', S['body']),
        _tip('Γρήγορη πρόσβαση: Win+R → %USERPROFILE%\\Documents\\MySchoolChecks → Enter'),
        PageBreak(),
    ]

    # ── 10. Αντιμετώπιση Προβλημάτων ──────────────────────────────────────────
    story += [
        Paragraph('10. Αντιμετώπιση Προβλημάτων', S['h1']),
        _ftable([
            ('Δεν ανοίγει η\nεφαρμογή',
             'Απεγκαταστήστε και επανεγκαταστήστε την εφαρμογή από το setup'),
            ('Browser δεν ανοίγει',
             'Βεβαιωθείτε ότι ο Chrome/Firefox είναι ενημερωμένος και τα στοιχεία σύνδεσης είναι σωστά'),
            ('Λάθος στοιχεία\nσύνδεσης',
             'Ρυθμίσεις → Σύνδεση → ξαναεισάγετε username + κωδικό MySchool'),
            ('Σφάλμα αποστολής\nemail',
             'Ελέγξτε τον κωδικό email στις Ρυθμίσεις. Δοκιμάστε πρώτα Test Mode'),
            ('Κενά αποτελέσματα',
             'Βεβαιωθείτε ότι κατεβάσατε τα αρχεία για σήμερα πριν τρέξετε τον έλεγχο'),
            ('Δεν βρίσκω\nτα αρχεία',
             'Έγγραφα → MySchoolChecks (ή Win+R → %USERPROFILE%\\Documents\\MySchoolChecks)'),
            ('Αδυνατεί η λήψη\nΤοποθετήσεων',
             'Βεβαιωθείτε ότι έχετε σύνδεση internet και σωστά credentials MySchool'),
            ('Άλλο πρόβλημα',
             f'Επικοινωνήστε: {AUTHOR} — {TEL} — {EMAIL}'),
        ]),
        _sp(1),
        _hr(),
        Paragraph(f'{ORG}  ·  {AUTHOR}  ·  {TEL}  ·  {EMAIL}',
            ParagraphStyle('_fc', fontName='Arial', fontSize=9,
                           textColor=DESC, alignment=TA_CENTER)),
        PageBreak(),
    ]

    # ── Ευχαριστήριες ─────────────────────────────────────────────────────────
    story += [
        _sp(4),
        Paragraph('Ευχαριστίες', ParagraphStyle('_ht', fontName='Arial-Bold',
            fontSize=22, leading=28, textColor=HDR, alignment=TA_CENTER)),
        _sp(0.4),
        HRFlowable(width='40%', color=ACCENT, thickness=2, spaceAfter=20),
        _sp(1),
        Paragraph(
            'Αφιερώνεται στον συνάδελφο Χρήστο Νιαβή, σε ένδειξη εκτίμησης για τη δημιουργική μας συνεργασία.',
            ParagraphStyle('_ded', fontName='Arial-Italic', fontSize=11, leading=17,
                           textColor=TEXT, alignment=TA_CENTER)),
        _sp(1.5),
        Paragraph(
            'Ευχαριστίες στους κ. Δήμου Παναγιώτη (ΔΙ.Δ.Ε. Αν. Θεσ/κης), '
            'κα Ορφανίδου Μαρία και κα Δρίνη Φωτεινή (ΔΙ.Π.Ε. Κιλκίς) '
            'και κα Τερζή Χριστίνα (ΔΙ.Π.Ε. Δυτ. Θεσ/κης) '
            'για τις πολύτιμες προτάσεις και παρατηρήσεις τους.',
            ParagraphStyle('_ded2', fontName='Arial-Italic', fontSize=11, leading=17,
                           textColor=TEXT, alignment=TA_CENTER)),
    ]

    doc.build(story, onFirstPage=_hf, onLaterPages=_hf)
    print(f'OK: {OUTPUT}  ({os.path.getsize(OUTPUT):,} bytes)')


if __name__ == '__main__':
    build()
