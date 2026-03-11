import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ─── Styles ───────────────────────────────────────────────────────────────────
HEADER_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
SUBHDR_FILL  = PatternFill("solid", fgColor="2E75B6")
SUBHDR_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
HIGH_FILL    = PatternFill("solid", fgColor="FF6B6B")
MED_FILL     = PatternFill("solid", fgColor="FFD166")
LOW_FILL     = PatternFill("solid", fgColor="D0E8C5")
ALT_FILL     = PatternFill("solid", fgColor="EBF3FB")
WRAP         = Alignment(wrap_text=True, vertical="top")
CENTER       = Alignment(horizontal="center", vertical="top")
THIN         = Side(border_style="thin", color="CCCCCC")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def style_header(cell, fill=None):
    cell.font = HEADER_FONT
    cell.fill = fill or HEADER_FILL
    cell.alignment = CENTER
    cell.border = BORDER

def style_cell(cell, fill=None, wrap=True):
    if fill:
        cell.fill = fill
    cell.alignment = WRAP if wrap else Alignment(vertical="top")
    cell.border = BORDER
    cell.font = Font(name="Calibri", size=10)

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 – MISSING CATEGORIES OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Missing Categories"

headers1 = [
    "Category Name",
    "Category Type",
    "Priority",
    "Evident Has Product?",
    "Evident Has Category Page?",
    "AMScope",
    "Leica",
    "Zeiss",
    "Nikon",
    "SEO Opportunity Notes",
    "Suggested URL Slug",
]
ws1.row_dimensions[1].height = 30
for col, h in enumerate(headers1, 1):
    c = ws1.cell(row=1, column=col, value=h)
    style_header(c)

col_widths1 = [35, 20, 10, 20, 22, 10, 10, 10, 10, 55, 40]
for i, w in enumerate(col_widths1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# Data rows: [Category Name, Type, Priority, Has Product, Has Page, AMScope, Leica, Zeiss, Nikon, Notes, Slug]
rows1 = [
    # ── HIGH PRIORITY ─────────────────────────────────────────────────────────
    ("Darkfield Microscopes",
     "Microscope Type",
     "High",
     "Yes",
     "No",
     "Yes", "", "", "",
     "Explicitly named in brief as an example gap. AMScope has both compound & specialty darkfield categories. High-intent search term.",
     "/en/products/darkfield-microscopes/"),

    ("Stereo Microscopes",
     "Microscope Type",
     "High",
     "Yes (SZX/SZ series)",
     "No",
     "Yes", "Yes", "Yes", "Yes",
     "All four competitors have a dedicated stereo microscope category. Evident likely sells stereo models (SZX) but has no /stereo/ path in sitemap.",
     "/en/products/stereo-microscopes/"),

    ("Fluorescence Microscopes",
     "Microscope Type",
     "High",
     "Yes (BX53, IX73, etc.)",
     "No",
     "Yes", "Yes", "Yes", "Yes",
     "Fundamental microscopy category. Evident's upright/inverted systems support fluorescence but there is no dedicated fluorescence category page.",
     "/en/products/fluorescence-microscopes/"),

    ("Phase Contrast Microscopes",
     "Microscope Type",
     "High",
     "Yes (capability on multiple models)",
     "No",
     "Yes", "", "", "",
     "AMScope lists phase contrast as a dedicated special microscope category. Common search term for biology labs.",
     "/en/products/phase-contrast-microscopes/"),

    ("Polarizing Microscopes",
     "Microscope Type",
     "High",
     "Yes (BX53-P listed under /upright/)",
     "No (lives under /upright/)",
     "Yes", "", "Yes (Axio Imager 2 Pol)", "Yes",
     "Nikon and Zeiss both have dedicated polarizing microscope categories. BX53-P is buried under /upright/ — missed SEO opportunity.",
     "/en/products/polarizing-microscopes/"),

    ("Multiphoton Microscopes",
     "Microscope Type",
     "High",
     "Yes (FV4000-MPE, FV5000-MPE)",
     "No (lives under /confocal/)",
     "", "Yes (DIVE)", "Yes", "Yes (AX R MP)",
     "Evident has strong multiphoton products but they are bundled under /confocal/. Competitors treat multiphoton as a standalone category.",
     "/en/products/multiphoton-microscopes/"),

    ("Super-Resolution Microscopes",
     "Microscope Type",
     "High",
     "Yes (mentioned on homepage)",
     "No",
     "", "Yes (STELLARIS STED)", "Yes", "Yes (N-STORM, STEDYCON)",
     "Leica, Zeiss, and Nikon all have dedicated super-resolution categories. Evident mentions super-resolution but has no /super-resolution/ category page.",
     "/en/products/super-resolution-microscopes/"),

    ("Slide Scanners / Whole Slide Imaging",
     "Imaging System",
     "High",
     "Yes (SLIDEVIEW VS200, VS-M1)",
     "No",
     "", "", "Yes (AxioScan 7)", "Yes",
     "Evident's SLIDEVIEW scanners are flagship products but appear to have no dedicated category page in the sitemap. Large opportunity in digital pathology search.",
     "/en/products/slide-scanners/"),

    ("3D Optical Profilometry / Surface Metrology",
     "Imaging System",
     "High",
     "Yes (LEXT OLS5500)",
     "No",
     "", "", "", "",
     "LEXT OLS5500 is highlighted on Evident's homepage but no /surface-metrology/ or /profilometry/ category page exists in sitemap.",
     "/en/products/3d-optical-profilometry/"),

    ("Epi-Fluorescence Microscopes",
     "Microscope Type",
     "High",
     "Yes (multiple systems)",
     "No",
     "Yes", "", "", "",
     "AMScope has a dedicated epi-fluorescence category. Evident's systems support epi-fluorescence but no category page targeting this term.",
     "/en/products/epi-fluorescence-microscopes/"),

    ("Materials Science Microscopes",
     "Application Category",
     "High",
     "Yes (BX53M, GX53 for materials)",
     "No",
     "", "", "Yes (Axiolab/Axio Imager materials)", "",
     "Evident has materials-specific microscopes (BX53M, GX53) but groups them under /upright/ and /inverted/. A materials science category page would capture this segment.",
     "/en/products/materials-science-microscopes/"),

    # ── MEDIUM PRIORITY ───────────────────────────────────────────────────────
    ("TIRF Microscopes",
     "Microscope Technique",
     "Medium",
     "Likely (IX-series supports TIRF)",
     "No",
     "", "", "", "Yes (Ti2-LAPP Photostimulation & TIRF)",
     "Nikon has a dedicated TIRF/photostimulation page. TIRF is a high-value research technique with dedicated search demand.",
     "/en/products/tirf-microscopes/"),

    ("Light Sheet Microscopes",
     "Microscope Type",
     "Medium",
     "No",
     "No",
     "", "Yes (DLS)", "Yes", "Yes",
     "Leica, Zeiss, and Nikon all offer light sheet systems. If Evident enters this space, a category page will be essential. Worth monitoring.",
     "/en/products/light-sheet-microscopes/"),

    ("Spinning Disk Confocal Microscopes",
     "Microscope Type",
     "Medium",
     "Unclear",
     "No",
     "", "", "", "Yes (Yokogawa CSU-W1, Crest X-Light)",
     "Nikon partners with Yokogawa for spinning disk confocal. Evident bundles all confocal under one category — missed opportunity for the spinning disk keyword segment.",
     "/en/products/spinning-disk-confocal-microscopes/"),

    ("Cell Culture Microscopes",
     "Application Category",
     "Medium",
     "Yes (CM30 inverted)",
     "No",
     "", "Yes (Mateo TL/FL)", "", "",
     "Leica's Mateo series specifically targets cell culture labs. Evident's CM30 inverted is ideal for this use case — no category page targets this audience.",
     "/en/products/cell-culture-microscopes/"),

    ("High Content Imaging Systems",
     "Imaging System",
     "Medium",
     "Likely (IXplore platform)",
     "No",
     "", "", "", "Yes",
     "Nikon has high content imaging as a product category. Evident's IXplore platform may support HCI workflows — gap in category-level visibility.",
     "/en/products/high-content-imaging/"),

    ("Veterinary Microscopes",
     "Application Category",
     "Medium",
     "Likely (upright compound systems)",
     "No",
     "Yes", "", "", "",
     "AMScope has a dedicated veterinary microscope category. Distinct search intent from human clinical — Evident's CX43/CX23 could target this segment.",
     "/en/products/veterinary-microscopes/"),

    ("Blood Analysis Microscopes",
     "Application Category",
     "Medium",
     "Yes (upright compound systems)",
     "No",
     "Yes", "", "", "",
     "AMScope lists blood analysis microscopes as a compound microscope subcategory. High-intent clinical market.",
     "/en/products/blood-analysis-microscopes/"),

    ("Industrial Inspection Microscopes",
     "Application Category",
     "Medium",
     "Yes (DSX2000, BX53M)",
     "No",
     "Yes", "", "", "",
     "AMScope has dedicated industrial inspection category under stereo microscopes. Evident's DSX and materials microscopes serve this market but no category page exists.",
     "/en/products/industrial-inspection-microscopes/"),

    ("Metallurgical Microscopes",
     "Microscope Type",
     "Medium",
     "Yes (BX53M, GX53)",
     "No",
     "Yes", "", "", "",
     "AMScope has a dedicated metallurgical microscope category. BX53M is explicitly for metallurgy — a dedicated page would capture this search segment.",
     "/en/products/metallurgical-microscopes/"),

    ("Forensics Microscopes",
     "Application Category",
     "Medium",
     "Likely (stereo + compound systems)",
     "No",
     "Yes (fiber/hair/forensics)", "", "", "",
     "AMScope has a forensics/fiber analysis application category. Distinct search intent worth targeting.",
     "/en/products/forensics-microscopes/"),

    ("Clinical / Medical Microscopes",
     "Application Category",
     "Medium",
     "Yes (CX series, upright systems)",
     "No (exists as section on homepage only)",
     "", "", "", "",
     "Evident's homepage highlights Clinical Diagnostics but no dedicated /clinical-microscopes/ or /medical-microscopes/ category page exists in sitemap.",
     "/en/products/clinical-microscopes/"),

    ("Automated Microscopes",
     "Microscope Feature",
     "Medium",
     "Yes (BX63, IX83 motorized)",
     "No",
     "", "Yes", "", "",
     "Leica has automated microscopes as a distinct category under light microscopes. Evident has motorized systems (BX63) but groups them under /upright/.",
     "/en/products/automated-microscopes/"),

    # ── CAMERA CATEGORIES ─────────────────────────────────────────────────────
    ("sCMOS Cameras",
     "Camera Type",
     "High",
     "Yes (multiple models)",
     "No",
     "", "Yes (K8, K3)", "Yes (Axiocam 820/807)", "",
     "Leica and Zeiss both have sCMOS as a camera specification category. Evident sells sCMOS cameras but no sCMOS-specific category page exists.",
     "/en/products/scmos-cameras/"),

    ("High-Speed Microscope Cameras",
     "Camera Type",
     "Medium",
     "Yes (likely)",
     "No",
     "Yes", "", "", "",
     "AMScope lists high-speed cameras as a digital microscope subcategory. Research demand exists for fast frame-rate imaging.",
     "/en/products/high-speed-cameras/"),

    ("Low-Light / Cooled Cameras",
     "Camera Type",
     "Medium",
     "Yes (EM-CCD under /cameras-and-accessories/)",
     "No (EM-CCD not surfaced as a category)",
     "Yes", "Yes", "", "",
     "AMScope and Leica both surface low-light/cooled cameras. Evident's EM-CCD lives under /cameras-and-accessories/ — not a findable category.",
     "/en/products/cooled-cameras/"),

    ("Fluorescence Microscope Cameras",
     "Camera Type",
     "Medium",
     "Yes (monochrome cameras)",
     "No",
     "", "Yes (dedicated fluorescence camera category)", "Yes (Axiocam 820/807 mono listed under fluorescence)", "",
     "Leica and Zeiss explicitly categorize cameras for fluorescence separately. Evident does not surface this segment.",
     "/en/products/fluorescence-cameras/"),

    ("USB Microscope Cameras",
     "Camera Type",
     "Low",
     "Possibly",
     "No",
     "Yes", "", "", "",
     "AMScope has a USB camera category. Lower relevance for Evident's positioning but captures entry-level search.",
     "/en/products/usb-cameras/"),

    # ── LOWER PRIORITY / AWARENESS ────────────────────────────────────────────
    ("FLIM Microscopes (Fluorescence Lifetime Imaging)",
     "Microscope Technique",
     "Low",
     "Unclear",
     "No",
     "", "Yes (STELLARIS FALCON)", "", "",
     "Leica's FALCON system enables FLIM. Niche but high-value technique. Worth monitoring if Evident develops capability.",
     "/en/products/flim-microscopes/"),

    ("Coherent Raman Scattering (CRS) / Label-Free Microscopes",
     "Microscope Technique",
     "Low",
     "No",
     "No",
     "", "Yes (STELLARIS CRS)", "", "",
     "Leica uniquely offers CRS/label-free imaging. Highly specialized — out of scope for now but worth watching.",
     "/en/products/label-free-microscopes/"),

    ("Dental Lab Microscopes",
     "Application Category",
     "Low",
     "Possibly (stereo)",
     "No",
     "Yes (stereo dental lab)", "Yes (surgical dental)", "", "",
     "Both AMScope (stereo) and Leica (surgical) have dental categories. Niche but addressable with stereo microscopes.",
     "/en/products/dental-microscopes/"),

    ("Gemology / Jewelry Microscopes",
     "Application Category",
     "Low",
     "Likely (stereo systems)",
     "No",
     "Yes", "", "", "",
     "AMScope has gemology/jewelry as both a stereo and application category. Out of core Evident focus but stereo products could target this.",
     "/en/products/gemology-microscopes/"),

    ("Mineralogy Microscopes",
     "Application Category",
     "Low",
     "Possibly (polarizing/stereo)",
     "No",
     "Yes", "", "", "",
     "AMScope has mineralogy under stereo microscopes. Niche geology market.",
     "/en/products/mineralogy-microscopes/"),

    ("Educational / Student Microscopes",
     "Application Category",
     "Low",
     "Yes (CX23 entry-level)",
     "No",
     "Yes", "", "Yes (Primo series)", "",
     "Zeiss Primo and AMScope have dedicated student/education categories. Evident's CX23 is entry-level — capturing education search traffic is possible.",
     "/en/products/student-microscopes/"),
]

priority_fills = {"High": HIGH_FILL, "Medium": MED_FILL, "Low": LOW_FILL}

for row_idx, row in enumerate(rows1, 2):
    fill = priority_fills.get(row[2], None)
    alt = ALT_FILL if row_idx % 2 == 0 else None
    for col_idx, value in enumerate(row, 1):
        c = ws1.cell(row=row_idx, column=col_idx, value=value)
        if col_idx == 3:  # Priority column gets color
            style_cell(c, fill=fill)
        else:
            style_cell(c, fill=alt)

ws1.freeze_panes = "A2"
ws1.auto_filter.ref = f"A1:{get_column_letter(len(headers1))}1"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 – EVIDENT SCIENTIFIC EXISTING CATEGORIES
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Evident Existing Categories")

headers2 = ["Category Name", "URL Path (from sitemap)", "Example Products", "Notes"]
ws2.row_dimensions[1].height = 30
for col, h in enumerate(headers2, 1):
    c = ws2.cell(row=1, column=col, value=h)
    style_header(c)

col_widths2 = [35, 45, 55, 50]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

evident_cats = [
    ("Upright Microscopes", "/en/products/upright/",
     "BX63, BX53, BX53-P, BX53M, BX46, BX43, BXFM, CX43, CX23",
     "Includes research, routine, polarizing (BX53-P), and materials (BX53M) models"),
    ("Inverted Microscopes", "/en/products/inverted/",
     "IXplore IX73, IXplore IX83, CKX53, CM30, GX53, APEXVIEW",
     "Research & cell culture inverted systems"),
    ("Digital Microscopes", "/en/products/digital/",
     "DSX2000",
     "Industrial digital microscopes — no eyepiece, all-digital imaging"),
    ("Confocal Microscopes", "/en/products/confocal/",
     "FLUOVIEW FV5000, FV4000, FV5000-MPE, FV4000-MPE",
     "Includes both confocal and multiphoton (-MPE suffix) systems"),
    ("Semiconductor Inspection", "/en/products/semiconductor-inspection/",
     "AL120-12",
     "Wafer and semiconductor inspection systems"),
    ("Cleanliness & Particle Analysis", "/en/products/cleanliness-and-particle-analysis/",
     "CIX100",
     "ISO 16232 / VDA 19 cleanliness analysis systems"),
    ("Digital Cameras", "/en/products/digital-cameras/",
     "DP75, SC180, EP50",
     "Color and specialty cameras for microscopy"),
    ("Cameras & Accessories (EM-CCD)", "/en/products/cameras-and-accessories/",
     "EM-CCD",
     "Electron-multiplying CCD cameras — not surfaced as a standalone category"),
    ("Objectives", "/en/products/objectives/",
     "TIRF-HR, BXC series, various UIS2 objectives",
     "General objectives product category"),
    ("IX Objectives (Standard)", "/en/products/ix-objectives-standard/",
     "N5702600, N5702400, etc.",
     "IX-series specific objective category"),
    ("Optics & Microscope Accessories", "/en/products/optics-and-microscope-accessories/",
     "N5229300, N5229400, N1478200",
     "Condensers, filters, adapters, and other accessories"),
    ("Software", "/en/products/software/",
     "cellSens",
     "Imaging and analysis software"),
]

for row_idx, row in enumerate(evident_cats, 2):
    fill = ALT_FILL if row_idx % 2 == 0 else None
    for col_idx, value in enumerate(row, 1):
        c = ws2.cell(row=row_idx, column=col_idx, value=value)
        style_cell(c, fill=fill)

ws2.freeze_panes = "A2"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 – COMPETITOR CATEGORY MATRIX
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Competitor Category Matrix")

headers3 = ["Category", "AMScope", "Leica", "Zeiss", "Nikon", "Source (Sitemap/Page)"]
ws3.row_dimensions[1].height = 30
for col, h in enumerate(headers3, 1):
    c = ws3.cell(row=1, column=col, value=h)
    style_header(c)

col_widths3 = [40, 14, 14, 14, 14, 70]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

YES = "✓"
NO  = "–"

# Rows: [Category, AMScope, Leica, Zeiss, Nikon, Source]
matrix_data = [
    # Microscope Types
    ("── MICROSCOPE TYPES ──", "", "", "", "", ""),
    ("Upright Microscopes",                YES, YES, YES, YES, "All sitemaps confirmed"),
    ("Inverted Microscopes",               YES, YES, YES, YES, "All sitemaps confirmed"),
    ("Stereo / Zoom Microscopes",          YES, YES, YES, YES, "AMScope: /collections/stereo-microscopes; Leica: /products/light-microscopes; Zeiss: light-microscopes page; Nikon: products page"),
    ("Digital Microscopes",                YES, YES, YES, YES, "AMScope: /collections/digital-microscopes; Leica: /products/digital-microscopes; Zeiss: products page; Nikon: products page"),
    ("Confocal Microscopes",               NO,  YES, YES, YES, "Leica: STELLARIS; Zeiss: LSM 910/990; Nikon: AX series"),
    ("Multiphoton Microscopes",            NO,  YES, YES, YES, "Leica: DIVE (STELLARIS platform); Zeiss: LSM Airyscan; Nikon: AX R MP"),
    ("Super-Resolution Microscopes",       NO,  YES, YES, YES, "Leica: STELLARIS STED; Zeiss: dedicated category; Nikon: N-STORM, STEDYCON, DeepSIM"),
    ("Light Sheet Microscopes",            NO,  YES, YES, YES, "Leica: /products/light-microscopess (DLS); Zeiss: dedicated category; Nikon: products page"),
    ("Darkfield Microscopes",              YES, NO,  NO,  NO,  "AMScope: /collections/compound-microscopes-darkfield; /collections/special-microscopes-darkfield"),
    ("Phase Contrast Microscopes",         YES, NO,  NO,  NO,  "AMScope: /collections/special-microscopes-phase-contrast"),
    ("Polarizing Microscopes",             YES, NO,  YES, YES, "AMScope: /collections/special-microscopes-polarizing; Zeiss: Axio Imager 2 Pol; Nikon: dedicated category"),
    ("Epi-Fluorescence Microscopes",       YES, NO,  NO,  NO,  "AMScope: /collections/special-microscopes-epi-fluorescence"),
    ("Fluorescence Microscopes",           YES, YES, YES, YES, "All competitors surface fluorescence as a top-level capability/category"),
    ("Metallurgical Microscopes",          YES, NO,  NO,  NO,  "AMScope: /collections/special-microscopes-metallurgical"),
    ("Surgical Microscopes",               NO,  YES, NO,  NO,  "Leica: dedicated surgical microscopes section (ophthalmology, neurosurgery, ENT, dental)"),
    ("Automated Microscopes",              NO,  YES, NO,  NO,  "Leica: /products/light-microscopes/ lists automated as distinct category"),
    ("Cell Culture Microscopes",           NO,  YES, NO,  NO,  "Leica: Mateo TL/FL series"),
    ("Spinning Disk Confocal",             NO,  NO,  NO,  YES, "Nikon: Yokogawa CSU-W1 SoRa, Crest X-Light"),
    ("TIRF Microscopes",                   NO,  NO,  NO,  YES, "Nikon: Ti2-LAPP Photostimulation & TIRF"),
    ("High Content Imaging",               NO,  NO,  NO,  YES, "Nikon: dedicated high content imaging category"),
    ("Cell Screening",                     NO,  NO,  NO,  YES, "Nikon: dedicated cell screening category"),
    ("Slide Scanning / WSI",               NO,  NO,  YES, YES, "Zeiss: AxioScan 7; Nikon: dedicated slide scanning category"),
    ("FLIM Microscopes",                   NO,  YES, NO,  NO,  "Leica: STELLARIS FALCON"),
    ("CRS / Label-Free Microscopes",       NO,  YES, NO,  NO,  "Leica: STELLARIS CRS"),
    ("Patch-Clamp Microscopes",            NO,  NO,  YES, NO,  "Zeiss: Axio Examiner (fixed-stage for electrophysiology)"),
    # Application Categories
    ("── APPLICATION CATEGORIES ──", "", "", "", "", ""),
    ("Industrial Inspection",              YES, NO,  NO,  NO,  "AMScope: /collections/applications-industrial-microscopes; /collections/stereo-microscopes-industrial-inspection"),
    ("Medical / Clinical",                 YES, NO,  NO,  NO,  "AMScope: /collections/applications-medical-microbiology-microscopes"),
    ("Veterinary / Zoology",               YES, NO,  NO,  NO,  "AMScope: /collections/applications-veterinary-zoology"),
    ("Blood Analysis",                     YES, NO,  NO,  NO,  "AMScope: /collections/compound-microscopes-blood-analysis"),
    ("Forensics / Fiber Analysis",         YES, NO,  NO,  NO,  "AMScope: /collections/applications-fiber-hair-forensics-microscopes"),
    ("Gemology / Jewelry",                 YES, NO,  NO,  NO,  "AMScope: /collections/stereo-microscopes-gemology-jewelry; /collections/applications-jewelry-gemology-microscopes"),
    ("Mineralogy",                         YES, NO,  NO,  NO,  "AMScope: /collections/stereo-microscopes-mineralogy"),
    ("Dental Lab",                         YES, YES, NO,  NO,  "AMScope: /collections/stereo-microscopes-dental-lab; Leica: surgical dental microscopes"),
    ("Botany",                             YES, NO,  NO,  NO,  "AMScope: /collections/applications-botany-microscopes"),
    ("Wafer / Semiconductor",              YES, NO,  NO,  NO,  "AMScope: /collections/wafer-semiconductor-microscopes"),
    ("Teaching / Student",                 YES, NO,  YES, NO,  "AMScope: /collections/compound-microscopes-teaching-training; Zeiss: Primo series"),
    ("Electronics / PCB",                  YES, NO,  NO,  NO,  "AMScope: /collections/applications-electronics-microscopes"),
    # Camera Categories
    ("── CAMERA CATEGORIES ──", "", "", "", "", ""),
    ("Microscope Cameras (general)",       YES, YES, YES, YES, "AMScope: /collections/microscope-cameras; Leica: camera product page; Zeiss: Axiocam; Nikon: Digital Sight series"),
    ("sCMOS Cameras",                      NO,  YES, YES, NO,  "Leica: K8, K3 series; Zeiss: Axiocam 820/807"),
    ("High-Speed Cameras",                 YES, NO,  NO,  NO,  "AMScope: /collections/digital-microscopes-high-speed-microscope-cameras"),
    ("Low-Light / Cooled Cameras",         YES, YES, NO,  NO,  "AMScope: /collections/digital-microscopes-low-light-microscope-cameras; Leica: passively cooled options"),
    ("USB Cameras",                        YES, NO,  NO,  NO,  "AMScope: /collections/usb-microscope-cameras"),
    ("HDMI Microscopes / Cameras",         YES, NO,  NO,  NO,  "AMScope: /collections/hdmi-digital-microscopes"),
    ("Wi-Fi Enabled Cameras",              YES, NO,  NO,  NO,  "AMScope: /collections/digital-microscopes-wi-fi-enabled-digital-microscope-cameras"),
    ("4K Compatible Cameras",              YES, NO,  NO,  NO,  "AMScope: /collections/digital-microscopes-4k-compatible"),
    ("Autofocus Cameras",                  YES, NO,  NO,  NO,  "AMScope: /collections/digital-microscopes-autofocus-microscope-cameras"),
    ("Fluorescence Cameras (mono)",        NO,  YES, YES, NO,  "Leica: fluorescence camera product line; Zeiss: Axiocam mono high-end category"),
    ("Polarization Cameras",               NO,  NO,  YES, NO,  "Zeiss: Axiocam 705 pol"),
]

SECTION_FILL = PatternFill("solid", fgColor="1F4E79")
SECTION_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)

for row_idx, row in enumerate(matrix_data, 2):
    is_section = row[0].startswith("──")
    fill = ALT_FILL if (row_idx % 2 == 0 and not is_section) else None
    for col_idx, value in enumerate(row, 1):
        c = ws3.cell(row=row_idx, column=col_idx, value=value)
        if is_section:
            c.font = SECTION_FONT
            c.fill = SECTION_FILL
            c.alignment = CENTER
            c.border = BORDER
        else:
            style_cell(c, fill=fill)
            if col_idx in (2, 3, 4, 5):
                c.alignment = CENTER

ws3.freeze_panes = "A2"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 – METHODOLOGY NOTES
# ══════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Methodology")
ws4.column_dimensions["A"].width = 25
ws4.column_dimensions["B"].width = 90

notes = [
    ("Research Date", "2026-03-11"),
    ("Analyst", "Claude (SEO Agent)"),
    ("", ""),
    ("SOURCES USED", ""),
    ("Evident Scientific sitemap", "https://evidentscientific.com/sitemap-en.xml  (from robots.txt)"),
    ("AMScope collections sitemap", "https://amscope.com/sitemap_collections_1.xml?from=240584655023&to=359037337775"),
    ("AMScope sitemap index", "https://amscope.com/sitemap.xml"),
    ("Zeiss microscopy sitemap", "https://www.zeiss.com/microscopy/us/home.sitemap.xml  (from zeiss.com/sitemap.xml index)"),
    ("Leica sitemap index", "https://www.leica-microsystems.com/sitemap.xml  (from leica-microsystems.com/robots.txt)"),
    ("Nikon HC sitemap index", "https://www.microscope.healthcare.nikon.com/sitemap.xml"),
    ("Nikon HC singles sitemap", "https://www.microscope.healthcare.nikon.com/sitemap-singles.xml"),
    ("Leica products page", "https://www.leica-microsystems.com/products/  (URL from sitemap)"),
    ("Leica digital microscopes page", "https://www.leica-microsystems.com/products/digital-microscopes/  (URL from sitemap)"),
    ("Leica cameras page", "https://www.leica-microsystems.com/products/microscope-cameras/  (URL from sitemap)"),
    ("Leica confocal page", "https://www.leica-microsystems.com/products/confocal-microscopes/  (URL from sitemap)"),
    ("Leica surgical page", "https://www.leica-microsystems.com/products/surgical-microscopes/  (URL from sitemap)"),
    ("Zeiss products overview", "https://www.zeiss.com/microscopy/en/products.html  (URL from sitemap)"),
    ("Zeiss widefield page", "https://www.zeiss.com/microscopy/en/products/light-microscopes/widefield-microscopes.html  (URL from sitemap)"),
    ("Zeiss cameras page", "https://www.zeiss.com/microscopy/en/products/cameras.html  (URL from sitemap)"),
    ("Nikon HC products page", "https://www.microscope.healthcare.nikon.com/products  (URL from sitemap-singles.xml)"),
    ("Evident homepage", "https://www.evidentscientific.com/en/  (to confirm product lines referenced on homepage)"),
    ("", ""),
    ("METHODOLOGY", ""),
    ("Step 1", "Retrieved robots.txt from evidentscientific.com and amscope.com to find sitemap index URLs"),
    ("Step 2", "Fetched sitemap index files for all four competitors to identify available sitemaps"),
    ("Step 3", "Fetched confirmed sitemap files — ONLY URLs sourced from sitemaps were crawled"),
    ("Step 4", "Fetched product/category pages whose URLs appeared in confirmed sitemaps"),
    ("Step 5", "Cross-referenced competitor categories against Evident Scientific's confirmed URL paths"),
    ("Step 6", "Categorized gaps by priority: High (Evident has product, no page), Medium (application opportunity), Low (niche/out-of-scope)"),
    ("", ""),
    ("PRIORITY DEFINITIONS", ""),
    ("High", "Evident Scientific has products that belong in this category but no dedicated category page exists in the sitemap"),
    ("Medium", "Application/segment category that Evident's existing products could target — moderate SEO opportunity"),
    ("Low", "Niche category or potentially out of scope for Evident's core B2B research/clinical positioning"),
]

for row_idx, (key, val) in enumerate(notes, 1):
    ck = ws4.cell(row=row_idx, column=1, value=key)
    cv = ws4.cell(row=row_idx, column=2, value=val)
    if key in ("SOURCES USED", "METHODOLOGY", "PRIORITY DEFINITIONS"):
        ck.font = Font(name="Calibri", bold=True, size=11)
        ck.fill = SUBHDR_FILL
        cv.fill = SUBHDR_FILL
    elif key:
        ck.font = Font(name="Calibri", bold=True, size=10)
    ck.alignment = Alignment(vertical="top")
    cv.alignment = Alignment(wrap_text=True, vertical="top")
    ws4.row_dimensions[row_idx].height = 18

# ── Save ──────────────────────────────────────────────────────────────────────
out_path = "/home/user/category-expansion/competitor_category_gaps.xlsx"
wb.save(out_path)
print(f"Saved: {out_path}")
