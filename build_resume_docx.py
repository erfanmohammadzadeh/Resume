# -*- coding: utf-8 -*-
"""Build a professional resume .docx using only the Python standard library."""
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).resolve().parent / "Erfan_Mohammadzadeh_Resume.docx"
DESKTOP = Path(r"C:\Users\Amvaj Negar - EM\Desktop") / "Erfan_Mohammadzadeh_Resume.docx"

NAVY = "1B365D"
ACCENT = "2C5F8A"
MUTED = "555555"
BODY = "222222"
RULE = "C5CDD6"


def t(text, xml_space=False):
    attrs = ' xml:space="preserve"' if xml_space else ""
    return f"<w:t{attrs}>{escape(text)}</w:t>"


def r(text, *, bold=False, italic=False, size=21, color=BODY, font="Calibri", space=False):
    rpr = [
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>',
        f'<w:sz w:val="{size}"/>',
        f'<w:szCs w:val="{size}"/>',
        f'<w:color w:val="{color}"/>',
    ]
    if bold:
        rpr.append("<w:b/><w:bCs/>")
    if italic:
        rpr.append("<w:i/><w:iCs/>")
    return f"<w:r><w:rPr>{''.join(rpr)}</w:rPr>{t(text, space)}</w:r>"


def hyperlink(rid, text, *, size=18, color=ACCENT):
    rpr = (
        f'<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        f'<w:color w:val="{color}"/><w:u w:val="single"/>'
    )
    return (
        f'<w:hyperlink r:id="{rid}" w:history="1">'
        f"<w:r><w:rPr>{rpr}</w:rPr>{t(text)}</w:r></w:hyperlink>"
    )


def p(runs, *, align="left", before=0, after=60, line=240, border=False, rtl=False):
    jc = f'<w:jc w:val="{align}"/>'
    sp = f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
    ind = ""
    bdr = ""
    if border:
        bdr = (
            f'<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" '
            f'w:color="{NAVY}"/></w:pBdr>'
        )
    if rtl:
        ind = "<w:bidi/>"
    return f"<w:p><w:pPr>{jc}{sp}{bdr}{ind}</w:pPr>{''.join(runs)}</w:p>"


def heading(text):
    return p(
        [r(text.upper(), bold=True, size=22, color=NAVY, font="Calibri")],
        before=200,
        after=80,
        border=True,
    )


def job_header(title, company, dates):
    return (
        p(
            [
                r(title, bold=True, size=22, color=BODY),
                r("  |  ", size=20, color=MUTED),
                r(company, bold=True, size=21, color=ACCENT),
            ],
            after=0,
        )
        + p([r(dates, italic=True, size=18, color=MUTED)], before=0, after=40)
    )


def bullet(text_parts):
    """text_parts is a list of r() fragments already built, or a plain string."""
    if isinstance(text_parts, str):
        runs = [r(text_parts, size=20)]
    else:
        runs = text_parts
    ppr = (
        '<w:pPr>'
        '<w:spacing w:before="20" w:after="40" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:left="360" w:hanging="200"/>'
        "</w:pPr>"
    )
    mark = r("•  ", size=20, color=ACCENT, space=True)
    return f"<w:p>{ppr}{mark}{''.join(runs)}</w:p>"


def rb(text, **kw):
    kw.setdefault("size", 20)
    return r(text, bold=True, **kw)


def rn(text, **kw):
    kw.setdefault("size", 20)
    return r(text, **kw)


body = []

# Header
body.append(
    p(
        [r("ERFAN MOHAMMADZADEH", bold=True, size=44, color=NAVY, font="Calibri")],
        align="center",
        after=40,
    )
)
body.append(
    p(
        [
            r(
                "Software Engineer  ·  Data Engineer  ·  C++ / Qt  ·  C# / .NET Web API",
                size=18,
                color=ACCENT,
            )
        ],
        align="center",
        after=80,
    )
)
body.append(
    p(
        [
            r("Iran  ·  ", size=18, color=MUTED),
            r("(+98) 913 406 5696  ·  ", size=18, color=MUTED),
            hyperlink("rIdEmail", "erfanmohammadzadeh.en@gmail.com", size=18),
            r("  ·  ", size=18, color=MUTED),
            hyperlink("rIdLi", "LinkedIn", size=18),
            r("  ·  ", size=18, color=MUTED),
            hyperlink("rIdGh", "GitHub", size=18),
        ],
        align="center",
        after=120,
        border=True,
    )
)

body.append(heading("Professional Summary"))
body.append(
    p(
        [
            rn(
                "Software and data engineer specializing in high-performance C++/Qt applications, "
                "sensor and signal pipelines, and structured geospatial reconstruction. I also build "
                "C# / .NET backends: ASP.NET Core Web APIs, REST CRUD endpoints, and database-backed "
                "persistence. I take raw measurements (cameras, LiDAR, ECG) through filtering, feature "
                "extraction, validation, visualization, and export—and store or expose results through "
                "SQL and structured APIs. My next step is to deepen production .NET service and database work."
            )
        ],
        after=80,
    )
)

body.append(heading("Technical Skills"))
skills = [
    ("Languages", "C++, C#, Python, SQL, QML"),
    (".NET / backend", "ASP.NET Core, Web API, REST CRUD, service-layer APIs"),
    ("Data & storage", "SQL databases, SQLite, XML, structured ETL-style pipelines"),
    ("Desktop & UI", "Qt (Widgets / QML)"),
    ("Vision & 3D", "OpenCV, PCL, VTK, CGAL, Open3D"),
    ("Geospatial", "PDAL, GDAL, QGIS, City4CFD / LoD modeling"),
    ("Systems", "Linux (LPIC-1), cross-platform desktop builds"),
    ("Domains", "ECG / biomedical DSP, camera calibration, ANPR, LiDAR city models"),
]
for label, val in skills:
    body.append(
        p(
            [rb(label + ":  ", size=20, space=True), rn(val)],
            before=20,
            after=20,
        )
    )

body.append(heading("Career Direction"))
body.append(
    p(
        [
            rn(
                "Growing toward C# / .NET backend engineering: ASP.NET Core Web APIs, CRUD over "
                "relational data, and connecting desktop and processing products to maintainable "
                "service and database layers."
            )
        ],
        after=80,
    )
)

body.append(heading("Work Experience"))

body.append(
    job_header(
        "Geospatial Data Engineer",
        "Image Horizon (Data Horizon)  ·  Tehran, Iran",
        "November 2025 – Present  (concurrent with Amvaj Negar)",
    )
)
body.append(
    bullet(
        [
            rn("Own the "),
            rb("City4CFD / QCity4CFD"),
            rn(" reconstruction path: point clouds and building footprints to "),
            rb("LoD 2.2"),
            rn(" city meshes for CFD, covering "),
            rb("20,000+ buildings"),
            rn("."),
        ]
    )
)
body.append(
    bullet(
        [
            rn("Built processing stages with "),
            rb("PCL, PDAL, GDAL, CGAL, and VTK"),
            rn(": filtering, segmentation, surface reconstruction, geometric regularization, rendering, and mesh QA."),
        ]
    )
)
body.append(
    bullet(
        [
            rn("Reported "),
            rb(">90% reconstruction accuracy"),
            rn(" on the high-detail building pipeline (Python / NumPy / Open3D / PDAL + GIS in QGIS / ArcGIS)."),
        ]
    )
)
body.append(
    bullet(
        "Bridged GIS operators and simulation teams by producing inspectable, simulation-ready geometry instead of raw point clouds."
    )
)

body.append(
    job_header(
        "Senior Software Engineer",
        "Amvaj Negar Sepahan Co.  ·  Isfahan, Iran",
        "April 2024 – Present",
    )
)
body.append(
    bullet(
        [
            rn("Designed and maintain "),
            rb("Holter ECG desktop software"),
            rn(" (C++ / Qt Widgets): long-term ECG ingest, "),
            rb("P-Q-R-S-T"),
            rn(" detection, beat-template classification, arrhythmia support, SQLite persistence, XML/PDF reporting."),
        ]
    )
)
body.append(
    bullet(
        [
            rn("Built "),
            rb("QCardio"),
            rn(", a validation harness for ECG libraries: converted "),
            rb("MIT-BIH Arrhythmia"),
            rn(" into a structured test set so algorithms are checked against physician-annotated ground truth."),
        ]
    )
)
body.append(
    bullet(
        "Own signal-processing correctness: filtering, feature extraction, classification, and regression tests that clinicians and engineers can both trust."
    )
)

body.append(
    job_header(
        "Software Engineer",
        "Data Image Rayan Co.  ·  Isfahan, Iran",
        "June 2024 – March 2025  (concurrent with Amvaj Negar)",
    )
)
body.append(
    bullet(
        [
            rn("Delivered "),
            rb("ANPR"),
            rn(" monitoring for parking access: live cameras, vehicle detection, plate crop, OCR validation, and barrier control."),
        ]
    )
)
body.append(
    bullet(
        "Reduced manual gate intervention by closing the loop from video frame to access decision with a real-time image pipeline."
    )
)

body.append(
    job_header(
        "Junior Software Engineer",
        "Tivan Sanat (Dade Pardazan Tivan Sanat)  ·  Isfahan, Iran",
        "April 2021 – December 2022",
    )
)
body.append(
    bullet(
        [
            rn("Shipped a cross-platform "),
            rb("C++/Qt + OpenCV"),
            rn(" camera-calibration tool: target detection, keypoints, intrinsics/extrinsics, distortion (radial/tangential), focal length, principal point, FOV, and correction matrices."),
        ]
    )
)
body.append(
    bullet(
        "Exported calibration results (SQLite / XML / PDF) for production optical QA."
    )
)

body.append(heading("Selected Projects"))

projects = [
    (
        "ASP.NET Core Web API (CRUD)",
        "Current backend practice",
        "REST endpoints for create / read / update / delete against a SQL database; request handling, data access, and API structure for service-oriented products.",
    ),
    (
        "QCity4CFD",
        "Data Horizon  ·  Jan 2025",
        "LoD 2.2 reconstruction from LiDAR and footprints; mesh regularization and CFD-oriented city models. Python, Open3D, PDAL, CGAL, QGIS/ArcGIS, City4CFD.",
    ),
    (
        "QCardio",
        "Amvaj Negar Sepahan  ·  Jun 2026",
        "ECG library validation against annotated MIT-BIH records; automated pass/fail on clinical waveforms.",
    ),
    (
        "ECG Holter Software",
        "Amvaj Negar Sepahan  ·  Apr 2024",
        "Desktop Holter analysis: wave detection, templates, arrhythmia flags, reports. C++, Qt Widgets, SQLite, XML/PDF.",
    ),
    (
        "Visible-camera parameter tester",
        "Tivan Sanat  ·  Dec 2022",
        "Production calibration of practical camera parameters for accurate rendering and metrology. C++, Qt, OpenCV.",
    ),
]
for name, meta, desc in projects:
    body.append(
        p(
            [rb(name, size=21), rn("  —  " + meta, size=19, color=MUTED)],
            before=80,
            after=20,
        )
    )
    body.append(p([rn(desc)], after=40))

body.append(heading("Education"))
body.append(
    p(
        [
            rb("B.Sc. Electrical Engineering — Communication Systems"),
        ],
        after=0,
    )
)
body.append(
    p(
        [r("Semnan University  ·  Semnan, Iran  ·  Oct 2019 – Jan 2024  ·  GPA 3.2 / 4.0", italic=True, size=18, color=MUTED)],
        after=40,
    )
)
body.append(
    p(
        [
            rn(
                "Coursework and practice in signal processing, communications, and control; applied DSP "
                "(filtering, features, real-time classification) in subsequent industry systems."
            )
        ]
    )
)

body.append(heading("Certificates"))
body.append(
    p(
        [rb("Foundations of Coding: Full-Stack"), rn("  —  Microsoft / Coursera  ·  Oct 2025")],
        after=20,
    )
)
body.append(
    p([hyperlink("rIdCert", "coursera.org/account/accomplishments/verify/XIAL95E2ZNBP", size=18)], after=80)
)

body.append(heading("Languages"))
body.append(p([rb("Persian (Farsi):  ", space=True), rn("Native")], after=20))
body.append(
    p(
        [
            rb("English:  ", space=True),
            rn(
                "Professional working proficiency (reading, writing, speaking, listening) — used for technical documentation, code, and work communication."
            ),
        ]
    )
)

document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {''.join(body)}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/>
    </w:sectPr>
  </w:body>
</w:document>
'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdEmail" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:erfanmohammadzadeh.en@gmail.com" TargetMode="External"/>
  <Relationship Id="rIdLi" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.linkedin.com/in/erfan-mohammadzade-076791178" TargetMode="External"/>
  <Relationship Id="rIdGh" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://github.com/erfan-mohammadzade" TargetMode="External"/>
  <Relationship Id="rIdCert" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.coursera.org/account/accomplishments/verify/XIAL95E2ZNBP" TargetMode="External"/>
</Relationships>
'''

styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="21"/><w:szCs w:val="21"/>
        <w:color w:val="222222"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="60" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
'''

core = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Erfan Mohammadzadeh — Resume</dc:title>
  <dc:creator>Erfan Mohammadzadeh</dc:creator>
  <cp:lastModifiedBy>Erfan Mohammadzadeh</cp:lastModifiedBy>
</cp:coreProperties>
'''

app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Word</Application>
</Properties>
'''


def write_docx(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)


write_docx(OUT)
try:
    write_docx(DESKTOP)
except OSError:
    pass
print(OUT)
print("bytes", OUT.stat().st_size)
