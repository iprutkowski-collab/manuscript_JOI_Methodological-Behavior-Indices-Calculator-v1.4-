Methodological Behavior Indices Calculator — version 1.4

What is new in v1.4
- Conservative bundle assignment rule.
- If sig_ordinal = 1, the document is not assigned to a pure
  bundle_likert_parametric_profile.
- Such cases are routed to bundle_likert_other_mixed unless explicit
  nonparametric-treatment evidence supports bundle_likert_nonparametric_profile.
- This is intended to avoid implying methodological correctness when ordinal
  measurement-level language is visible together with parametric cues.

Recommended Windows build command

cd /d C:\JOI_APP
py -3 -m pip install --upgrade pip
py -3 -m pip install pyinstaller python-docx pymupdf pypdf pdfplumber striprtf pillow pytesseract pywin32
py -3 -m PyInstaller --onefile --windowed --name MethodologicalIndicesApp methodological_indices_app_v1_4_conservative.py
