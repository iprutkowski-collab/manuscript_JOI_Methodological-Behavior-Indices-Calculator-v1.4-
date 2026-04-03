# manuscript_JOI_Methodological-Behavior-Indices-Calculator-v1.4-
A prototype local tool for measuring visible methodological behavior in scholarly texts using SPI, MVI, BCS, and AUS.
# Methodological Behavior Indices Calculator (v1.4)

A prototype local software tool for measuring **visible methodological behavior** in scholarly texts.

This application implements the bounded indicator logic developed in the manuscript:

**“Measuring Methodological Behavior in the Scientific Record: A Framework for Metadata- and Text-Based Indicator Construction and Validation.”**

The tool detects operational methodological signals in article text and computes four indices:

- **SPI** — Signal Presence Index
- **MVI** — Methodological Visibility Index
- **BCS** — Bundle Coherence Score
- **AUS** — Ambiguity / Uncertainty Score

It also assigns a bounded bundle profile and produces a short interpretive report.

## Current version

**Version:** 1.4  
**Bundle logic:** conservative  
**Important rule:** if `sig_ordinal = 1`, the document is **not** assigned to a pure parametric bundle.

## Main functions

The tool can read:

- PDF
- DOCX
- DOC (Windows + Microsoft Word + `pywin32`)
- TXT
- MD
- CSV
- RTF

The program:

1. loads article text,
2. optionally excludes the References section,
3. detects methodological signals,
4. computes SPI, MVI, BCS, and AUS,
5. assigns a bounded bundle profile,
6. exports a report.

## Bundle profiles

The current bounded implementation uses three bundle profiles:

- `bundle_likert_parametric_profile`
- `bundle_likert_nonparametric_profile`
- `bundle_likert_other_mixed`

These bundle labels should be interpreted as **co-occurrence profiles**, not as ontological statements about the scale itself and not as automatic judgments of methodological correctness.

## Supported interpretation logic

The software is designed to measure:

- breadth of visible methodological traces,
- intensity of methodological disclosure,
- coherence of visible bundle structure,
- uncertainty caused by ambiguity or weak contextualization.

It is **visibility-oriented**, not correctness-oriented.

## Input and output

### Input
A scholarly text file in one of the supported formats.

### Output
A local report containing:

- word count,
- total methodological hits,
- ambiguous hits,
- active signals,
- active classes,
- SPI,
- MVI,
- BCS,
- AUS,
- dominant bundle,
- signal-level flags,
- class-level flags,
- bundle membership scores,
- interpretation summary.

## Requirements

See `requirements.txt`.

Recommended Python version:

- **Python 3.10+**
- Windows 64-bit recommended

## Installation

Create and activate a virtual environment:

```bash
python -m venv venv
