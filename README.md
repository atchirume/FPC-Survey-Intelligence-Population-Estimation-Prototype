# FPC Survey Intelligence & Population Estimation Platform

**Public Prototype Application**  
**Developed by Admire Tarisirayi Chirume**

## Overview

The FPC Survey Intelligence & Population Estimation Platform is a prototype Streamlit application for Foreign Private Capital (FPC) survey processing, statistical quality assurance, BPM6/BOP/IIP compiler-support analytics, sampling-frame management, survey weighting and population estimation.

The platform combines:

- PDF and OCR-based survey extraction
- Survey review and editing
- Accounting validation
- Data cleaning and quality assurance
- Sampling-frame profiling
- Design-weight calculation
- Nonresponse adjustment
- Optional ratio calibration
- Population and domain estimation
- BOP/IIP compiler-support views
- Dashboard analytics
- Excel, Word, PowerPoint and PDF reporting
- Export and audit workflows

## Population-estimation workflow

1. Upload a statistical business register or survey sampling frame.
2. Select the relevant worksheet and enterprise identifiers.
3. Define survey estimation strata.
4. Define sample-selection and response status.
5. Calculate design weights.
6. Apply within-stratum nonresponse adjustment.
7. Optionally apply ratio calibration to a known benchmark such as employment or turnover.
8. Calculate national and domain estimates.
9. Review weight diagnostics, zero-response strata and estimation QA.
10. Generate sample or population reports.

## Demonstration data

The repository contains a **synthetic demonstration sampling frame** with 2,500 establishments covering Zimbabwe's provinces and ISIC Rev.4 sections.

Enterprise names in the public demonstration workbook have been anonymised as synthetic enterprise identifiers. The workbook is intended for testing the application's methodology and workflow and **must not be interpreted as official statistics or an official population frame**.

## Important public-prototype notice

This application is intended for research, analytical and demonstration purposes.

It is **not an officially authorised institutional system** and should not be interpreted as representing the policies, decisions, systems or official statistics of any government or other institution.

Do **not** upload confidential, personal, commercially sensitive or production survey data to the public demonstration application.

## Technology

- Python
- Streamlit
- Pandas
- NumPy
- PyMuPDF
- Tesseract OCR
- Plotly
- OpenPyXL
- python-docx
- python-pptx
- ReportLab

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Author

**Admire Tarisirayi Chirume**

The application brings together statistical data extraction, data engineering, quantitative analysis, external-sector statistical concepts and software development into an integrated FPC survey workflow.

## Status

**Public prototype / demonstration release.**
