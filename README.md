# Esophageal Motility Study Pipeline

A streamlined pipeline for processing esophageal manometry reports in `.docx` format. It automates the extraction of patient data, diagnostic tables, embedded images, and textual findings — all neatly structured for analysis and integration.

---

## Features

- Extracts patient info and clinical tables to `.csv`
- Extracts and renames embedded images from reports
- Removes grid lines from diagrams using OpenCV
- Parses diagnostic text sections into structured `.json`
- Outputs organized, analysis-ready folders

---

## ⚙️ How It Works

### 1. Text & Table Extraction
- Reads `.docx` tables to extract:
  - Patient details
  - Summary metrics
  - Esophageal & UES motility data
- Saves each as a separate `.csv`

### 2. Image Extraction & Naming
- Extracts diagrams embedded in the document
- Names them using:
  - Custom defaults (e.g., Swallow Composite)
  - Patient metadata from CSV

### 3. Image Preprocessing
- Detects and removes grid lines via:
  - Canny edge detection
  - Hough Transform
  - Inpainting (OpenCV)
- Saves clean diagrams to `processed_images/`

### 4. Diagnostic Text Extraction
- Captures key sections like:
  - Chicago Classification Findings
  - Procedure, Indications, Impressions
- Exports to a structured `.json`

---

## Output Structure

```
.
├── extracted_data/
│   ├── Patient_details.csv
│   ├── Esophageal_Manometry_Summary.csv
│   ├── Lower_Esophageal_Sphincter.csv
│   ├── Esophageal_Motility.csv
│   ├── Upper_Esophageal_Sphincter.csv
│   ├── Pharyngeal_UES_Motility.csv
│   ├── Image_filenames.csv
│   └── chicago_classification_findings.json
├── images/
│   └── *.png         # Original extracted images
├── processed_images/
│   └── *.png         # Grid-line removed versions
└── subj.docx         # Input report
```

---

## 🔧 Requirements

Install dependencies with:

```bash
pip install -r requirements.txt
```



