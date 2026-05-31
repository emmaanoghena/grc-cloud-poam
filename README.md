# GRC Project 2 – Cloud Security Risk Register & POA&M Generator

## Overview

This project automates the ingestion of mock cloud security findings, applies a multi-factor risk scoring model, and produces a color-coded Excel workbook in **POA&M (Plan of Action & Milestones)** format aligned to **NIST SP 800-53**.

---

## Project Structure

grc-cloud-poam/
├── data/
│   └── mock_findings.csv          # Mock cloud security findings (AWS, Prisma Cloud, Wiz)
├── scripts/
│   └── generate_poam.py           # Main script: scoring + workbook generation
├── output/
│   └── cloud_security_poam.xlsx   # Generated POA&M workbook (git-ignored)
├── .gitignore
└── README.md

## Risk Scoring Formula

risk_score = (
(severity_weight  × 4) +   # Critical=10, High=7, Medium=4, Low=2
(asset_criticality × 3) +  # High=3, Medium=2, Low=1
(exposure          × 2) +  # Internet=3, Internal=2, Isolated=1
(data_sensitivity  × 1)    # PII=3, Confidential=2, Public=1
) × 2.5


| Level    | Score Range |
|----------|-------------|
| Critical | 80 – 100+   |
| High     | 60 – 79     |
| Medium   | 40 – 59     |
| Low      | 0 – 39      |

## Usage

```bash
git clone <your-repo-url>
cd grc-cloud-poam
pip install openpyxl pandas
python scripts/generate_poam.py
open output/cloud_security_poam.xlsx
```

## Controls Mapping

Findings are tagged to NIST SP 800-53 control families:
AC · AU · CA · CM · CP · IA · IR · SC · SI

## Extending This Project

Replace `data/mock_findings.csv` with real exports from:
- **AWS Security Hub** → `aws securityhub get-findings`
- **Prisma Cloud** → Alerts API or CSV export
- **Wiz** → Issues export from the Wiz console
