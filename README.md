# 🦟 Malaria Rule-Based Data Validation and Indicator App

![Streamlit](https://img.shields.io/badge/Streamlit-App-red)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-black)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel%20Support-green)

A **Streamlit web application** for validating malaria datasets using **custom rule-based Python functions**.  
Users upload an Excel workbook and a rule file, and the app applies validation rules, generates processed sheets, and produces an **Error Summary report**.

---

## Features

- Upload malaria Excel datasets  
- Dynamic rule engine using `malaria_*` functions  
- Sheet-by-sheet rule selection  
- Validation and cleaning for malaria datasets  
- Automatic **Error Summary** generation  
- Downloadable processed Excel workbook  
- Support for **Positive**, **Aggregate**, and **REACH Recheck** workflows  

---

## Project Structure

```bash
project/
│
├── app.py          # Streamlit application
├── rules.py        # Rule functions
└── README.md
How It Works

Upload a rule file (.py)

Upload a malaria Excel workbook

Select the rule for each sheet

Click Run

Preview processed outputs

Download the processed workbook

Rule File Requirements

The uploaded rule file must contain functions whose names start with:

malaria_

Example:

def malaria_positive(df):
    # validation / cleaning logic
    return df
Output Workbook

The processed file will look like:

malaria_processed_with_summary.xlsx
│
├── Error Summary
├── Processed - Positive
├── Processed - Aggregate
└── Processed - REACH_RECHECK



Notes

Only selected sheets are processed

Original sheets remain unchanged

Processed sheets are added to a new workbook

Error Summary is generated automatically

malaria_recheck requires both Positive and Aggregate data

Adding New Rules (Developer)

To add a new rule, define a function in rules.py:

def malaria_example(df):
    return df

Rules must:

Start with malaria_

Accept a DataFrame

Return a DataFrame

Use a COMMENT column for validation messages

License

Internal project for malaria validation.
- check your README after you commit  
- or help you add **real screenshots into GitHub** 📸
