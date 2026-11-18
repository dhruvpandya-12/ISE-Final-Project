MEDICAL DEVICE RECALL PREVENTION PROJECT
How to run
Work in progress

What you need
Python 3.10 or newer. Check with python3 --version on Mac, or python --version on Windows.
The two FDA zip files from https://open.fda.gov/data/downloads/ in the Medical Device section:
device-recall-0001-of-0001_json.zip (Device Recall)
device-enforcement-0001-of-0001_json.zip (Device Recall Enforcement Reports)
If your downloaded names differ slightly, rename them to exactly these names and put them in the data folder.
The raw FDA zip files are large and are not included in this repo.

Folder layout
recall_project_v2/
    README.txt
    requirements.txt
    step1_extract.py
    step2_analysis.py
    step3_excel.py
    step4_excel_ise250.py
    data/
    outputs/
        study_events.csv and the result tables in CSV form
        Recall_Analysis_Calculations.xlsx
        Recall_Analysis_ISE250_Style.xlsx

Run it
1. Open a terminal in the recall_project_v2 folder.
2. Install the libraries once: python3 -m pip install -r requirements.txt
3. Run the steps in order:
   python3 step1_extract.py extracts the needed fields from the FDA zip files.
   python3 step2_analysis.py builds the event table and runs the analysis, including the logistic regression.
   python3 step3_excel.py builds Recall_Analysis_Calculations.xlsx.
   python3 step4_excel_ise250.py builds Recall_Analysis_ISE250_Style.xlsx, laid out like the ISE 250 class material.
   Open the workbooks in Excel. They have many formulas, so the first open can take a few seconds.

What you should see after step 2, with data downloaded 27 Oct 2026
Recall events studied, 2013 to 2025: 13,410, covering 36,287 recalled products.
Class I events 2020 to 2025: 37 to 125.
Class I share 3.5 percent in 2016 to 2018, and 11.5 percent in 2023 to 2025, z = 12.16.
Root cause not determined: 20.7 percent of events.
Top priorities: Production process, and Design (hardware).
Logistic regression: 10,640 events, McFadden R2 = 0.216, AUC = 0.837.
If you download newer FDA files, the numbers will change slightly because FDA adds new recalls weekly.

Dataset and outputs
outputs/study_events.csv is the main dataset, with the 13,410 events analysed.
outputs/table_by_year.csv
outputs/table_categories.csv
outputs/table_root_cause_groups.csv
outputs/table_pareto_detailed_causes.csv
outputs/table_pchart_quarterly.csv
outputs/table_radiology.csv
outputs/table_logistic_regression.csv
outputs/Recall_Analysis_Calculations.xlsx
outputs/Recall_Analysis_ISE250_Style.xlsx

Common problems
No module named: run the pip install step again.
No such file or directory, data: check the zip files are inside the data folder with the exact names above.
python3 command not found on Windows: use python instead of python3.

Note
The written report is still in progress and is not included in this repo.
