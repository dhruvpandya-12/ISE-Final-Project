MEDICAL DEVICE RECALL PREVENTION PROJECT - HOW TO RUN
=====================================================

What you need
-------------
- Python 3.10 or newer (check with:  python3 --version  on Mac, or  python --version  on Windows)
- The two FDA zip files from https://open.fda.gov/data/downloads/  (section "Medical Device"):
    device-recall-0001-of-0001_json.zip       (Device Recall)
    device-enforcement-0001-of-0001_json.zip  (Device Recall Enforcement Reports)
  If your downloaded names differ slightly, rename them to exactly these names.

Folder layout
-------------
recall_project_v2/
    README.txt
    requirements.txt
    step1_extract.py
    step2_analysis.py
    step3_excel.py
    data/
        device-recall-0001-of-0001_json.zip
        device-enforcement-0001-of-0001_json.zip

Run it (Mac: use the Terminal app; Windows: use Command Prompt and type "python" instead of "python3")
-----------------------------------------------------------------------------------
1. Open a terminal in the recall_project folder:
       cd path/to/recall_project
2. Install the libraries (once):
       python3 -m pip install -r requirements.txt
3. Run the three steps in order:
       python3 step1_extract.py     (1-3 min)  -> data/recalls.csv, data/enforcement.csv
       python3 step2_analysis.py    (< 1 min)  -> outputs/ tables, charts, results.json
                                               (includes the logistic regression; needs statsmodels)
       python3 step3_excel.py       (about 30 s) -> outputs/Recall_Analysis_Calculations.xlsx
       python3 step4_excel_ise250.py (about 30 s) -> outputs/Recall_Analysis_ISE250_Style.xlsx
         (same results, laid out like the ISE 250 class material: DMAIC order, Minitab-style output,
          textbook KIND Karz table formats, p chart workbook layout, FMEA SEV x OCC x DET table)
   Open the workbook in Excel. It has about 41,000 formulas, so the first open can take a few seconds.

What you should see after step 2 (data downloaded 27 Oct 2026)
---------------------------------------------------------------
Recall events studied (2013-2025): 13,410 (36,287 recalled products)
Class I events 2020 -> 2025: 37 -> 125 (+238%)
Class I share 3.5% (2016-18) -> 11.5% (2023-25), z = 12.16
Root cause not determined: 20.7% of events
Top priorities: ['Production process', 'Design (hardware)']
Logistic regression: 10,640 events, McFadden R2 = 0.216, AUC = 0.837, root-cause LR test chi2 = 158.8 (df 6)
  (then a table of odds ratios; Design (hardware) = 2.012, Period 2023-2025 = 2.539)
If you download newer FDA files, the numbers will change slightly because FDA adds new recalls weekly.

Outputs
-------
outputs/chart1_class1_by_year.png     the problem: Class I recalls rising
outputs/chart2_pareto.png             Pareto of root causes
outputs/chart3_severity.png           Class I rate by root cause
outputs/chart4_pchart.png             p-chart of Class I share by quarter
outputs/chart5_priority.png           risk priority number ranking
outputs/chart6_logistic_odds_ratios.png  logistic regression odds ratios (adjusted)
outputs/table_logistic_regression.csv every coefficient, standard error, p-value and odds ratio
outputs/table_*.csv                   every result table (open in Excel)
outputs/study_events.csv              the 13,410 events analysed
outputs/results.json                  every number used in the report
outputs/Recall_Analysis_Calculations.xlsx   Excel workbook with formulas

Common problems
---------------
"No module named ..."            -> run step 2 of "Run it" again (pip install).
"No such file or directory: data/..." -> check the zip files are inside the data folder with the exact names.
"python3: command not found" (Windows) -> use "python" instead of "python3".

Documents
---------
docs/LINKS.txt   links to the final report, the from-scratch explanation, the defense speech and the implementation guide

What changed in version 2
-------------------------
- step2_analysis.py section 12b: logistic regression (Class I ~ root cause + device class + specialty + period),
  with firm-clustered standard errors, likelihood-ratio tests, McFadden R2 and AUC. New chart 6 and table_logistic_regression.csv.
- study_events.csv has three new columns: devclass, period, spec_model.
- step3_excel.py: new sheet "9 Logistic regression" (odds ratios, model checks, unadjusted vs adjusted, what-if calculator);
  Events sheet columns N-P compute each event's predicted probability and log-likelihood with formulas.
- requirements.txt adds statsmodels. Run the pip install line again.
