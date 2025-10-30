# Excel-Macros-VBA-Automation-Projects
A collection of practical Excel VBA macros projects automating data processing, analysis, and reporting workflows - showcasing expertise in automation, dynamic dashboards, conditional logic, and data visualization to turn repetitive tasks into efficient, scalable business solutions; streamlining business analytics and decision-making.
&nbsp;

**OVERVIEW**

This repository showcases multiple **Excel VBA and Macro Automation** projects demonstrating my capability to automate data preparation, formatting, visualization, and reporting processes using Excel’s developer environment.

Each project represents a practical business scenarios - automating repetitive manual operations, improving consistency, and increasing reporting efficiency through VBA scripting and dynamic macros.

The repository highlights how data analysts can leverage Excel’s developer tools for:
- Process optimization
- Automated reporting
- Dynamic data formatting and charting
- Conditional logic and event-driven automation

All project outputs are consolidated as **macro-enabled Excel workbooks** (.xlsm), each demonstrating a distinct VBA concept and automation workflow.
&nbsp;

**SKILLS DEMONSTRATED**

- Excel VBA Programming & Macro Recording
- Data Cleaning, Preparation & Standardization
- Dynamic Range Selection & Automation
- Chart Creation & Visualization Automation
- Conditional Logic & Control Structures (If...Then, Exit Sub)
- Formula Automation using R1C1 Notation
- Conditional Formatting via VBA
- Interactive Button, Image, and Shortcut-Based Macro Execution
- Multi-Sheet Automation & Dynamic Referencing
- Preventive Error Handling & Logic Validation
- Process Documentation, Code Annotation, and Debugging
- Workflow Automation for Business Reporting
  &nbsp;

**PROJECT SUMMARIES**
&nbsp;

**A) Project 1 – Financial Data Formatting** (Macro: “Format”) 

(_File: Macros_Project1.xlsm_)
&nbsp;

**Goal:**

Automate formatting and chart creation for weekly financial datasets.

**Steps:**

1. Used dynamic range selection (Ctrl+Shift+Arrow) to auto-detect data.
2. Applied consistent header formatting and alignment.
3. Automated column chart generation and repositioning.
4. Linked macro to button (“Format”) for one-click execution.
5. Verified by running macro across Week_1, Week_2, Week_3.
6. Added another macro (“Highlighting”) to shade selected rows and linked it to an image button for interactive use.

**Before V/S After Transformation:**

<img width="1920" height="1080" alt="Raw_Data1" src="https://github.com/user-attachments/assets/9ee4523d-dda4-4086-9434-2623ee13b923" />
&nbsp;

<img width="1920" height="1080" alt="Financial_Format_Chart" src="https://github.com/user-attachments/assets/88d8ce92-b9f1-4b1d-a091-044fc6a93bfa" />

**Outcome:**

 A reusable workbook to format, chart, and highlight weekly financial reports in one click.

**Skills:**

Data formatting automation, dynamic range selection, chart creation, macro-linked buttons, and UI integration for quick reporting.

&nbsp;

**B) Project 2 – Customer Data Cleaning** (Macro: “Clean_Data”) 

(_File: Macros_Project2.xlsm_)
&nbsp;

**Goal:**

Clean and structure customer payment data through automation.

**Steps:**

1. Added new column dynamically before “Balance Due”.
2. Renamed headers programmatically.
3. Recorded macro to split full names into first/last using Text to Columns.
4. Applied conditional formatting to flag customers with dues > 0.
5. Linked macro to a “Clean Data” button on a separate sheet.

**Before V/S After Transformation:**

<img width="1920" height="1080" alt="Raw_Data2" src="https://github.com/user-attachments/assets/86223452-a3ed-4b3e-94f1-7bd91ecc9592" />
&nbsp;

<img width="1920" height="1080" alt="CustomerData_Cleaned" src="https://github.com/user-attachments/assets/e6c229d6-6c74-4ad5-902a-5a08d78fe9da" />

**Outcome:**

Automated cleaning process that standardizes name fields and identifies pending balances instantly.

**Skills:**

Text parsing, data cleaning & structuring, conditional formatting via VBA, dynamic column creation, and button-based macro execution.

&nbsp;

**C) Project 3 – Loan Report Automation** (Macro: “Loans_Report”) 

(_File: Macros_Project3.xlsm_)
&nbsp;

**Goal:**

Generate formatted, validated weekly loan reports with computed ratios and highlights.

**Steps:**

1. Applied Calibri font and styled headers (blue background, white bold text).
2. Converted currency columns to Dollar format (Loan Amount, Installment, etc.).
3. Removed redundant columns.
4. Added computed column Debt-to-Income = Installment / (Annual_Income / 12) using formula insertion (R1C1 style).
5. Replaced Delinquent → Charged Off in Loan_Status.
6. Applied conditional formatting (red fill for Charged Off).
7. Auto-fit columns for a clean layout.
8. Added meaningful comments within the VBA editor for clarity.

**Before V/S After Transformation:**

<img width="1920" height="1080" alt="Raw_Data3" src="https://github.com/user-attachments/assets/c80f006a-145c-486d-9b48-a445d8f9c7f1" />
&nbsp;

<img width="1920" height="1080" alt="LoanReport_ConditionalFormatting" src="https://github.com/user-attachments/assets/7be64173-f312-4f3e-be0f-09b10f608d2e" />

**Outcome:**

 Automated weekly loan report standardization with dynamic metric calculation and conditional styling.

**Skills:**

Data transformation, formula automation with R1C1, sorting & replacement, conditional formatting, calculated columns, validation, documentation, and multi-sheet automation.

&nbsp;

**D) Project 4 – Departmental Performance Chart Automation** (Macro: "CreateChart")

(_File: Macros_Project4.xlsm_)
&nbsp;

**Goal:**

Automate visualization generation (combo charts) across multiple departmental sheets using a single reusable macro.

**Sheets:**

Procurement, Finance, Marketing, Sales, Empty Sheet

**Steps:**

1. Formatted cell A1 as department title (bold, larger font).
2. Auto-fit all columns dynamically.
3. Created a combo chart with formatted bars, datalabels, adjust transparency of trendline, custom legend position and background color.
4. Programmed chart title to reference Range("A1").Value for dynamic department labeling.
5. Added conditional logic: If Range("A3").Value <> "Position" Or Range("A4").Value = "" Then Exit Sub Ensuring no chart generation on empty or invalid sheets.
6. Assigned shortcut key Ctrl + C for macro execution.
7. Tested macro across Procurement, Finance, Marketing, and Sales sheets - confirmed auto-title and format updates.
8. Ran on Empty Sheet → No action executed (logic validation successful).

**Before V/S After Transformation:**

<img width="1920" height="1080" alt="Raw_Data4" src="https://github.com/user-attachments/assets/e3b92d32-04bf-4bde-9b22-71d2d60ebeac" />
&nbsp;

<img width="1920" height="1080" alt="Departmental_Chart_Output" src="https://github.com/user-attachments/assets/d430b54f-8af0-4565-849e-68a86945177b" />

**Outcome:**

A fully automated multi-sheet chart generator that ensures department-wise visual consistency and conditional control.

**Skills:**

Visualization automation, conditional VBA logic, code referencing, control structures, preventive error handling, visualization formatting, and shortcut-based execution.

&nbsp;

**DELIVERABLES**

Excel Workbooks:
- Macros_Project1.xlsm
- Macros_Project2.xlsm
- Macros_Project3.xlsm
- Macros_Project4.xlsm

Screenshots (for visual reference):
Embedded within README for both raw and output data views.
&nbsp;

**KEY HIGHLIGHTS**

- Four automation projects showing end-to-end process automation in Excel.
- Combines both recorded macros and manually edited VBA code.
- Demonstrates data cleaning, transformation, chart automation, and error handling.
- Focuses on dynamic referencing, conditional checks, and reusable logic.
- Structured and documented for clarity and reproducibility.
  &nbsp;

**TECH STACK**

| Category | Tools / Skills |
|-----------|----------------|
| **Data Automation** | Excel VBA, Macros |
| **Data Cleaning & Structuring** | Conditional Formatting, Text Parsing |
| **Visualization Automation** | Chart Objects, Dynamic Titles |
| **Logic & Control** | If...Then, Loops, Exit Conditions |
| **Formatting & Styling** | Cell, Range, and Chart Formatting |
| **Documentation** | VBA Comments, Screenshot Documentation |

**OUTCOME**

This project set demonstrates hands-on automation and analytical capability using Excel VBA.
It reflects proficiency in optimizing manual data processes into reusable automated workflows - transforming Excel from a data entry tool into a process automation platform.
&nbsp;

**FUTURE ENHANCEMENTS**

- Integrate macros with Power Query refresh automation.
- Build User Forms for interactive inputs (e.g., dynamic chart range selection).
- Add error handling modules with message boxes for invalid entries.
- Explore cross-file automation for consolidating multiple workbooks.
  &nbsp;

**HOW TO USE**

1. Clone or download this repository.
2. Open any .xlsm file in Excel.
3. Enable macros when prompted.
4. Run macros via:
   - Assigned buttons on sheets, or
   - Keyboard shortcuts (e.g., Ctrl + C).
5. View transformation or chart outputs across sheets.
6. Explore VBA code via Developer → Visual Basic Editor.
   &nbsp;

**REPOSITORY STRUCTURE**
```
Excel_Macros_VBA_Automation_Projects/
│
├── Macros_Project1.xlsm        ➜  Financial Data Formatting
├── Macros_Project2.xlsm        ➜  Customer Data Cleaning
├── Macros_Project3.xlsm        ➜  Loan Report Automation
├── Macros_Project4.xlsm        ➜  Department Chart Automation
│
└── README.md   ➜  Includes all embedded raw & output screenshots
```

**_Designed and documented by [Muskan Tayal](https://www.linkedin.com/in/muskan-tayal-820145225)_**
