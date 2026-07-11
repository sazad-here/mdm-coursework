# MDM Coursework Portfolio

**Managerial Decision Making (MDM)** — Business Analytics coursework  
Lamar University · MS Management Information Systems

**Author:** [Shakila Azad](https://github.com/sazad-here) · [@sazad-here](https://github.com/sazad-here)  
**LinkedIn:** [linkedin.com/in/shakila-azad2005](https://www.linkedin.com/in/shakila-azad2005/)  
**GitHub repo:** [github.com/sazad-here/mdm-coursework](https://github.com/sazad-here/mdm-coursework)  
**Course:** MDM (Managerial Decision Making)  
**Textbook:** *Business Analytics: Data Analysis and Decision Making* (Albright et al., 8e)

## Project summary (LinkedIn / portfolio)

Curated Managerial Decision Making assignments: regression, descriptive statistics, Solver optimization, decision analysis, and a retirement-investment case study. Includes datasets, visualizations, step-by-step guides, and Python automation for Excel. Built for MS MIS coursework so students and professionals can follow the analysis end to end.

**Skills demonstrated:** Excel · Data Analysis ToolPak · Excel Solver · Regression · Descriptive statistics · Pivot tables / COUNTIFS · Linear programming · Decision analysis · Python (pandas, openpyxl, xlsxwriter)

Full rubrics, textbook PDFs, and quizzes are excluded.

## Start here (if you came from LinkedIn)

| Want to see… | Go to |
|--------------|--------|
| Best showcase (case study + charts) | [case-study-retirement-investment](case-study-retirement-investment/) · open [`index.html`](case-study-retirement-investment/index.html) |
| Python + Excel automation | [04-variable-relationships](04-variable-relationships/) |
| Charts & histograms | [03-descriptive-statistics](03-descriptive-statistics/) |
| Solver / optimization | [05-linear-programming](05-linear-programming/) · [06-optimization-solver](06-optimization-solver/) |
| Full assignment list | See table below |

Connect on LinkedIn: [Shakila Azad](https://www.linkedin.com/in/shakila-azad2005/)

## How to use this repo (students & visitors)

1. Open an assignment folder and read **How to follow this assignment**.
2. Download files from that folder’s **`data/`** directory.
3. Recreate the analysis in Excel (or run the Python scripts where provided).
4. Compare your charts and Solver output to the **Visualizations** section.
5. Use this as a **study reference** — do your own work for graded submissions.

> Academic integrity: Learn from the structure and methods shown here. Do not copy submissions as your own graded work.

## Assignments

| Folder | MDM topic | Tools |
|--------|-----------|--------|
| [01-regression-fundamentals](01-regression-fundamentals/) | Regression basics | Excel |
| [02-regression-inference](02-regression-inference/) | Regression inference | Excel |
| [03-descriptive-statistics](03-descriptive-statistics/) | Descriptive statistics | Excel, Python |
| [04-variable-relationships](04-variable-relationships/) | Variable relationships | Excel, Python |
| [05-linear-programming](05-linear-programming/) | Linear programming | Excel Solver |
| [06-optimization-solver](06-optimization-solver/) | Solver optimization | Excel Solver |
| [07-decision-analysis](07-decision-analysis/) | Decision analysis | Excel |
| [case-study-retirement-investment](case-study-retirement-investment/) | Case study C09_04 | Excel, HTML |

## Datasets by assignment

| Folder | Datasets in `data/` |
|--------|---------------------|
| 01-regression-fundamentals | Bank Salaries, Cost of Power, Drugstore Sales, Overhead Costs |
| 02-regression-inference | Use Albright Ch. 10 companion files (not stored locally) |
| 03-descriptive-statistics | P02_03, P02_06, P02_20, Supermarket_Transactions |
| 04-variable-relationships | P03_02, P03_08, P03_21 |
| 05-linear-programming | Production Mix 1 & 2, Production Planning |
| 06-optimization-solver | Worker Scheduling, Transportation, Investing |
| 07-decision-analysis | Assignment 7 Problem 10 workbook |
| case-study-retirement-investment | C09_04 |

## Visual previews

Each assignment README embeds charts and Solver screenshots.

| Assignment | Preview |
|------------|---------|
| [01-regression-fundamentals](01-regression-fundamentals/) | Regression scatterplot |
| [02-regression-inference](02-regression-inference/) | Coefficient & CI charts |
| [03-descriptive-statistics](03-descriptive-statistics/) | Histograms, column charts, time series |
| [04-variable-relationships](04-variable-relationships/) | Crosstabs, pivots, correlation |
| [05-linear-programming](05-linear-programming/) | Solver optimal solutions |
| [06-optimization-solver](06-optimization-solver/) | Scheduling, transportation, investing |
| [07-decision-analysis](07-decision-analysis/) | Decision equation |
| [case-study-retirement-investment](case-study-retirement-investment/) | Scatterplots + HTML report |

![Case study — salary vs percent invested](case-study-retirement-investment/screenshots/salary.png)

## Highlights

- **`solve_assignment4.py`** — COUNTIFS crosstabs, pivot tables, correlation, charts
- **`solve_prob6.py`** — Descriptive statistics output with pandas and xlsxwriter
- **Case study** — Retirement investment regression (Team 27, interactive HTML)

## Setup

```bash
pip install -r requirements.txt
```

```bash
cd 04-variable-relationships
python solve_assignment4.py
```

```bash
cd 03-descriptive-statistics
python solve_prob6.py
```

## License

Educational portfolio. Dataset copyrights belong to their respective sources and the textbook publisher.
