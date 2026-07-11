# MDM Coursework

**Managerial Decision Making** — Business Analytics coursework, Lamar University (MS Management Information Systems).

**Author:** [Shakila Azad](https://github.com/shakilaazad08-bot)  
**LinkedIn:** [linkedin.com/in/shakila-azad2005](https://www.linkedin.com/in/shakila-azad2005/)  
**Course:** MDM (Managerial Decision Making)  
**Textbook:** *Business Analytics: Data Analysis and Decision Making* (Albright et al., 8e)

Curated assignments and team case study from the MDM course. Each folder includes **textbook datasets** (`data/`) so students using Albright 8e can download the same files and practice alongside the sample work.

Full rubrics, textbook PDFs, and quizzes are excluded.

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
| 03-descriptive-statistics | P02_03, P02_06, P02_20, Supermarket_Transactions |
| 04-variable-relationships | P03_02, P03_08, P03_21 |
| 05-linear-programming | Production Mix 1 & 2, Production Planning |
| 06-optimization-solver | Worker Scheduling, Transportation, Investing |
| case-study-retirement-investment | C09_04 |

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
