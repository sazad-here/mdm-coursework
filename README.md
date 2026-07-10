# Business Analytics Portfolio

Coursework from **Managerial Decision Making** (Business Analytics), Lamar University — MS Management Information Systems.

**Author:** [Shakila Azad](https://github.com/shakilaazad08-bot)  
**LinkedIn:** [linkedin.com/in/shakila-azad2005](https://www.linkedin.com/in/shakila-azad2005/)  
**Textbook:** *Business Analytics: Data Analysis and Decision Making* (Albright et al., 8e)

This repository is a **curated portfolio** of assignments and a team case study. It highlights analysis, Excel modeling, and Python automation — not a full dump of course materials (rubrics, textbook PDFs, and quizzes are excluded).

## Projects

| Folder | Topic | Tools |
|--------|--------|--------|
| [01-regression-fundamentals](01-regression-fundamentals/) | Scatterplots, simple & multiple regression | Excel |
| [02-regression-inference](02-regression-inference/) | Significance, confidence intervals, model selection | Excel |
| [03-descriptive-statistics](03-descriptive-statistics/) | Histograms, categorical data, distributions | Excel, Python |
| [04-variable-relationships](04-variable-relationships/) | Crosstabs, pivot tables, correlation | Excel, Python |
| [05-linear-programming](05-linear-programming/) | Product mix & production planning | Excel Solver |
| [06-optimization-solver](06-optimization-solver/) | Scheduling, transportation, investing | Excel Solver |
| [07-decision-analysis](07-decision-analysis/) | Decision analysis (Ch. 7) | Excel |
| [case-study-retirement-investment](case-study-retirement-investment/) | Retirement plan investment regression (C09_04) | Excel, HTML |

## Highlights

- **`solve_assignment4.py`** — Builds Excel workbooks with COUNTIFS crosstabs, pivot tables, correlation sheets, and charts (Windows + Excel).
- **`solve_prob6.py`** — Generates formatted descriptive-statistics output with pandas and xlsxwriter.
- **Case study** — Interactive HTML report on factors affecting tax-deferred retirement contributions.

## Setup (Python scripts)

```bash
pip install -r requirements.txt
```

Run Assignment 4 automation (Excel must be closed):

```bash
cd 04-variable-relationships
python solve_assignment4.py
```

Run Assignment 3 formatter from its `data/` folder context:

```bash
cd 03-descriptive-statistics
python solve_prob6.py
```

## Note on course materials

Publisher slides, full rubrics with model answers, quizzes, and the textbook PDF are intentionally omitted for copyright and academic integrity.

## License

Coursework portfolio for educational and professional demonstration. Dataset copyrights belong to their respective sources and the textbook publisher.
