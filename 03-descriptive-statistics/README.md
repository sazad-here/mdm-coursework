# Descriptive Statistics

**Chapters 1–2** — Describing distributions and categorical data (Albright 8e).

## How to follow this assignment

| Problem | Dataset | What to do |
|---------|---------|------------|
| 3 | `data/P02_03.xlsx` | Name categorical / nominal / ordinal variables; build column charts; recode Gender, Children, Salary, Opinion |
| 6 | `data/P02_06.xlsx` | Frequency table (5 classes, width 10); histogram with labels; choose mean vs median from shape |
| 20 | `data/P02_20.xlsx` | Time-series chart; year-over-year change |
| 36 | `data/Supermarket_Transactions.xlsx` | Summarize a large transaction file (counts, categories, charts) |

Textbook pages: #3 p.50 · #6 p.63 · #20 p.70 · #36 p.78

## Script (optional automation)

```bash
python solve_prob6.py
```

Reads `data/` and writes a formatted answers workbook with **pandas** + **xlsxwriter**.

## Sample work

- `sample-work/Problem_3_Answers.xlsx` — example formatted submission

## Visualizations

### Problem 3 — Categorical variables & column charts

![Problem 3 part a](screenshots/3a.png)

![Problem 3 gender column chart](screenshots/3b%20column%20chart%20gender.png)

![Problem 3 age column chart](screenshots/3b%20column%20chart%20age.png)

![Problem 3 opinion column chart](screenshots/3b%20column%20chart%20oopinion.png)

![Problem 3 state column chart](screenshots/3b%20column%20chart%20state.png)

![Problem 3 salary recode chart](screenshots/3c%20salary.png)

### Problem 6 — Frequency distribution & histogram

![Problem 6 frequency distribution](screenshots/6a.png)

![Problem 6 histogram part b](screenshots/6b.png)

![Problem 6 histogram part c](screenshots/6c.png)

![Problem 6 mean vs median](screenshots/6d.png)

### Problem 20 — Time series

![Problem 20 time series](screenshots/20.a.png)

![Problem 20 year-over-year](screenshots/20.b.png)

### Problem 36 — Supermarket transactions

![Problem 36 part a](screenshots/36.a.png)

![Problem 36 part b](screenshots/36.b.png)

![Problem 36 part c](screenshots/36.c.png)

![Problem 36 part d](screenshots/36.d.png)

## Skills

Frequency tables, histograms, empirical rule, categorical recoding, Python Excel formatting
