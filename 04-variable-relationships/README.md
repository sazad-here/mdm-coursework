# Variable Relationships

**Chapter 3** — Relationships among variables (Albright 8e).

## Textbook datasets (`data/`)

| File | Problems | Textbook reference |
|------|----------|-------------------|
| `P03_02.xlsx` | 2, 34 | Ch. 3, pp. 90 & 125 |
| `P03_08.xlsx` | 8 | Ch. 3, p. 96 |
| `P03_21.xlsx` | 21 | Ch. 3, p. 104 |

`P03_08` and `P03_21` share the same employee salary data (first seven columns). Start here before building COUNTIFS tables, pivot tables, and correlation outputs.

## Problems covered

| Problem | Deliverable |
|---------|-------------|
| 2 & 34 | COUNTIFS crosstabs + column charts; pivot table row/column % |
| 8 | Excel Table, filters, summary stats by group |
| 21 | Correlation matrix, scatterplots, strongest predictor |

## Scripts

- **`solve_assignment4.py`** — Builds `outputs/` workbooks from the `data/` files (requires Windows Excel via COM).
- **`_verify_assignment.py`** — Checks workbook structure and COUNTIFS formulas.

```bash
python solve_assignment4.py
python _verify_assignment.py
```

## Sample submissions (`outputs/`)

Completed workbooks and `screenshots/` for crosstabs, pivots, and correlation.

## Skills

COUNTIFS, pivot tables, correlation, Excel Tables, Python/openpyxl automation
