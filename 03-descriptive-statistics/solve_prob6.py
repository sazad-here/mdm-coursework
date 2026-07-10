import pandas as pd
import xlsxwriter
import os
from pathlib import Path

DATA = Path(__file__).resolve().parent / "data"

df3 = pd.read_excel(DATA / "P02_03.xlsx", sheet_name="Data")
df6 = pd.read_excel(DATA / "P02_06.xlsx", sheet_name="Data")
df20 = pd.read_excel(DATA / "P02_20.xlsx", sheet_name="Data")
df36 = pd.read_excel(DATA / "Supermarket_Transactions.xlsx", sheet_name="Data")

output_file = 'Assignment_3_Answers.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

# --- Add formats ---
header_format = workbook.add_format({
    'bg_color': '#203764', 'font_color': 'white', 'bold': True, 'font_size': 11, 'font_name': 'Calibri', 'valign': 'vcenter'
})
subheader_format = workbook.add_format({
    'bg_color': '#D9E1F2', 'font_color': 'black', 'bold': True, 'font_size': 11, 'font_name': 'Calibri', 'bottom': 1
})
normal_format = workbook.add_format({
    'font_size': 11, 'font_name': 'Calibri', 'valign': 'vcenter', 'text_wrap': True
})
gray_format = workbook.add_format({
    'bg_color': '#F2F2F2', 'font_size': 11, 'font_name': 'Calibri', 'valign': 'vcenter', 'text_wrap': True
})
percent_format = workbook.add_format({
    'font_size': 11, 'font_name': 'Calibri', 'valign': 'vcenter', 'num_format': '0.00%'
})

# =========================================================
# PROBLEM 3
# =========================================================
worksheet_data3 = workbook.add_worksheet('Problem 3 Data')

columns3 = list(df3.columns) + ['Gender Recoded', 'Children Recoded', 'Salary Recoded', 'Opinion Recoded']
for col_num, value in enumerate(columns3):
    worksheet_data3.write(0, col_num, value, subheader_format)

for row_num in range(len(df3)):
    excel_row = row_num + 2
    worksheet_data3.write(row_num + 1, 0, df3.iloc[row_num]['Person'], normal_format)
    worksheet_data3.write(row_num + 1, 1, df3.iloc[row_num]['Age'], normal_format)
    worksheet_data3.write(row_num + 1, 2, df3.iloc[row_num]['Gender'], normal_format)
    worksheet_data3.write(row_num + 1, 3, df3.iloc[row_num]['State'], normal_format)
    worksheet_data3.write(row_num + 1, 4, df3.iloc[row_num]['Children'], normal_format)
    worksheet_data3.write(row_num + 1, 5, df3.iloc[row_num]['Salary'], normal_format)
    worksheet_data3.write(row_num + 1, 6, df3.iloc[row_num]['Opinion'], normal_format)
    
    worksheet_data3.write_formula(row_num + 1, 7, f'=IF(C{excel_row}=1, "Male", "Female")', normal_format)
    worksheet_data3.write_formula(row_num + 1, 8, f'=IF(E{excel_row}=0, "No children", "At least one child")', normal_format)
    worksheet_data3.write_formula(row_num + 1, 9, f'=IF(F{excel_row}<40000, "Less than 40K", IF(F{excel_row}<=70000, "Between 40K and 70K", IF(F{excel_row}<=100000, "Between 70K and 100K", "Greater than 100K")))', normal_format)
    worksheet_data3.write_formula(row_num + 1, 10, f'=IF(G{excel_row}="Strongly disagree", 1, IF(G{excel_row}="Disagree", 2, IF(G{excel_row}="Neutral", 3, IF(G{excel_row}="Agree", 4, IF(G{excel_row}="Strongly agree", 5, 0)))))', normal_format)

worksheet_data3.set_column('A:K', 15)

ws3 = workbook.add_worksheet('Problem 3 Answers')
ws3.merge_range('A1:I1', 'Problem 3', header_format)

# Part a
ws3.write('A3', 'a)', subheader_format)
ws3.write('B3', '', subheader_format)
ws3.write('A4', 'Name categorical variables', gray_format)
ws3.write('B4', 'Age, Gender, State, Opinion', gray_format)
ws3.write('A5', 'Name nominal:', normal_format)
ws3.write('B5', 'Gender, State', normal_format)
ws3.write('A6', 'Name ordinal:', gray_format)
ws3.write('B6', 'Age, Opinion', gray_format)

# Part b
ws3.write('A9', 'b)', subheader_format)
ws3.write('B9', '', subheader_format)
ws3.write('A10', 'Column Chart', gray_format)

ws3.write('A12', 'Age Group', subheader_format)
ws3.write('B12', 'Frequency', subheader_format)
ages = ['Young', 'Middle-aged', 'Elderly']
for i, age in enumerate(ages):
    row = 13 + i
    fmt = gray_format if i % 2 == 0 else normal_format
    ws3.write(row - 1, 0, age, fmt)
    ws3.write_formula(row - 1, 1, f"=COUNTIF('Problem 3 Data'!B:B, A{row})", fmt, value=len(df3[df3['Age']==age]))
ws3.write(15, 0, 'Total', subheader_format)
ws3.write_formula(15, 1, '=SUM(B13:B15)', subheader_format, value=len(df3))


chart_age = workbook.add_chart({'type': 'column'})
chart_age.add_series({
    'name': 'Frequency',
    'categories': "='Problem 3 Answers'!$A$13:$A$15",
    'values': "='Problem 3 Answers'!$B$13:$B$15",
    'fill': {'color': '#4472C4'},
    'gap': 150
})
chart_age.set_title({
    'name': 'Column Chart (Age Distribution)',
    'name_font': {'bold': True, 'size': 18}
})
chart_age.set_x_axis({
    'name': 'Age Group',
    'name_font': {'bold': True},
    'major_gridlines': {'visible': False}
})
chart_age.set_y_axis({
    'name': 'Count',
    'name_font': {'bold': True},
    'major_gridlines': {'visible': True},
    'major_unit': 50,
    'max': 250,
    'min': 0
})
chart_age.set_legend({'none': True})
ws3.insert_chart('E10', chart_age, {'x_scale': 1.2, 'y_scale': 1.1})


# Gender
ws3.write('A17', 'Gender', subheader_format)
ws3.write('B17', 'Frequency', subheader_format)
ws3.write('A18', 'Male', gray_format)
ws3.write_formula('B18', "=COUNTIF('Problem 3 Data'!C:C, 1)", gray_format, value=len(df3[df3['Gender']==1]))
ws3.write('A19', 'Female', normal_format)
ws3.write_formula('B19', "=COUNTIF('Problem 3 Data'!C:C, 2)", normal_format, value=len(df3[df3['Gender']==2]))
ws3.write(19, 0, 'Total', subheader_format)
ws3.write_formula(19, 1, '=SUM(B18:B19)', subheader_format, value=len(df3))

chart_gender = workbook.add_chart({'type': 'column'})
chart_gender.add_series({
    'name': 'Frequency',
    'categories': "='Problem 3 Answers'!$A$18:$A$19",
    'values': "='Problem 3 Answers'!$B$18:$B$19",
    'fill': {'color': '#ED7D31'}
})
chart_gender.set_title({'name': 'Column Chart (Gender)'})
chart_gender.set_x_axis({'name': 'Gender Code'})
chart_gender.set_y_axis({'name': 'Count'})
chart_gender.set_legend({'none': True})
ws3.insert_chart('E25', chart_gender, {'x_scale': 1.2, 'y_scale': 1.1})

# State
ws3.write('A22', 'State', subheader_format)
ws3.write('B22', 'Frequency', subheader_format)
states = ['Texas', 'Virginia', 'California', 'Michigan', 'Illinois', 'Arizona', 'Minnesota', 'Florida', 'New York', 'Ohio']
for i, s in enumerate(states):
    row = 23 + i
    fmt = gray_format if i % 2 == 0 else normal_format
    ws3.write(row - 1, 0, s, fmt)
    ws3.write_formula(row - 1, 1, f"=COUNTIF('Problem 3 Data'!D:D, A{row})", fmt, value=len(df3[df3['State']==s]))
ws3.write(32, 0, 'Total', subheader_format)
ws3.write_formula(32, 1, '=SUM(B23:B32)', subheader_format, value=len(df3))

chart_state = workbook.add_chart({'type': 'column'})
chart_state.add_series({
    'name': 'Frequency',
    'categories': "='Problem 3 Answers'!$A$23:$A$32",
    'values': "='Problem 3 Answers'!$B$23:$B$32",
    'fill': {'color': '#A5A5A5'}
})
chart_state.set_title({'name': 'Column Chart (State)'})
chart_state.set_x_axis({'name': 'State'})
chart_state.set_y_axis({'name': 'Count'})
chart_state.set_legend({'none': True})
ws3.insert_chart('E40', chart_state, {'x_scale': 1.2, 'y_scale': 1.1})

# Opinion
ws3.write('A35', 'Opinion', subheader_format)
ws3.write('B35', 'Frequency', subheader_format)
opinions = ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree']
for i, o in enumerate(opinions):
    row = 36 + i
    fmt = gray_format if i % 2 == 0 else normal_format
    ws3.write(row - 1, 0, o, fmt)
    ws3.write_formula(row - 1, 1, f"=COUNTIF('Problem 3 Data'!G:G, A{row})", fmt, value=len(df3[df3['Opinion']==o]))
ws3.write(40, 0, 'Total', subheader_format)
ws3.write_formula(40, 1, '=SUM(B36:B40)', subheader_format, value=len(df3))

chart_opinion = workbook.add_chart({'type': 'column'})
chart_opinion.add_series({
    'name': 'Frequency',
    'categories': "='Problem 3 Answers'!$A$36:$A$40",
    'values': "='Problem 3 Answers'!$B$36:$B$40",
    'fill': {'color': '#FFC000'}
})
chart_opinion.set_title({'name': 'Column Chart (Opinion)'})
chart_opinion.set_x_axis({'name': 'Opinion'})
chart_opinion.set_y_axis({'name': 'Count'})
chart_opinion.set_legend({'none': True})
ws3.insert_chart('E55', chart_opinion, {'x_scale': 1.2, 'y_scale': 1.1})

# Part c
ws3.write('A90', 'c) Recoding Summary', subheader_format)
ws3.write('B90', '', subheader_format)
ws3.write('A92', 'Gender', subheader_format)
ws3.write('B92', "The gender variable was recoded from numerical codes (1 and 2) to 'Male' and 'Female' to transform the raw identifier into meaningful, human-readable nominal categories. This makes the data immediately interpretable in future pivot tables or charts.", normal_format)
ws3.write('A94', 'Gender', gray_format)
ws3.write('B94', 'Count', gray_format)
ws3.write('A95', 'Female', normal_format)
ws3.write_formula('B95', "=COUNTIF('Problem 3 Data'!H:H, A95)", normal_format)
ws3.write('A96', 'Male', gray_format)
ws3.write_formula('B96', "=COUNTIF('Problem 3 Data'!H:H, A96)", gray_format)

ws3.write('A99', 'Children', subheader_format)
ws3.write('B99', "The number of children was recoded into two categories: 'No children' (for 0) and 'At least one child' (for values 1 or greater). This converts a discrete quantitative variable into a binary nominal variable, which is highly useful in marketing analytics to isolate the spending habits of parents vs. non-parents.", normal_format)
ws3.write('A101', 'Children Category', gray_format)
ws3.write('B101', 'Count', gray_format)
ws3.write('A102', 'At least one child', normal_format)
ws3.write_formula('B102', "=COUNTIF('Problem 3 Data'!I:I, A102)", normal_format)
ws3.write('A103', 'No children', gray_format)
ws3.write_formula('B103', "=COUNTIF('Problem 3 Data'!I:I, A103)", gray_format)

ws3.write('A106', 'Salary', subheader_format)
ws3.write('B106', "Salary values were grouped into four distinct categories (bins) using $30K intervals. This transforms a continuous quantitative variable into an ordinal categorical variable. Creating these structured income brackets allows us to easily segment customers into Low, Lower-Middle, Upper-Middle, and High-income tiers for demographic analysis.", normal_format)
ws3.write('A108', 'Salary Category', gray_format)
ws3.write('B108', 'Count', gray_format)
sal_bins = ['Less than 40K', 'Between 40K and 70K', 'Between 70K and 100K', 'Greater than 100K']
for i, b in enumerate(sal_bins):
    row = 109 + i
    fmt = normal_format if i % 2 == 0 else gray_format
    ws3.write(row - 1, 0, b, fmt)
    ws3.write_formula(row - 1, 1, f"=COUNTIF('Problem 3 Data'!J:J, A{row})", fmt)

ws3.write('A114', 'Opinion', subheader_format)
ws3.write('B114', "The opinion responses were converted from text labels to numerical codes ranging from 1 to 5. This translates the qualitative Likert scale into ordinal numeric data, which is mathematically necessary if we want to calculate an 'average opinion' or run regression models on customer satisfaction.", normal_format)
ws3.write('A116', 'Opinion Code', gray_format)
ws3.write('B116', 'Count', gray_format)
op_bins = ['1 (Strongly disagree)', '2 (Disagree)', '3 (Neutral)', '4 (Agree)', '5 (Strongly agree)']
for i, b in enumerate(op_bins):
    row = 117 + i
    fmt = normal_format if i % 2 == 0 else gray_format
    ws3.write(row - 1, 0, b, fmt)
    ws3.write_formula(row - 1, 1, f"=COUNTIF('Problem 3 Data'!K:K, {i+1})", fmt)

ws3.write('A124', 'Salary Category Distribution', subheader_format)
ws3.write('B124', 'The following chart displays the distribution of individuals across the newly defined salary categories.', normal_format)

chart_salary = workbook.add_chart({'type': 'column'})
chart_salary.add_series({
    'name': 'Frequency',
    'categories': "='Problem 3 Answers'!$A$109:$A$112",
    'values': "='Problem 3 Answers'!$B$109:$B$112",
    'fill': {'color': '#ED7D31'},
    'gap': 150
})
chart_salary.set_title({'none': True})
chart_salary.set_x_axis({
    'name': 'Salary Category',
    'name_font': {'bold': True},
    'major_gridlines': {'visible': False}
})
chart_salary.set_y_axis({
    'major_gridlines': {'visible': True},
    'major_unit': 50
})
chart_salary.set_legend({'none': True})
ws3.insert_chart('A126', chart_salary, {'x_scale': 1.5, 'y_scale': 1.5})

ws3.set_column('A:A', 35)
ws3.set_column('B:B', 100)

# =========================================================
# PROBLEM 6
# =========================================================
worksheet_data6 = workbook.add_worksheet('Problem 6 Data')
for col_num, value in enumerate(df6.columns):
    worksheet_data6.write(0, col_num, value, subheader_format)

for row_num in range(len(df6)):
    worksheet_data6.write(row_num + 1, 0, df6.iloc[row_num]['Metropolitan_Area'], normal_format)
    worksheet_data6.write(row_num + 1, 1, df6.iloc[row_num]['Commute Time'], normal_format)

worksheet_data6.set_column('A:B', 30)

ws6 = workbook.add_worksheet('Problem 6 Answers')
ws6.merge_range('A1:G1', 'Problem 6', header_format)

# Part a: Frequency Distribution and Histogram
ws6.write('A3', 'a) Frequency Distribution (Class width = 10) & Histogram', subheader_format)
ws6.write('B3', '', subheader_format)

# Add user's analytical text
ws6.merge_range('A4:D6', "Based on the 'Commute Time' data, a frequency distribution using 5 categories with a class width of 10, starting from 30 minutes. All 379 records fall within the range of 30 to 80 minutes.", normal_format)

ws6.write('A8', 'Commute Time (Bin)', gray_format)
ws6.write('B8', 'Frequency', gray_format)

bins = [
    ('30 to < 40', '>=30', '<40'),
    ('40 to < 50', '>=40', '<50'),
    ('50 to < 60', '>=50', '<60'),
    ('60 to < 70', '>=60', '<70'),
    ('70 to < 80', '>=70', '<80')
]

for i, (label, cond1, cond2) in enumerate(bins):
    row = 9 + i
    fmt = normal_format if i % 2 == 0 else gray_format
    ws6.write(row - 1, 0, label, fmt)
    ws6.write_formula(row - 1, 1, f'=COUNTIFS(\'Problem 6 Data\'!B:B, "{cond1}", \'Problem 6 Data\'!B:B, "{cond2}")', fmt)

ws6.write(14, 0, 'Total', subheader_format)
ws6.write_formula(14, 1, '=SUM(B9:B13)', subheader_format)

ws6.merge_range('A16:D17', 'The majority of commute times fall within the 40-50 minute range, followed by the 50-60 minute range.', normal_format)

# Part 2: Histogram
ws6.write('A20', '2) Histogram', subheader_format)
ws6.write('B20', '', subheader_format)

chart_hist = workbook.add_chart({'type': 'column'})
chart_hist.add_series({
    'name': 'Frequency',
    'categories': "='Problem 6 Answers'!$A$9:$A$13",
    'values': "='Problem 6 Answers'!$B$9:$B$13",
    'fill': {'color': '#5B9BD5'},
    'gap': 150
})
chart_hist.set_title({'name': 'Histogram of Commute Times'})
chart_hist.set_x_axis({'name': 'Commute Time (Minutes)'})
chart_hist.set_y_axis({'name': 'Frequency'})
chart_hist.set_legend({'none': True})
ws6.insert_chart('E5', chart_hist, {'x_scale': 1.2, 'y_scale': 1.1})

# Part b: Representative Measure
ws6.write('A23', 'b) Representative Measure', subheader_format)
ws6.write('B23', '', subheader_format)

analysis_text = """The analysis of the 'Commute Time' data in the 'Data' table reveals the following statistics:
Mean: 48.09
Median: 47.33
Skewness: 0.67

Analysis and Recommendation
The skewness value of 0.67 indicates that the data is moderately right-skewed (positively skewed). In such distributions, the mean is pulled higher by the longer commute times in the upper tail, which is why the mean (48.09) is greater than the median (47.33).

Recommendation:
The median is the most appropriate representative measure for this dataset. Because the data is skewed, the median provides a better "typical" value that is less influenced by the extreme values in the tail of the distribution compared to the mean.

Data Samples
There are 379 records in total. Here are the first 5 records from the table:
1. Abilene, TX: 37.57
2. Akron, OH: 49.38
3. Albany, GA: 44.52
4. Albany-Schenectady-Troy, NY: 48.37
5. Albuquerque, NM: 48.85"""

ws6.merge_range('A25:G32', analysis_text, normal_format)

ws6.write('A34', 'Mean', gray_format)
ws6.write_formula('B34', '=AVERAGE(\'Problem 6 Data\'!B:B)', gray_format)
ws6.write('A35', 'Median', normal_format)
ws6.write_formula('B35', '=MEDIAN(\'Problem 6 Data\'!B:B)', normal_format)
ws6.write('A36', 'Skewness', gray_format)
ws6.write_formula('B36', '=SKEW(\'Problem 6 Data\'!B:B)', gray_format)

# Part c: Measure of Variability
ws6.write('A39', 'c) Measure of Variability', subheader_format)
ws6.write('B39', '', subheader_format)
ws6.merge_range('A41:F41', 'The Standard Deviation and Variance represent the spread of the commute times around the mean.', normal_format)

ws6.write('A43', 'Standard Deviation', gray_format)
ws6.write_formula('B43', '=STDEV(\'Problem 6 Data\'!B:B)', gray_format)
ws6.write('A44', 'Variance', normal_format)
ws6.write_formula('B44', '=VAR(\'Problem 6 Data\'!B:B)', normal_format)
ws6.write('A45', 'Range', gray_format)
ws6.write_formula('B45', '=MAX(\'Problem 6 Data\'!B:B) - MIN(\'Problem 6 Data\'!B:B)', gray_format)

# Part d: Empirical Rule (95%)
ws6.write('A48', 'd) Empirical Rule (95% of travel times)', subheader_format)
ws6.write('B48', '', subheader_format)

emp_text = """Based on the analysis of the 'Commute Time' data in the 'Data' table, here are the calculated statistics and the evaluation of the empirical rule:

Statistical Summary
Mean Commute Time: 48.09
Standard Deviation (SD): 7.51
Empirical Rule Range (Mean +/- 2*SD): 33.08 to 63.11

Empirical Rule Evaluation
The empirical rule predicts that approximately 95% of the data should fall within two standard deviations of the mean.
Actual Percentage within Range: 95.25%
Total Records: 379
Records within Range: 361 (e.g., Abilene, TX: 37.57; Akron, OH: 49.38; Albany, GA: 44.52; Albany-Schenectady-Troy, NY: 48.37)

Conclusion:
The empirical rule is approximately correct for this dataset, as the actual percentage of values falling within the calculated range (95.25%) is very close to the predicted 95%."""

ws6.merge_range('A50:G62', emp_text, normal_format)

# Live Formula Verification block
ws6.write('A64', 'Live Formula Verifications', subheader_format)
ws6.write('B64', '', subheader_format)

ws6.write('A65', 'Mean', gray_format)
ws6.write_formula('B65', '=AVERAGE(\'Problem 6 Data\'!B:B)', gray_format)
ws6.write('A66', 'Standard Deviation', normal_format)
ws6.write_formula('B66', '=STDEV(\'Problem 6 Data\'!B:B)', normal_format)

ws6.write('A67', 'Lower Bound (-2 SD)', gray_format)
ws6.write_formula('B67', '=B65 - (2 * B66)', gray_format)
ws6.write('A68', 'Upper Bound (+2 SD)', normal_format)
ws6.write_formula('B68', '=B65 + (2 * B66)', normal_format)

ws6.write('A69', 'Records in Range', gray_format)
ws6.write_formula('B69', '=COUNTIFS(\'Problem 6 Data\'!B:B, ">="&B67, \'Problem 6 Data\'!B:B, "<="&B68)', gray_format)
ws6.write('A70', 'Total Records', normal_format)
ws6.write_formula('B70', '=COUNT(\'Problem 6 Data\'!B:B)', normal_format)
ws6.write('A71', 'Actual Percentage', gray_format)
ws6.write_formula('B71', '=B69/B70', gray_format)

ws6.set_column('A:A', 35)
ws6.set_column('B:B', 30)

# =========================================================
# PROBLEM 20
# =========================================================
worksheet_data20 = workbook.add_worksheet('Problem 20 Data')

# Format purchase date explicitly as a string in python before writing it to excel so it doesn't get messed up.
df36['Purchase Date'] = pd.to_datetime(df36['Purchase Date'], errors='coerce').dt.strftime('%m/%d/%Y').fillna('')

# Write headers
for col_num, value in enumerate(df20.columns):
    worksheet_data20.write(0, col_num, value, subheader_format)
worksheet_data20.write(0, 2, 'Year-to-Year Difference', subheader_format)

for row_num in range(len(df20)):
    excel_row = row_num + 2
    worksheet_data20.write(row_num + 1, 0, df20.iloc[row_num]['Year'], normal_format)
    worksheet_data20.write(row_num + 1, 1, df20.iloc[row_num]['Healthcare expenditure as a percent of GDP'], percent_format)
    
    if row_num == 0:
        worksheet_data20.write(row_num + 1, 2, 'N/A', normal_format)
    else:
        # Multiply by 100 to convert decimal (e.g. 0.0018) to percentage points (0.18) so chart axis plots as standard numbers
        worksheet_data20.write_formula(row_num + 1, 2, f'=(B{excel_row}-B{excel_row-1})*100', normal_format)

worksheet_data20.set_column('A:C', 35)

ws20 = workbook.add_worksheet('Problem 20 Answers')
ws20.merge_range('A1:G1', 'Problem 20', header_format)

# Part a: Time series graph
ws20.write('A3', 'a) Time Series Graph & Trends', subheader_format)
ws20.write('B3', '', subheader_format)

chart_ts = workbook.add_chart({'type': 'line'})
chart_ts.add_series({
    'name': 'Healthcare % of GDP',
    'categories': f"='Problem 20 Data'!$A$2:$A${len(df20)+1}",
    'values': f"='Problem 20 Data'!$B$2:$B${len(df20)+1}",
    'line': {'color': '#4472C4', 'width': 2.25}
})
chart_ts.set_title({'name': 'Healthcare Expenditure as a % of GDP (1960-2020)'})
chart_ts.set_x_axis({
    'name': 'Year',
    'text_axis': True
})
chart_ts.set_y_axis({
    'name': 'Percentage of GDP',
    'min': 0,
    'max': 0.25,
    'major_unit': 0.05,
    'num_format': '0%'
})
chart_ts.set_legend({'none': True})
ws20.insert_chart('A5', chart_ts, {'x_scale': 1.5, 'y_scale': 1.5})

ws20.write('A20', 'Most Striking Feature & Trends:', subheader_format)
trend_text = "The most striking feature is the sharp spike in 2020 (nearly 19.7%), likely reflecting the COVID-19 pandemic's impact.\nOverall, there is a consistent, long-term upward trend, with healthcare spending more than tripling over the 60-year period."
ws20.merge_range('A21:G24', trend_text, normal_format)

# Part b: Year-to-year differences
ws20.write('A26', 'b) Year-to-Year Differences', subheader_format)
ws20.write('B26', '', subheader_format)

diff_text = """Year-over-Year Change Analysis
The graph of these differences reveals the annual shifts in healthcare spending as a share of the GDP:
• Significant Spike in 2020: The most dramatic change in the entire dataset occurred in 2020, where expenditure jumped by 2.10 percentage points—the largest annual increase recorded.
• General Growth: Expenditure typically increases by an average of 0.2 to 0.5 percentage points annually.
• Occasional Stability: There were brief periods where spending remained stable or even slightly decreased, such as the mid-1980s and the early 2010s."""
ws20.merge_range('A28:G33', diff_text, normal_format)

ws20.write('A35', 'Year-over-Year Difference in Healthcare Expenditure (% of GDP)', subheader_format)
ws20.write('B35', 'The following chart visualizes these annual changes:', normal_format)

chart_diff = workbook.add_chart({'type': 'line'})
chart_diff.add_series({
    'name': 'Year-to-Year Difference',
    'categories': f"='Problem 20 Data'!$A$3:$A${len(df20)+1}",
    'values': f"='Problem 20 Data'!$C$3:$C${len(df20)+1}",
    'line': {'color': '#ED7D31', 'width': 2.25}
})
chart_diff.set_title({'name': 'Year-to-Year Difference in Healthcare Expenditure %'})
chart_diff.set_x_axis({
    'name': 'Year',
    'text_axis': True,
    'label_position': 'low',
    'num_font': {'rotation': -90},
    'interval_unit': 2
})
chart_diff.set_y_axis({
    'name': 'Difference from Previous Year (Percentage Points)',
    'min': -0.5,
    'max': 2.5,
    'major_unit': 0.5,
    'num_format': '0.0',
    'major_gridlines': {'visible': True}
})
chart_diff.set_legend({'none': True})
ws20.insert_chart('A37', chart_diff, {'x_scale': 1.5, 'y_scale': 1.5})

ws20.set_column('A:A', 35)
ws20.set_column('B:B', 30)

# =========================================================
# PROBLEM 36
# =========================================================
worksheet_data36 = workbook.add_worksheet('Problem 36 Data')

# Write headers
for col_num, value in enumerate(df36.columns):
    worksheet_data36.write(0, col_num, value, subheader_format)

# Write data
for row_num in range(len(df36)):
    for col_num in range(len(df36.columns)):
        worksheet_data36.write(row_num + 1, col_num, df36.iloc[row_num, col_num], normal_format)

worksheet_data36.autofilter(0, 0, len(df36), len(df36.columns) - 1)

# Highlight format
highlight_format = workbook.add_format({
    'bg_color': '#FFFF00', # Yellow highlight
    'font_size': 11, 'font_name': 'Calibri', 'valign': 'vcenter', 'bold': True
})

# =========================================================
# PROBLEM 36 UNIQUE CUSTOMERS SHEET
# =========================================================
worksheet_unique36 = workbook.add_worksheet('Problem 36 Unique Data')

# Write headers for the unique sheet
worksheet_unique36.write(0, 0, 'Customer ID', subheader_format)
worksheet_unique36.write(0, 1, 'Marital Status', subheader_format)
worksheet_unique36.write(0, 2, 'Homeowner', subheader_format)

# Drop duplicates from the pandas dataframe
df_unique = df36.drop_duplicates(subset=['Customer ID']).reset_index(drop=True)

# Write the unique data
for row_num in range(len(df_unique)):
    worksheet_unique36.write(row_num + 1, 0, df_unique.iloc[row_num]['Customer ID'], normal_format)
    worksheet_unique36.write(row_num + 1, 1, df_unique.iloc[row_num]['Marital Status'], normal_format)
    worksheet_unique36.write(row_num + 1, 2, df_unique.iloc[row_num]['Homeowner'], normal_format)

worksheet_unique36.autofilter(0, 0, len(df_unique), 2)
worksheet_unique36.set_column('A:C', 20)

ws36 = workbook.add_worksheet('Problem 36 Answers')
ws36.merge_range('A1:G1', 'Problem 36: Supermarket Transactions', header_format)

ws36.write('A3', 'Total Transactions (Denominator)', subheader_format)
ws36.write_formula('B3', "=COUNT('Problem 36 Data'!A:A)", subheader_format)

# Part a
ws36.write('A5', 'a) Proportion of married customers', subheader_format)
ws36.write('B5', '', subheader_format)
ws36.write('A6', 'Count of Married (M)', gray_format)
ws36.write_formula('B6', '=COUNTIF(\'Problem 36 Data\'!E:E, "M")', gray_format)
ws36.write('A7', 'Proportion', highlight_format)
ws36.write_formula('B7', '=B6/$B$3', highlight_format)
ws36.write('C7', '<-- (Value between 0 and 1)', normal_format)
text_a = "Examples: There are 6,866 transactions where the customer is married, including transaction IDs: 2, 3, 4, 6, and 8."
ws36.merge_range('A8:C8', text_a, normal_format)

# Part b
ws36.write('A10', 'b) Proportion of customers who do not own a home', subheader_format)
ws36.write('B10', '', subheader_format)
ws36.write('A11', 'Count of Non-Homeowners (N)', gray_format)
ws36.write_formula('B11', '=COUNTIF(\'Problem 36 Data\'!F:F, "N")', gray_format)
ws36.write('A12', 'Proportion', highlight_format)
ws36.write_formula('B12', '=B11/$B$3', highlight_format)
ws36.write('C12', '<-- (Value between 0 and 1)', normal_format)
text_b = "Examples: There are 5,615 transactions where the customer is not a homeowner, including transaction IDs: 3, 10, 11, 13, and 14."
ws36.merge_range('A13:C13', text_b, normal_format)

# Part c
ws36.write('A15', 'c) Proportion of customers with at least one child', subheader_format)
ws36.write('B15', '', subheader_format)
ws36.write('A16', 'Count of Children >= 1', gray_format)
ws36.write_formula('B16', '=COUNTIFS(\'Problem 36 Data\'!G:G, ">=1")', gray_format)
ws36.write('A17', 'Proportion', highlight_format)
ws36.write_formula('B17', '=B16/$B$3', highlight_format)
ws36.write('C17', '<-- (Value between 0 and 1)', normal_format)
text_c = "Examples: There are 12,715 transactions where the customer has at least one child, including transaction IDs: 1, 2, 3, 4, and 5."
ws36.merge_range('A18:C18', text_c, normal_format)

# Part d
ws36.write('A20', 'd) Proportion of UNIQUE single customers who own a home', subheader_format)
ws36.write('B20', '', subheader_format)
ws36.write('A21', 'Note:', gray_format)
ws36.write('B21', 'To check this work, you must first copy Column C (Customer ID), Column E (Marital Status), and Column F (Homeowner) to a new sheet and use the "Remove Duplicates" tool on Customer ID to get the 5,404 unique shoppers. Then apply the filters "S" and "Y".', gray_format)

ws36.write('A22', 'Total Unique Customers', normal_format)
ws36.write_formula('B22', "=COUNT('Problem 36 Unique Data'!A:A)", normal_format)

ws36.write('A23', 'Unique Single (S) Homeowners (Y)', gray_format)
ws36.write_formula('B23', '=COUNTIFS(\'Problem 36 Unique Data\'!B:B, "S", \'Problem 36 Unique Data\'!C:C, "Y")', gray_format) 

ws36.write('A24', 'Proportion', highlight_format)
ws36.write_formula('B24', '=B23/B22', highlight_format)
ws36.write('C24', '<-- (Value between 0 and 1)', normal_format)
text_d = "Examples: There are 1,299 unique customers who are single and own a home, including Customer IDs: 7223, 1900, 9673, 20, and 2670."
ws36.merge_range('A25:C25', text_d, normal_format)

ws36.set_column('A:A', 50)
ws36.set_column('B:B', 20)
ws36.set_column('C:C', 30)

workbook.close()
print("Successfully generated Assignment_3_Answers.xlsx")
