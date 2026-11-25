# Mutual Fund Overview & Insights

## About the Project
This project analyses mutual fund data using Python. It performs data cleaning, exploratory data analysis, fund scoring, ranking, and generates a Top 30 mutual fund shortlist.

## Project Objectives
- Clean and preprocess mutual fund data.
- Analyse historical returns and expenses.
- Understand fund performance using statistical methods.
- Calculate weighted scores for mutual funds.
- Rank funds and shortlist the Top 30.
- Generate reports and visualisations.

## Technologies Used
- Python
- Pandas
- NumPy
- Matplotlib
- Seaborn
- Scikit-learn
- OpenPyXL

## Project Workflow
1. Load mutual fund data from CSV.
2. Clean missing values and text data.
3. Perform exploratory data analysis.
4. Calculate statistical summaries.
5. Normalise returns and expense ratios.
6. Calculate weighted fund scores.
7. Rank mutual funds.
8. Export the Top 30 funds to Excel.
9. Generate insights and charts.

## Scoring Methodology

| Parameter | Weight |
|---|---:|
| 3-Year Returns | 45% |
| Expense Ratio | 25% |
| Fund Age | 15% |
| Consistency | 15% |

The final score is calculated using a weighted scoring approach.

## Installation

Clone the repository:

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
```

Install the required libraries:

```bash
pip install -r requirements.txt
```

## Run the Project

```bash
python mutual_fund_analysis.py --data comprehensive_mutual_funds_data.csv --out outputs
```

## Generated Outputs
- `clean_funds.csv` – Cleaned dataset.
- `Top30_Funds.xlsx` – Top 30 ranked funds.
- `insights.txt` – Analysis and findings.
- PNG charts – Statistical visualisations.

## Disclaimer
This project is for educational and analytical purposes only. Historical returns and calculated rankings do not guarantee future investment performance. This project does not provide financial advice.

## Author
Created as a Python Data Analysis project.
