# 📊 Customer Analytics Project

An Excel-based customer analytics project that turns raw customer records into live, formula-driven insights on demographics, revenue, purchase behavior, ratings, satisfaction, and discount impact.

All analysis tables live on the **`Analysis`** sheet, and every figure is calculated with formulas that pull directly from the **`Customer_Data`** sheet, so the tables update automatically when the data changes.

---

## 📁 Workbook Structure

| Sheet | Purpose |
|-------|---------|
| `Customer_Data` | Raw customer-level dataset (source of truth) |
| `Analysis` | All analysis tables (Sections 1–7), formula-driven |

---

## 🗂️ Dataset Overview

**Records:** More than 350 customers  |  **Columns:** 12

| Column | Description | Type |
|--------|-------------|------|
| `Customer ID` | Unique customer identifier | Integer |
| `Gender` | Customer gender (Male / Female) | Text |
| `Age` | Customer age in years | Integer |
| `City` | City of residence | Text |
| `Membership Type` | Membership tier (Gold / Silver / Bronze) | Text |
| `Total Spend` | Total amount spent by the customer | Decimal |
| `Items Purchased` | Number of items purchased | Integer |
| `Average Rating` | Average rating given by the customer (out of 5) | Decimal |
| `Discount Applied` | Whether a discount was used (TRUE / FALSE) | Boolean |
| `Days Since Last Purchase` | Days elapsed since the last purchase | Integer |
| `Satisfaction Level` | Satisfied / Neutral / Unsatisfied | Text |

### Sample Data

| Customer ID | Gender | Age | City | Membership Type | Total Spend | Items Purchased | Average Rating | Discount Applied | Days Since Last Purchase | Satisfaction Level |
|---|---|---|---|---|---|---|---|---|---|---|
| 101 | Female | 29 | New York | Gold | 1120.20 | 14 | 4.6 | TRUE | 25 | Satisfied |
| 102 | Male | 34 | Los Angeles | Silver | 780.50 | 11 | 4.1 | FALSE | 18 | Neutral |
| 103 | Female | 43 | Chicago | Bronze | 510.75 | 9 | 3.4 | TRUE | 42 | Unsatisfied |
| 104 | Male | 30 | San Francisco | Gold | 1480.30 | 19 | 4.7 | FALSE | 12 | Satisfied |
| 105 | Male | 27 | Miami | Silver | 720.40 | 13 | 4.0 | TRUE | 55 | Unsatisfied |
| 106 | Female | 37 | Houston | Bronze | 440.80 | 8 | 3.1 | FALSE | 22 | Neutral |
| 107 | Female | 31 | New York | Gold | 1150.60 | 15 | 4.5 | TRUE | 28 | Satisfied |
| 108 | Male | 35 | Los Angeles | Silver | 800.90 | 12 | 4.2 | FALSE | 14 | Neutral |
| 109 | Female | 41 | Chicago | Bronze | 495.25 | 10 | 3.6 | TRUE | 40 | Unsatisfied |
| 110 | Male | 28 | San Francisco | Gold | 1520.10 | 21 | 4.8 | FALSE | 9 | Satisfied |

---

## 🎯 Project Objectives

### Section 1 — Customer Demographics Analysis
- 1.1 Customer Count by Gender
- 1.2 Customer Count by Membership Type
- 1.3 Customer Count by City
- 1.4 Average Age by Gender
- 1.5 Average Age by Membership Type

### Section 2 — Revenue & Spending Analysis
- 2.1 Average Total Spend by Gender
- 2.2 Average Total Spend by Membership Type
- 2.3 Average Total Spend by City
- 2.4 Total Revenue by Membership Type
- 2.5 Total Revenue by City
- 2.6 Total Revenue by Gender and Membership Type
- 2.7 Average Spend by Gender and Membership Type

### Section 3 — Purchase Behavior Analysis
- 3.1 Average Items Purchased by Membership Type
- 3.2 Average Items Purchased by Gender
- 3.3 Average Days Since Last Purchase by Membership Type
- 3.4 Average Days Since Last Purchase by City

### Section 4 — Customer Rating Analysis
- 4.1 Average Rating by Membership Type
- 4.2 Average Rating by City
- 4.3 Average Rating by City and Membership Type

### Section 5 — Satisfaction Analysis
- 5.1 Customer Count by Satisfaction Level
- 5.2 Average Spend by Satisfaction Level
- 5.3 Average Rating by Satisfaction Level
- 5.4 Average Items Purchased by Satisfaction Level
- 5.5 Total Spend by Satisfaction Level and Membership Type

### Section 6 — Discount Impact Analysis
- 6.1 Customer Count by Discount Applied (TRUE/FALSE)
- 6.2 Average Spend — With vs. Without Discount
- 6.3 Average Rating — With vs. Without Discount
- 6.4 Average Items Purchased — With vs. Without Discount

### Section 7 — Combined Business Analysis
- 7.1 Customer Count by City and Membership Type
- 7.2 Customer Count, Total Revenue, Avg Spend & Avg Rating by Membership

---

## 🧮 Formulas Used

The analysis relies on standard Excel aggregate functions referencing `Customer_Data`.

| Purpose | Function | Example |
|---------|----------|---------|
| Count by category | `COUNTIF` | `=COUNTIF(Customer_Data!B:B, "Female")` |
| Count by two categories | `COUNTIFS` | `=COUNTIFS(Customer_Data!D:D, "Chicago", Customer_Data!E:E, "Bronze")` |
| Average by category | `AVERAGEIF` | `=AVERAGEIF(Customer_Data!E:E, "Gold", Customer_Data!F:F)` |
| Average by two categories | `AVERAGEIFS` | `=AVERAGEIFS(Customer_Data!F:F, Customer_Data!B:B, "Male", Customer_Data!E:E, "Silver")` |
| Total by category | `SUMIF` | `=SUMIF(Customer_Data!D:D, "New York", Customer_Data!F:F)` |
| Total by two categories | `SUMIFS` | `=SUMIFS(Customer_Data!F:F, Customer_Data!L:L, "Satisfied", Customer_Data!E:E, "Gold")` |
| Discount comparison | `AVERAGEIF` | `=AVERAGEIF(Customer_Data!I:I, TRUE, Customer_Data!F:F)` |

> **Column reference:** A = Customer ID, B = Gender, C = Age, D = City, E = Membership Type, F = Total Spend, G = Items Purchased, H = Average Rating, I = Discount Applied, J = Column1, K = Days Since Last Purchase, L = Satisfaction Level.

---

## 🚀 Getting Started

1. **Clone the repository**
   ```bash
   git clone https://github.com/lachuayyadhurai3/Excel_Practice.git
   cd https://github.com/lachuayyadhurai3/Excel_Practice.git
   ```
2. **Open the workbook** in Microsoft Excel (or Google Sheets / LibreOffice Calc).
3. **Go to the `Analysis` sheet** to view all tables.
4. **Update `Customer_Data`**: add or edit rows and the analysis tables recalculate automatically.

---

## 🛠️ Tools & Skills Demonstrated

- Microsoft Excel
- Conditional aggregation (`COUNTIFS`, `SUMIFS`, `AVERAGEIFS`)
- Cross-sheet formula referencing
- Segmentation analysis (gender, city, membership, satisfaction)
- Business insight reporting

---

## 📌 Key Questions Answered

- Which membership tier and city drive the most revenue?
- How do spending, ratings, and purchase frequency differ by gender and membership?
- How does satisfaction relate to spend, rating, and items purchased?
- Do discounts lead to higher spend, better ratings, or more items per customer?
- Which customer segments are at risk (high days since last purchase)?

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

---

## 🤝 Contributing

Suggestions and improvements are welcome. Open an issue or submit a pull request.
