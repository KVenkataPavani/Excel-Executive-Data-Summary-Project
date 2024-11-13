# Quarter One Report - Excel Workbook

## Project Overview

This project involves the analysis and reformatting of a Microsoft Excel workbook called *Quarter One Report.xlsx*. The workbook contains sales data for a series of products across two years, 2022 and 2023, with key information on product prices, quantities sold, and sales totals. The goal of this project was to organize and format the data to create a clear, professional summary that allows for easy analysis and reporting of sales performance across both years, with a specific focus on the first quarter.

## Key Steps in the Project

### 1. Reformatting and Organizing Data
- **Column Adjustments**: The first task was to adjust column widths to ensure data (such as month names) was displayed clearly. A new column was inserted to include "Product ID", and headings were added to enhance clarity and organization.
- **Heading Formatting**: Bold text, background fill colors, and centered headings were applied to key cells, improving the readability and impact of the worksheet's structure. The *Format Painter* tool was used to quickly copy formatting across different sections.

### 2. Data Customization
- **Text Case Correction**: The text in column G (product names) was in uppercase, which was corrected using the `PROPER` function in Excel to ensure proper capitalization (e.g., "MOUNTAIN BIKES" changed to "Mountain Bikes").
- **Data Sorting and Hiding**: Data was sorted by order date (oldest to newest), and irrelevant columns were hidden to keep the worksheet focused on the required data.
- **Freezing Panes**: To make data navigation easier during analysis, certain rows and columns (headers) were frozen, ensuring they remain visible as the user scrolls.

### 3. Using Excel Formulas for Data Analysis
- **Extracting Date Components**: The `MONTH` and `YEAR` functions were used to separate the date values into month and year columns (e.g., extracting "January" as month 1, and the year 2022).
- **Calculating Sales and Tax**:
  - The formula `=N2*O2` was used to calculate total sales for each product (Retail Price * Quantity Sold).
  - The `IF` function, `=IF(P2>2000, P2*5%, 0)`, was used to apply tax (5%) on sales amounts exceeding $2,000, where tax is calculated as 5% of the sales value.
- **Summing Sales by Year**: The `SUMIF` function was used to sum sales based on specific criteria (such as year or month), allowing for easy comparison between the two years:
  - `=SUMIF(L2:L246, 2022, R2:R246)` summed sales for 2022.
  - `=SUMIF(L2:L246, 2023, R2:R246)` summed sales for 2023.
- **Month-by-Month Sales Totals**: Monthly totals were calculated using similar `SUMIF` formulas with the month criteria (e.g., January = 1, February = 2) to break down sales by month.

### 4. Profit Margin Comparison Between Years
- **Percentage Difference**: A formula was created to calculate the percentage difference in sales between 2022 and 2023. The formula `(C6 - B6) / B6` was used to calculate the percentage increase or decrease in sales between the two years, which was then formatted as a percentage.
- The same approach was applied for monthly comparisons (e.g., January 2022 vs. January 2023), helping to visualize trends and growth.

## Formulas Used in the Workbook

- **PROPER Function**:  
  `=PROPER(G2)`  
  Converts text in cell G2 to proper case (capitalizing the first letter of each word).
  
- **MONTH Function**:  
  `=MONTH(J2)`  
  Extracts the month number from a date in cell J2.

- **YEAR Function**:  
  `=YEAR(J2)`  
  Extracts the year from a date in cell J2.

- **Multiplication Formula**:  
  `=N2 * O2`  
  Calculates total sales by multiplying retail price (N2) by quantity sold (O2).

- **IF Function for Tax Calculation**:  
  `=IF(P2 > 2000, P2 * 5%, 0)`  
  Applies a 5% tax if the total sales (P2) exceed $2,000, otherwise returns 0.

- **SUMIF Function** (for summing based on criteria):  
  `=SUMIF(L2:L246, 2022, R2:R246)`  
  Sums sales (R2:R246) for entries with the year 2022 (L2:L246).  
  `=SUMIF($K$2:$K$103, 1, $R$2:$R$103)`  
  Sums sales for January 2022 (Month = 1).

- **Percentage Difference Formula**:  
  `=(C6 - B6) / B6`  
  Calculates the percentage increase in sales from 2022 (B6) to 2023 (C6).

## Conclusion

This project demonstrates effective use of Excel’s data management, analysis, and formula features to manipulate, organize, and summarize sales data for the first quarter of two years. Through the use of sorting, hiding irrelevant columns, applying formulas like `SUMIF`, `IF`, and `PROPER`, and formatting tools like the *Format Painter*, the report was successfully restructured for clear and efficient analysis. The final workbook provides an insightful overview of the sales trends and comparisons across both years, with an emphasis on quarterly performance, profit margins, and tax calculations.
