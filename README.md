E-Commerce Sales Analysis

Project Overview

This project uses Microsoft Excel to clean, standardize, prepare, and
analyze an e-commerce sales dataset. It combines sales transactions with
customer, product, and store information, then uses Excel formulas,
PivotTables, summary statistics, a dashboard, and a forecast worksheet
to explore sales performance.

**Project type:** Excel Data Analytics Mini Project\
**Tools:** Microsoft Excel, Power Query / Excel data-cleaning tools,
PivotTables, formulas, charts and dashboard\
**Dataset file:** `Ecommerce_Sales_Dataset.xlsx`\
**Analysis period in the workbook:** 14 October 2023 -- 13 October 2025

Objectives

-   Clean inconsistent IDs and text fields.
-   Check missing values and duplicate records.
-   Standardize customer, product, category, country, and store
    information.
-   Prepare transaction data for analysis.
-   Use formulas to calculate and validate useful metrics.
-   Analyze sales by region, store type, category, gender, year, product
    brand, discount, and other dimensions.
-   Present findings through PivotTables, a dashboard, and a forecast
    worksheet.

 Workbook Structure

  -----------------------------------------------------------------------
  Worksheet                           Purpose
  ----------------------------------- -----------------------------------
  `Sales_Fact`                        Main transaction-level dataset,
                                      including sales, customer, product,
                                      store, and sales measures.

  `Customer_Dim`                      Customer attributes such as name,
                                      age, gender, city, state, country,
                                      and loyalty level.

  `Product_Dim`                       Product ID, product brand/name,
                                      category, sub-category, cost, and
                                      stock.

  `Store_Dim`                         Store name, region, city, and store
                                      type.

  `sales_dim`                         A sales-focused transaction table
                                      used for analysis.

  `summary statistics`                Summary-statistics worksheet.

  `Final Dashboard`                   Dashboard presentation sheet.

  `store vs region vs total amt`      Sales amount comparison by
                                      store/region.

  `category vs gender vs total`       Category and gender-related sales
                                      analysis.

  `product brand vs category vs to`   Product/category sales comparison.

  `year vs category,gender vs tota`   Year and category breakdown.

  `category vs total vs dis`          Category sales and average discount
                                      comparison.

  `forecast`                          Historical total amount with
                                      forecast and confidence-bound
                                      columns.
  -----------------------------------------------------------------------

> Some worksheets contain extra `Unnamed` columns that appear to be
> blank or assignment notes. Exclude irrelevant blank/helper columns
> from the final analysis view, while retaining any notes needed to
> document the cleaning process.

Dataset Overview

The workbook contains a transaction table and supporting dimension
tables. The main sales table includes fields such as:

-   **Transaction identifiers:** `Sales_ID`, `Customer_ID`,
    `Product_ID`, `Store_ID`
-   **Date and customer attributes:** `Order_Date`, `Age`, `Gender`,
    `State`, `Loyalty_Level`
-   **Product and store attributes:** `Product_Brand`, `Category`,
    `Region`, `Store_Type`
-   **Sales measures:** `Quantity`, `Discount`, `Unit_Price`, `Cost`,
    `Total_Amount`
-   **Payment field:** `Payment_Type`

The customer, product, and store dimension sheets provide additional
descriptive attributes for grouping and reporting.

 Data Cleaning and Preparation

The following cleaning steps are documented for this project. Apply them
to the relevant columns and verify the result after each transformation.

1. Remove unnecessary columns

-   Review blank `Unnamed` columns and remove them from the
    analysis-ready table when they do not contain useful information.
-   Keep assignment notes separately if they document a transformation
    or business rule.
-   Confirm that the final table has clear, meaningful column headers.

2. Standardize Customer IDs

-   Identify inconsistent formats such as `C-1464`.
-   Use **Find and Replace / Replace Values** to standardize the
    identifier to the agreed format, for example `CUST1464`.
-   Apply one consistent pattern across the entire `Customer_ID` column.
-   Check that customer IDs match between the sales and customer tables.

3. Standardize Product IDs

-   Identify inconsistent formats such as `P-3`.
-   Replace the inconsistent pattern with the agreed standard, for
    example `PRO3`.
-   Check that `Product_ID` values match between the sales and product
    tables.

4. Clean text fields

Use Excel **TRIM**, **CLEAN**, and **PROPER** where appropriate to
remove extra spaces, non-printing characters, and inconsistent
capitalization.

Example helper formula (if the source text is in `B2`):

``` excel
=PROPER(TRIM(CLEAN(B2)))
```

Apply suitable text cleaning to fields such as: - Customer `Name` -
`City` and `State` - `Store_Name` and store `City` - `Product_Brand` /
product name, where appropriate

**Important:** `PROPER` changes capitalization and may alter brand
names, abbreviations, or intentional mixed-case names. Review the
results before replacing the original values.

5. Standardize Country values

-   Replace `United States of America` with `USA` using **Replace
    Values**.
-   Use one consistent country label throughout the dataset.

6. Standardize category and loyalty values

-   Review `Category`, `Sub_Category`, `Loyalty_Level`, `Region`,
    `Store_Type`, `Gender`, and `Payment_Type` for spelling variations,
    extra spaces, inconsistent capitalization, and unexpected labels.
-   Map variations to approved category names.
-   Do not assign missing or `Unknown` loyalty levels without a
    documented business rule.

7. Check missing values

-   Filter each column for blanks.
-   Use formulas such as `COUNTBLANK` to measure missing values.
-   Fill a blank only when the correct value can be derived reliably
    from a trusted source or an approved rule.
-   Do not replace every blank with zero: zero and missing have
    different meanings.
-   The workbook includes some blank fields and helper/notes columns;
    assess them separately from business-data columns.

8. Check duplicate records

-   Use **Data → Remove Duplicates** on the appropriate table and key
    columns.
-   Check duplicate transaction IDs separately from fully duplicated
    rows.
-   Confirm that repeated values are genuinely duplicate records before
    deleting them.

9. Set appropriate data types and formats

-   `Order_Date`: Excel Date format.
-   `Quantity`, `Age`, `Stock`: whole-number format where appropriate.
-   `Discount`: percentage or decimal format according to the documented
    meaning of the field.
-   `Unit_Price`, `Cost`, `Total_Amount`: consistent currency/number
    format.
-   IDs: Text format, to preserve their exact standardized patterns.

10. Validate relationships and calculations

-   Confirm sales `Customer_ID` values exist in `Customer_Dim`.
-   Confirm sales `Product_ID` values exist in `Product_Dim`.
-   Confirm sales `Store_ID` values exist in `Store_Dim`.
-   Check whether transaction values agree with the approved calculation
    rules.
-   Keep the original `Total_Amount` until the definition of `Discount`
    and the amount calculation is confirmed. Do not overwrite source
    values with an assumed formula.

Excel Formulas Used / Recommended

The examples below use ordinary cell references. Adjust the column
letters to match the actual layout, or convert the data range to an
Excel Table and use structured references. These formulas support
cleaning, quality checks, and analysis.

Text cleaning

  -----------------------------------------------------------------------------------------------
  Task                    Formula example                                 Purpose
  ----------------------- ----------------------------------------------- -----------------------
  Remove extra spaces     `=TRIM(B2)`                                     Removes
                                                                          leading/trailing spaces
                                                                          and reduces repeated
                                                                          spaces between words.

  Remove non-printing     `=CLEAN(B2)`                                    Removes common
  characters                                                              non-printing
                                                                          characters.

  Clean and standardize   `=PROPER(TRIM(CLEAN(B2)))`                      Combines text cleaning
  capitalization                                                          with title case; review
                                                                          brand names and
                                                                          abbreviations.

  Standardize country     `=IF(G2="United States of America","USA",G2)`   Maps the long country
  label                                                                   label to `USA`
                                                                          (assuming country is in
                                                                          `G`).

  Check for a blank       `=IF(A2="","Missing","Available")`              Flags missing values in
                                                                          the selected column.
  -----------------------------------------------------------------------------------------------

Missing-value and duplicate checks

  ---------------------------------------------------------------------------------------------------
  Task                    Formula example                                     Purpose
  ----------------------- --------------------------------------------------- -----------------------
  Count blank cells       `=COUNTBLANK(A2:A2002)`                             Counts blanks in a
                                                                              selected range.

  Count populated cells   `=COUNTA(A2:A2002)`                                 Counts non-empty cells.

  Flag duplicate IDs      `=IF(COUNTIF($A$2:A2,A2)>1,"Duplicate","Unique")`   Flags repeated values
                                                                              as the list is
                                                                              evaluated downward.
                                                                              Replace `A` with the ID
                                                                              column.

  Count unique values     `=COUNTA(UNIQUE(A2:A2002))`                         Counts distinct values
  (Microsoft 365)                                                             in a range.

  Flag missing quantity   `=IF(F2="","Missing","Available")`                  Example if `Quantity`
                                                                              is in column F.
  ---------------------------------------------------------------------------------------------------

### Descriptive statistics

  Metric                                         Formula example
  ---------------------------------------------- ----------------------
  Total sales amount                             `=SUM(S2:S2002)`
  Number of sales records with numeric amounts   `=COUNT(S2:S2002)`
  Average sales amount                           `=AVERAGE(S2:S2002)`
  Minimum sales amount                           `=MIN(S2:S2002)`
  Maximum sales amount                           `=MAX(S2:S2002)`
  Total quantity sold                            `=SUM(F2:F2002)`
  Average unit price                             `=AVERAGE(O2:O2002)`
  Count of sales rows                            `=COUNTA(A2:A2002)`

These ranges are examples based on a header row and approximately 2,001
transaction rows. Update the ranges to match the final cleaned table.

### Conditional and category-based analysis

  ----------------------------------------------------------------------------------------------
  Task                    Formula example                                Purpose
  ----------------------- ---------------------------------------------- -----------------------
  Label price band        `=IF(O2>=500,"High Price","Standard Price")`   Classifies unit price
                                                                         using a threshold of
                                                                         500. Change the
                                                                         threshold if the
                                                                         project defines another
                                                                         rule.

  Sum sales for           `=SUMIF(M2:M2002,"Electronics",S2:S2002)`      Adds `Total_Amount` for
  Electronics                                                            Electronics, assuming
                                                                         Category is M and
                                                                         Total_Amount is S.

  Count sales below a     `=COUNTIF(O2:O2002,"<100")`                    Counts unit prices
  price threshold                                                        below 100.

  Sum sales by region     `=SUMIF(Q2:Q2002,"North",S2:S2002)`            Example for North
                                                                         region, assuming Region
                                                                         is Q.

  Sum quantity for a      `=SUMIF(M2:M2002,"Clothing",F2:F2002)`         Totals quantity for
  category                                                               Clothing.

  Average discount by     `=AVERAGEIF(M2:M2002,"Beauty",N2:N2002)`       Averages discount for
  category                                                               Beauty, assuming
                                                                         Discount is N.
  ----------------------------------------------------------------------------------------------

### Extracting ID components

Only use these formulas if the ID format actually contains the component
you intend to extract.

  -----------------------------------------------------------------------
  Task                    Formula example         Purpose
  ----------------------- ----------------------- -----------------------
  First two characters    `=LEFT(A2,2)`           Extracts the first two
                                                  characters.

  Last two characters     `=RIGHT(A2,2)`          Extracts the last two
                                                  characters.

  Middle characters       `=MID(A2,3,2)`          Extracts two characters
                                                  starting at position 3.
  -----------------------------------------------------------------------

For IDs such as `CUST1464` or `PROD97`, these formulas return character
fragments, not necessarily meaningful dates or country codes. Confirm
the ID specification before interpreting extracted text.

### Sales amount validation

The dataset contains a `Discount` column and a pre-existing
`Total_Amount` column. Before calculating a replacement amount, confirm
whether `Discount` represents a percentage, a fixed monetary amount, or
another measure. If the business definition confirms a percentage
discount, a possible formula is:

``` excel
=QuantityCell*UnitPriceCell*(1-DiscountCell)
```

If the discount is a fixed amount per unit, the calculation may instead
be:

``` excel
=QuantityCell*(UnitPriceCell-DiscountCell)
```

These are alternative examples, **not confirmed formulas for this
workbook**. Compare calculated results against the source `Total_Amount`
and document the approved rule before using it in reporting.

## Analysis and PivotTables

The workbook includes PivotTable-style analysis and a dashboard. Useful
analysis questions include:

1.  What is the total sales amount by region and store type?
2.  Which product categories contribute the most sales?
3.  How does sales amount vary by gender and category?
4.  How do category sales change year by year?
5.  What is the average discount by category?
6.  Which product brands/categories contribute most to sales?
7.  How does the sales mix differ across stores and regions?
8.  What does the forecast worksheet suggest about future sales, and how
    wide are its confidence bounds?

### Suggested PivotTable setup

-   **Rows:** `Category`, `Region`, `Store_Type`, `Product_Brand`, or
    Year from `Order_Date`
-   **Columns:** `Gender`, `Category`, or `Store_Type`
-   **Values:** Sum of `Total_Amount`, Sum of `Quantity`, Average of
    `Discount`, Count of `Sales_ID`
-   **Filters/Slicers:** `Order_Date`, `Region`, `Category`, `Gender`,
    `Store_Type`, `Payment_Type`

Refresh PivotTables after the source data is cleaned or updated.

## Dashboard Recommendations

A useful Excel dashboard can include: - KPI cards: Total Sales, Number
of Transactions, Total Quantity, Average Order Amount - Sales by
Category - Sales by Region and Store Type - Sales Trend by Month/Year -
Sales by Payment Type - Average Discount by Category - Interactive
slicers for Date, Category, Region, and Store Type

Ensure each chart has a clear title, readable labels, consistent number
formats, and a stated date range.

## Key Data Quality Notes

-   The workbook contains approximately 2,001 rows in `Sales_Fact`, 500
    customers, 100 products, and 20 stores.
-   The main transaction data spans 14 October 2023 to 13 October 2025.
-   Some sheets contain blank helper columns or notes, so remove only
    columns confirmed to be unnecessary.
-   Some customer loyalty values are recorded as `Unknown`; preserve
    them unless an approved rule provides the correct level.
-   Validate the meaning of `Discount` and the calculation of
    `Total_Amount` before recalculating sales values.
-   Confirm IDs and relationships across fact and dimension tables
    before drawing final conclusions.

## Deliverables

-   Cleaned and standardized Excel data
-   Documented data-cleaning workflow
-   Formula-based data-quality checks and descriptive metrics
-   PivotTables for category, region, store, product, gender, year, and
    discount analysis
-   Excel dashboard
-   Forecast worksheet for trend exploration

## How to Use

1.  Download or clone this repository.
2.  Open `Ecommerce_Sales_Dataset.xlsx` in Microsoft Excel.
3.  Review the fact and dimension worksheets.
4.  Apply or verify the documented cleaning steps.
5.  Update formula ranges to match the cleaned table.
6.  Refresh PivotTables and dashboard elements.
7.  Validate calculations and document insights before sharing results.

## Limitations

-   Formula examples assume a conventional header row and may require
    updated column references after cleaning.
-   `Total_Amount` should not be recalculated until the discount
    definition and business logic are verified.
-   Forecast results are estimates, not guarantees; evaluate the
    forecasting method and confidence bounds before making decisions.
-   Findings depend on the quality and completeness of the supplied
    workbook.

## Skills Demonstrated

**Excel Data Cleaning · Data Standardization · Data Quality Checks ·
Excel Formulas · Descriptive Statistics · PivotTables · Data
Visualization · Dashboard Reporting · Sales Trend Analysis · Forecast
Interpretation**

------------------------------------------------------------------------

*Prepared as a GitHub README for an Excel-based E-Commerce Sales
Analysis mini project.*

