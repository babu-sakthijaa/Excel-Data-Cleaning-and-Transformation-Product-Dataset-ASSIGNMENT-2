# 🧹 Excel Data Cleaning and Transformation – Product Dataset

## 📌 Project Overview

This project is part of my journey toward becoming a **Data Analyst**. The objective of this assignment is to develop practical skills in **data cleaning, data preparation, standardization, and transformation using Microsoft Excel**.

Real-world datasets often contain missing values, inconsistent text, duplicate records, and poorly structured information. Before performing analysis, these data quality issues need to be identified and handled appropriately.

In this project, I cleaned and transformed a **Product Dataset** by handling missing values, correcting inconsistencies, removing duplicates, restructuring columns, applying number formatting, and using conditional formatting.

---

## 🎯 Project Objectives

The main objectives of this project are to:

* Identify and handle missing values.
* Detect and correct inconsistent data.
* Standardize text values using Find and Replace.
* Identify and remove duplicate records.
* Split and restructure information from the Product ID.
* Merge product and brand information.
* Apply appropriate number and date formatting.
* Use conditional formatting to improve data readability.
* Prepare a clean dataset for further analysis.

---

# 📁 Dataset

The dataset contains product-level information with the following attributes:

| Column       | Description                                                             |
| ------------ | ----------------------------------------------------------------------- |
| Product ID   | Unique identifier containing manufacturing date and country information |
| Product Name | Name of the product                                                     |
| Brand Name   | Brand associated with the product                                       |
| Quantity     | Quantity of the product                                                 |
| Category     | Product category                                                        |
| Price        | Product price                                                           |

---

# 🧹 Data Cleaning and Transformation Performed

## 1. Handling Missing Values

### Missing Price Values

I checked the **Price** column for missing values.

For products with missing price information, the appropriate approach is to investigate the reason for the missing value before replacing it.

Possible strategies include:

* Check the original source for the correct price.
* Use the product's available information to identify the correct value.
* If appropriate, use a category-level or product-level estimate.
* If the price cannot be reliably determined, retain it as missing rather than introducing an incorrect value.

The chosen approach should depend on the business context and the amount of missing data.

### Missing Category Values

I also checked the **Category** column for missing values.

Possible approaches include:

* Identify the category using the Product Name or Brand Name.
* Compare similar products and assign the appropriate category.
* Use the most frequent category only when it is logically appropriate.
* Keep the value as missing if there is insufficient information to determine the category.

The goal is to avoid introducing incorrect information into the dataset.

---

# 2. 🔤 Correcting Inconsistent Data

## Product Name Standardization

I inspected the **Product Name** column for inconsistent text formatting.

Examples of issues that can occur include:

* Different capitalization.
* Extra spaces.
* Inconsistent spelling.
* Variations of the same product name.

I used Excel's **Find and Replace** functionality to standardize inconsistent product names.

For example:

```text
iphone 15
IPhone 15
IPHONE 15
```

can be standardized to:

```text
iPhone 15
```

This improves consistency and makes the dataset easier to analyze.

---

## Category Correction

I also checked the **Category** column for spelling mistakes and inconsistent category names.

For example:

```text
Electonics
```

could be corrected to:

```text
Electronics
```

Excel's **Find and Replace** feature was used to correct identified typos and standardize category values.

---

# 3. ♻️ Removing Duplicate Records

Duplicate records can affect analysis by causing products to be counted more than once.

I checked the dataset for duplicate rows based on the **entire row**.

Excel's built-in:

**Data → Remove Duplicates**

feature was used to identify and remove duplicate records where applicable.

This helps ensure that each unique product record is represented correctly in the cleaned dataset.

---

# 4. 🔄 Splitting and Merging Data

## Splitting Product ID

The **Product ID** contains multiple pieces of information.

For example:

```text
28-JAN-US
```

The Product ID was separated into:

* Manufacturing Date
* Country Code

Unnecessary characters such as separators were removed during the transformation.

The resulting structure is intended to make the information easier to analyze.

---

## Manufacturing Date

The date information extracted from the Product ID was converted into a proper Excel date format.

The required display format is:

```text
DD-MM-YYYY
```

For example:

```text
28-01-2026
```

This allows the date to be used properly in sorting, filtering, and date-based analysis.

---

## Country Code

The country information was extracted from the Product ID and stored in a separate **Country Code** column.

Example:

```text
Product ID: 28-JAN-US
Country Code: US
```

---

## Merging Brand Name and Product Name

A new column named **Product Brand** was created by combining the **Brand Name** and **Product Name** columns.

Example:

```text
Brand Name: Apple
Product Name: iPhone 15
```

Result:

```text
Apple - iPhone 15
```

An Excel formula that can be used for this transformation is:

```excel
=B2&" - "&C2
```

This creates a single descriptive field containing both the brand and product name.

---

# 5. 💰 Number Formatting

## Price Formatting

The **Price** column was formatted as currency to improve readability and clearly represent monetary values.

Example:

```text
1000
```

is displayed as:

```text
$1,000.00
```

Currency formatting also makes the dataset easier to interpret in reports and dashboards.

---

## Manufacturing Date Formatting

The **Manufacturing Date** column was formatted using:

```text
DD-MM-YYYY
```

Example:

```text
28-01-2026
```

This provides a consistent date representation throughout the dataset.

---

# 6. 🎨 Conditional Formatting

## Price Column

Conditional formatting was applied to the **Price** column using a **Data Bar or Color Scale**.

This makes it easier to visually identify:

* Lower-priced products.
* Medium-priced products.
* Higher-priced products.

Data bars provide a quick visual comparison between product prices.

---

## Electronics Category

A custom conditional formatting rule was created for the **Category** column to highlight cells where the category is:

```text
Electronics
```

This allows Electronics products to be identified quickly within the dataset.

A custom Excel formula can be used such as:

```excel
=F2="Electronics"
```

The exact cell reference should be adjusted according to the final position of the Category column.

---

# 📊 Data Cleaning Workflow

The overall cleaning process followed this workflow:

```text
Raw Product Dataset
        ↓
Check Missing Values
        ↓
Correct Inconsistent Text
        ↓
Fix Category Typos
        ↓
Remove Duplicate Rows
        ↓
Split Product ID
        ↓
Create Manufacturing Date
        ↓
Create Country Code
        ↓
Merge Brand + Product Name
        ↓
Apply Number Formatting
        ↓
Apply Conditional Formatting
        ↓
Clean Dataset
        ↓
Ready for Analysis
```

---

# 🛠️ Excel Features Used

| Excel Feature / Function | Purpose                           |
| ------------------------ | --------------------------------- |
| Find & Replace           | Standardize inconsistent text     |
| Remove Duplicates        | Remove duplicate records          |
| Text to Columns          | Split Product ID information      |
| Concatenation `&`        | Merge Brand Name and Product Name |
| Currency Formatting      | Format Price values               |
| Date Formatting          | Format Manufacturing Date         |
| Conditional Formatting   | Highlight important values        |
| Data Bars                | Visualize price differences       |
| Color Scales             | Compare price values              |
| Custom Rules             | Highlight Electronics category    |

---

# 📋 Data Quality Improvements

Before cleaning, the dataset could contain common real-world data quality issues such as:

* Missing values
* Inconsistent product names
* Category spelling errors
* Duplicate rows
* Combined information in a single column
* Inconsistent date representation
* Unformatted price values

After cleaning and transformation, the dataset becomes:

* More consistent
* Easier to read
* Easier to filter and sort
* More suitable for analysis
* More suitable for visualization
* Better prepared for reporting

---

# 💡 Key Learnings

Through this project, I practiced important data preprocessing concepts:

* Identifying missing values.
* Making decisions about missing data.
* Standardizing text values.
* Correcting spelling inconsistencies.
* Removing duplicate records.
* Splitting combined information into separate columns.
* Combining multiple columns into a single field.
* Formatting numerical and date data.
* Applying conditional formatting.
* Preparing raw data for further analysis.

---

# 🚀 Future Improvements

As I continue developing my Data Analyst skills, I plan to extend this project by:

* Creating an Excel PivotTable from the cleaned dataset.
* Building an interactive Excel dashboard.
* Adding charts and KPIs.
* Performing category-level analysis.
* Analyzing product and brand performance.
* Performing additional data validation.
* Recreating the cleaning process using **Python and Pandas**.
* Loading the cleaned dataset into **SQL** for further analysis.
* Comparing Excel, SQL, and Python data-cleaning workflows.

---

# 📂 Project Structure

```text
excel-data-cleaning-transformation/
│
├── README.md
│
└── Assignment 2 - Data Cleaning and Transformation.xlsx
```

---

# 👨‍💻 About Me

I am building my skills toward a career as a **Data Analyst** by working on practical projects and developing a portfolio.

This project is my second Excel-based portfolio project and focuses on **data cleaning and transformation**.

I am continuously developing my skills in:

* 📊 Microsoft Excel
* 🗄️ SQL
* 🐍 Python
* 📈 Data Visualization
* 📉 Statistics
* 🔍 Data Analysis
* 🧹 Data Cleaning

More projects will be added as I continue my Data Analyst learning journey.

---

⭐ **This project is part of my Data Analyst portfolio and learning journey.**
