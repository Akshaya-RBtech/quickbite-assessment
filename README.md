# QuickBite – Food Delivery Data Analysis

## Project Overview

QuickBite is a food delivery application operating in Bengaluru, Hyderabad, Chennai and Mumbai.

This project focuses on cleaning, combining and analysing food delivery data to understand area performance, customer food preferences and delivery partner performance.

## Dataset Description

The provided Excel workbook contains six tables:

- **Orders:** Order details, amount, payment mode, status and delivery time.
- **Customers:** Customer information.
- **Restaurants:** Restaurant names and cuisine details.
- **Areas:** Area names and city information.
- **Delivery Partners:** Delivery partner details and vehicle types.
- **Feedback:** Customer ratings and comments.

## Project Tasks

### 1. Data Cleaning

- Removed duplicate records.
- Handled missing values.
- Standardized city, area, cuisine, payment mode and order status values.
- Corrected inconsistent text formats.
- Converted date and numeric columns into proper formats.
- Handled invalid ratings and delivery times.

**Output:** `cleaned_data/Quick_Bite_Data_Cleaning.xlsx`

### 2. Data Combining

- Combined all cleaned tables using their ID columns.
- Used left joins with Orders as the main table.
- Combined multiple feedback records into one order-level record.
- Created `order_day` from the order date.
- Created `is_late` based on delivery time greater than 45 minutes.
- Ensured each order appears only once.

**Output:** `cleaned_data/Quick_Bite_Combined_Data.xlsx`

### 3. Data Analysis

The analysis uses delivered orders only.

**Area Performance**
- Total orders
- Total revenue
- Average order value
- Average delivery time
- Late delivery count and percentage
- Average customer rating

**Top Cuisine by Day**
- Identified the most ordered cuisine for each day.
- Calculated order count and percentage share.

**Vehicle Performance**
- Total deliveries by vehicle type
- Average delivery time
- Late delivery count and percentage
- Average customer rating
- Revenue handled

**Output:** `analysis/Quick_Bite_Analysis.xlsx`

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Microsoft Excel
- VS Code

## Project Structure

```text
quickbite-assessment/
│
├── data/
│   └── QuickBite_Assessment.xlsx
│
├── scripts/
│   ├── cleaning.py
│   └── analysis.py
│
├── cleaned_data/
│   ├── Quick_Bite_Data_Cleaning.xlsx
│   └── Quick_Bite_Combined_Data.xlsx
│
├── analysis/
│   └── Quick_Bite_Analysis.xlsx
│
└── README.md
```

## How to Run the Project

**Step 1: Install the required libraries**

```bash
pip install pandas openpyxl
```

**Step 2: Run the data cleaning script**

```bash
python scripts/cleaning.py
```

**Step 3: Run the data analysis script**

```bash
python scripts/analysis.py
```

## Conclusion

This project provides insights into food delivery business performance through data cleaning, data integration and analysis. It helps identify area-wise revenue, popular cuisines by day and delivery partner performance.