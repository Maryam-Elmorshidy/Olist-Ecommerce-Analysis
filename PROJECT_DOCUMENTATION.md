# Olist Brazilian E-Commerce Data Analysis Using Excel

## Project Overview

This project involved analyzing the **Brazilian E-Commerce Public Dataset by Olist** using Microsoft Excel.

The work included understanding the datasets, cleaning and transforming the data using Power Query, building a data model, creating calculated columns and DAX measures, exploring the data using PivotTables and charts, and developing four dashboards.

The project was completed through an exploratory process. During the analysis, I investigated data quality issues, examined relationships between different variables, and formulated questions to better understand the data.

---

## Table of Contents

1. [Understanding the Data](#1-understanding-the-data)
2. [Data Cleaning](#2-data-cleaning)
   - [Sellers](#sellers)
   - [Product Category Name Translation](#product-category-name-translation)
   - [Products](#products)
   - [Order Reviews](#order-reviews)
   - [Orders](#orders)
   - [Order Payments](#order-payments)
   - [Order Items](#order-items)
   - [Geolocation](#geolocation-1)
   - [Customers](#customers-1)
3. [Additional Columns](#3-additional-columns)
4. [Data Model](#4-data-model)
5. [DAX Calculations](#5-dax-calculations)
6. [Data Exploration and Analysis](#6-data-exploration-and-analysis)
   - [Products Analysis](#products-analysis)
   - [Sales Analysis](#sales-analysis)
   - [Orders Analysis](#orders-analysis)
7. [Final Dashboards](#7-final-dashboards)
8. [Conclusion](#conclusion)

---

## 1. Understanding the Data

I started by loading all the datasets into Power Query.

To make the work easier, I simplified the names of the tables as follows:

| Original Table Name | New Table Name |
|---|---|
| Olist_customers_dataset | Customers |
| Olist_geolocation_dataset | Geolocation |
| Olist_order_items_dataset | Order_items |
| Olist_order_payments_dataset | Order_payments |
| Olist_order_reviews_dataset | Order_reviews |
| Olist_orders_dataset | Orders |
| Olist_products_dataset | Products |
| Olist_sellers_dataset | Sellers |
| Product_category_name_translation | Kept the original name temporarily because I had not yet understood the purpose of this table |

### Initial Examination of the Tables

#### Sellers
- No duplicate values were found in `seller_id`.

#### Product Category Name Translation
- The first row contained the column names, so I corrected it.
- The table contained no duplicate rows.

#### Products
- No duplicate values were found in `product_id`.

#### Order Reviews
- The table did not have a primary key.

#### Orders
- No duplicate values were found in `order_id`.

#### Geolocation
I noticed that `geolocation_zip_code_prefix` contained duplicate values.

After examining the corresponding values in the other columns, I found that:
- `geolocation_city` and `geolocation_state` remained the same for the duplicated ZIP code prefixes.
- `geolocation_lat` and `geolocation_lng` contained very similar values, with only minor differences.

Based on these observations, I decided to remove duplicates from `geolocation_zip_code_prefix`, making it unique so that it could be used as a primary key for the table.

#### Customers
- No duplicate values were found in `customer_id`, so I identified it as the primary key.

---

## 2. Data Cleaning

After examining the structure and quality of each table, I started cleaning the datasets individually.

### Sellers

#### 2.1. Correcting an Incorrect City Value

In the `seller_city` column, I found the value `04482255`.

The corresponding values in `seller_zip_code_prefix` and `seller_state` consistently indicated Rio de Janeiro.

Therefore, I replaced `04482255` with `rio de janeiro`.

#### 2.2. Cleaning the `seller_city` Column

The `seller_city` column contained many errors.

I followed these steps:
1. Applied **Trim** to remove extra spaces.
2. Because the column contained numerous errors and was difficult to clean efficiently through the standard interface, I edited the code directly in the **Advanced Editor**.
3. I performed 52 replacement operations. These operations did not appear individually in the Applied Steps section because I handled them directly in the code.
4. Later, I modified the code with AI assistance to correct additional errors.

#### 2.3. Cleaning the `seller_state` Column

I extracted the values that needed correction into separate tables and then appended the corrected values to the original table.

I created separate files for the values under the name `Modified_Values_For_Sellers`.

While working on the second column, I noticed extra spaces, so I applied Trim.

I also corrected the fourth value that I had identified during the cleaning process.

#### 2.4. Checking for Duplicates Again

After making the corrections, I checked for duplicate values again. This helped reveal additional errors in the final column.

I found 10 duplicate values that resulted in incorrect values in that column.

I decided to correct these values using the M language, with assistance from AI, because I was not proficient in M.

#### 2.5. Resolving the Formula.Firewall Issue

I postponed the append operation because I encountered a `Formula.Firewall` error.

I resolved the issue by changing some Power Query settings, although I do not think this was necessarily the best approach.

Later, I discovered that the append operation could have been handled using an `if` statement. However, since I had already spent considerable time on my original approach, I decided to keep my existing implementation.

#### 2.6. Identifying Values Missing from Geolocation

I created a separate table named `ValuesInSellersOnly` and placed it in the `Missing` file.

While working with PivotTables, I discovered that seven values existed in Sellers but were missing from Geolocation, even though Geolocation was supposed to be the reference table.

#### Reflection on Cleaning Sellers

Although I spent a considerable amount of time cleaning this table, AI identified and corrected several additional errors when I provided it with the resulting data. I adopted those corrections.

✅ The **Sellers** table was now cleaned.

---

### Product Category Name Translation

#### 2.7. Cleaning `product_category_name_english`

I sorted the column alphabetically and reviewed its 71 values to identify errors.

I found three incorrect values:
- `Costruction_tools_garden`
- `Costruction_tools_tools`
- `fashio_female_clothing`

#### 2.8. Examining `product_category_name`

The values in this column were written in a language I did not understand, so I did not know how to handle them at that stage.

✅ The **Product Category Name Translation** table was now cleaned.

---

### Products

#### 2.9. Investigating Null Values in `product_category_name`

I found 610 null values in the relevant section of the `product_category_name` column.

The main difficulty was that I did not have another reference source to verify these values.

Until I could determine how to handle them, I separated them into a table named `NullValues_In_Products` in the `Missing` file.

I then created a copy of the Products table and removed all columns except `product_category_name`.

Next, I merged this column with the corresponding column in `Product_category_name_translation` and examined the values that were null in the translation table.

After comparing the two tables, I discovered that the Products table contained two values that did not exist in `Product_category_name_translation`, which was the reference table for the category translations.

I tried to find matching or similar values in the translation table, but I could not identify any sufficiently close matches.

Therefore, I separated these values into a table named `ValuesInProducts_NotFoundIn_Product_category_name_translation` in the `Missing` file.

I left these values outside the main table for the time being so that they would not affect the other data.

#### 2.10. Investigating `product_name_lenght`

There was no column containing the actual product names in this table or any other available table that I could use to verify the product name lengths.

Therefore, I left this column unchanged.

#### 2.11. Investigating `product_description_lenght`

There was no column containing the product descriptions in this table or any other available table that I could use to verify the description lengths.

Therefore, I left this column unchanged.

#### 2.12. Correcting Negative Values in `product_photos_qty`

I found three negative values in this column. These seemed illogical because they did not represent a meaningful number of product photos.

I tried to find a relationship between this column and the other columns, but I could not identify one.

I considered two possible approaches:
- Change the negative values to positive values.
- Treat the values as suspicious, keep them outside the main table, and investigate them further.

Since I did not have a reference source for these values, I chose the first approach and converted the negative values to positive values.

#### 2.13. Removing a Row Containing Null Values

I found one row containing null values in the last four columns.

I checked the related `Order_items` table, but I could not find the `product_id` associated with that row.

Based on this, I decided to delete the row.

✅ The **Products** table was now cleaned.

---

### Order Reviews

#### 2.14. Investigating `review_id` and `order_id`

I noticed that some values in `review_id` were duplicated, but the duplication was unusual.

The entire row appeared twice with the same date, time, name, and comment, while the `order_id` value differed.

As a result, each duplicated `review_id` was associated with two different `order_id` values.

Although I did not find a use for `review_id` in the data model, I decided to leave it unchanged rather than delete it.

#### 2.15. Examining `review_score`

I did not find any issues in this column.

#### 2.16. Handling Review Comments

I expected the `review_comment_title` and `review_comment_message` columns to contain empty values, so I left them unchanged.

#### 2.17. Calculating Review Response Time

I added a column named `Review Response Time - Days` to calculate the difference between `review_creation_date` and `review_answer_timestamp`.

The purpose was to obtain a measure of how quickly reviews were answered.

I then converted the resulting duration into days to make it easier to analyze.

✅ The **Order_reviews** table was now cleaned.

---

### Orders

#### 2.18. Checking `order_id` and `customer_id`

I found no duplicate values in either of the first two columns, `order_id` and `customer_id`.

#### 2.19. Investigating Missing Delivery Dates

The `order_delivered_carrier_date` and `order_delivered_customer_date` columns contained many null values.

These missing values could affect analyses comparing `order_delivered_customer_date` with `order_estimated_delivery_date`, for example.

#### 2.20. Investigating Null Values in `order_approved_at`

The `order_approved_at` column contained 160 null values.

Among these 160 rows, only 14 also contained null values in both `order_delivered_customer_date` and `order_estimated_delivery_date`. The remaining rows had null values in all three columns.

Since `order_approved_at` indicates that an order was approved, a null value in this column together with null values in the two subsequent delivery-related columns could indicate that the order had not progressed through the process.

Therefore, I decided to remove only the rows in which all three columns were null.

I used the following M code:

```m
#"Filtered Rows" =
    Table.SelectRows(
        #"Changed Type",
        each not (
            [order_approved_at] = null and
            [order_delivered_carrier_date] = null and
            [order_delivered_customer_date] = null
        )
    ),
```

#### 2.21. Exploring Null Values in Relation to `order_status`

While examining the remaining null values alongside `order_status`, I made the following observations.

**Delivered Orders**
There were eight cases where the order status was recorded as `delivered`, even though the delivery date was missing.

I decided to leave these values empty because I treated them as potentially unrecorded values rather than incorrect values.

**Canceled Orders**
All orders with the status `canceled` contained null values, except for six cases.

This could indicate, for example, that the customer canceled the order at the time of delivery or refused to receive it.

**Approved Orders**
There were two cases with the status `approved`. The values in the other columns appeared consistent with these statuses.

**Invoiced Orders**
There were 314 cases with the status `invoiced`, and all had values that appeared consistent with the neighboring columns.

**Processing Orders**
There were 301 cases with the status `processing`, and all had values that appeared consistent with the neighboring columns.

**Shipped Orders**
The cases with the status `shipped` appeared consistent with the order process. Null values appeared only in `order_delivered_customer_date`.

**Unavailable Orders**
It was reasonable for orders with the status `unavailable` to contain many null values.

#### 2.22. Considering an Estimated Delivery Duration Column

I considered adding a column to calculate the expected time between the purchase timestamp and the estimated delivery date:

```
order_estimated_delivery_date - order_purchase_timestamp
```

This would indicate the expected delivery duration from the time the order was placed.

✅ The **Orders** table was now cleaned.

---

### Order Payments

I noticed that `order_id` contained duplicate values, which was expected for this table.

I did not identify any other apparent issues in the remaining columns.

✅ The **Order_payments** table was now cleaned.

---

### Order Items

#### 2.23. Handling a Null Value in `price`

I found a null value in the `price` column.

To determine a possible correction, I checked `product_id` and compared the price of the same product in another order.

Based on that comparison, I identified a possible correction.

#### 2.24. Investigating Negative Prices

I found three negative values in the `price` column, which did not seem logical.

**First Negative Value: -110**
When I checked the product's prices, I found values of 110 and 150.

I searched the Products table but did not find evidence that the product had different sizes that could explain the price difference.

I considered several possible explanations, such as price changes over time, discounts, or other price adjustments.

Therefore, I left the value of 150 unchanged and corrected only the negative value.

**Second Negative Value**
The second negative value was also relatively straightforward to correct based on the available data.

**Third Negative Value**
For the third negative value, I followed the same reasoning. Since the data contained negative prices that appeared to have been entered incorrectly, I leaned toward correcting this value in the same way.

#### 2.25. Investigating Differences in Product Prices

I noticed that some products had different prices across orders. Some differences were minor, while others were larger.

However, I did not have an external reference source to verify the correct prices.

Since price differences could have several explanations, including different purchase times or customer-specific circumstances, I decided to leave the `price` column unchanged apart from the corrections described above.

I did not identify any apparent issues in the remaining columns.

✅ The **Order_items** table was now cleaned.

---

### Geolocation

The Geolocation table presented a problem similar to the one I encountered in Sellers.

To speed up the cleaning process, I used AI assistance.

After converting the values in `geolocation_city` into English, I used AI to help identify values that needed correction or replacement. I then removed the extra spaces.

Cleaning this table took a very long time and required several attempts because of the different issues it contained.

I believe using Python might have made this process faster, but I chose to complete the work without it.

✅ The **Geolocation** table was now cleaned.

---

### Customers

I used AI assistance to identify errors in the Customers table and correct them in the Advanced Editor.

I found two values that I could not identify directly, so I referred to the Sellers and Geolocation tables to help correct them.

✅ The **Customers** table was now cleaned.

---

## 3. Additional Columns

During the preparation and analysis stages, I added columns to support the analysis.

### 3.1. Estimated Delivery Time - Days

**Table:** Orders

I added this column to calculate the expected delivery duration from the order purchase timestamp to the estimated delivery date.

The calculation was based on:

```
order_estimated_delivery_date - order_purchase_timestamp
```

### 3.2. Review Response Time - Days

**Table:** Order_reviews

I added this column to calculate the difference between `review_creation_date` and `review_answer_timestamp`.

This measure could help evaluate how quickly reviews were answered.

I converted the resulting duration into days to make the analysis easier.

---

## 4. Data Model

After cleaning the tables, I created relationships between them to connect the data for analysis.

| Relationship Type | Key Used | First Table | Second Table |
|---|---|---|---|
| Many-to-One | `product_category_name` | Products | Product_category_name_translation |
| Many-to-One | `product_id` | Order_items | Products |
| Many-to-One | `seller_id` | Order_items | Sellers |
| Many-to-One | `order_id` | Order_items | Orders |
| Many-to-One | `order_id` | Order_payments | Orders |
| Many-to-One | `Geolocation_zip_code_prefix` / `seller_zip_code_prefix` | Sellers | Geolocation |
| Many-to-One | `Geolocation_zip_code_prefix` / `customer_zip_code_prefix` | Customers | Geolocation |
| Many-to-One | `customer_id` | Customers | Orders |
| Many-to-One | `order_id` | Order_reviews | Orders |

The data model connected the tables through their relevant identifiers, allowing me to explore information from different tables together.

---

## 5. DAX Calculations

During the analysis stage, I realized that I needed three additional columns.

### 5.1. Actual Delivery Time - Days

I should have added this column earlier, when I created `Estimated Delivery Time - Days`, but I forgot to do so at that stage.

I added it later using the following DAX formula:

```dax
Actual Delivery Time - Days =
DATEDIFF(
    Orders[order_purchase_timestamp],
    Orders[order_delivered_customer_date],
    DAY
)
```

I used this column together with `Estimated Delivery Time - Days` to compare the expected delivery duration with the actual delivery duration.

### 5.2. Year and Quarter Columns

I added `Year` and `Quarter` columns to the Orders table.

The reason was that when I used the standard date grouping in a PivotTable and then added a Slicer, the quarters appeared nested within each year. The resulting chart showed the years rather than allowing me to compare quarters across years in the way I wanted.

To address this issue, I created separate columns.

**Year**

```dax
Year =
FORMAT(
    Orders[order_estimated_delivery_date],
    "yyyy"
)
```

**Quarter**

```dax
Quarter =
CONCATENATE(
    "Qtr",
    INT(
        (MONTH(Orders[order_estimated_delivery_date]) + 2) / 3
    )
)
```

### 5.3. DAX Measures Used During Exploration

I created four DAX measures during the exploration stage.

**1. Average Price**

```dax
Average Price :=
AVERAGE(Order_items[price])
```

This measure calculates the average value in the `price` column.

**2. Total Price**

```dax
Total Price :=
SUMX(Order_items, Order_items[price])
```

This measure calculates the sum of the values in the `price` column across the `Order_items` table.

**3. TotalSales**

```dax
TotalSales :=
CALCULATE(
    SUM(Order_items[price]),
    Orders[order_status] = "delivered"
)
```

I used this measure to calculate total sales by summing the item prices associated with orders whose status was `delivered`.

**4. Average Review Response Time**

```dax
Average Review Response Time :=
AVERAGE(Order_reviews[Review Response Time - Days])
```

This measure calculates the average review response time in days.

---

## 6. Data Exploration and Analysis

After completing the cleaning, modeling, and calculation stages, I moved to a stage I called **Exploration**.

I kept the worksheet I used during this stage in the Excel file. I also prepared a separate file containing my attempts to understand the data more deeply, particularly after connecting the tables, along with the questions I asked during the exploration stage.

I conducted the analysis by reviewing the data again, defining the columns, asking exploratory questions, and using PivotTables and charts to investigate possible answers.

I divided the dashboards into three main sections:
- Products
- Sales
- Orders

Accordingly, I created three worksheets:
- `Product _ pivot table`
- `Sales _ pivot table`
- `Orders _ pivot table`

After dividing the analysis into these sections, I began organizing the questions relevant to each one.

---

### Products Analysis

**Worksheet:** `Product _ pivot table`

#### 6.1. Which product categories have the highest number of orders? (Top 10)

**Objective:** Identify the product categories with the highest number of orders.

I also tried to understand why these categories received so many orders.

One possible explanation is that some product categories contain a larger number of available products than others. This could lead to more variety and, consequently, more orders.

I also examined whether product prices could explain the high order volumes, but I did not find a relationship between the two in this analysis.

#### 6.2. Which product categories are most ordered in the top 5 cities by order volume?

**Objective:** Identify the five cities with the highest order volumes and determine which product categories are most frequently ordered in those cities.

This analysis allowed me to examine product category preferences within the cities with the highest order volumes.

#### 6.3. Which product categories have the lowest number of orders? (Bottom 10)

**Objective:** Identify the product categories with the lowest order volumes.

Understanding which categories receive the fewest orders could help identify areas for further investigation.

One possible explanation is that some categories contain fewer available products than others, which may result in fewer orders.

#### 6.4. Top 10 Products by Sales

**Objective:** Identify the top 10 products by sales.

I observed a strong positive relationship between sales and order volume.

This suggested that higher sales could be associated with a higher number of orders. One possible explanation is that categories with more available products offer greater variety, potentially resulting in more orders.

#### 6.5. Which categories have the highest average product price?

**Objective:** Identify the product categories with the highest average prices.

The high average prices of these categories may simply reflect the fact that their products are generally more expensive.

#### 6.6. What share of total orders comes from the top 5 product categories?

**Objective:** Determine the proportion of total orders represented by the five most frequently ordered product categories.

A possible explanation for their high order volumes is that these categories contain a larger number of available products, offering more variety and potentially attracting more orders.

#### 6.7. Timeline

After exploring these questions, I created a **Timeline** to examine the data over time.

#### Products Dashboard

I brought together the product-related analysis in the `Products_Dashboard`.

---

### Sales Analysis

**Worksheet:** `Sales _ pivot table`

#### 6.8. Do sales increase or decrease over time?

**Objective:** Understand how sales changed over time.

The analysis showed that sales had been increasing continuously. However, at the beginning of 2018, this growth stopped, and sales began to decline gradually.

During the third quarter of 2018, sales decreased sharply.

I then examined the relationship between sales and order volume and found a strong relationship between them.

After examining the data and investigating its date range, I discovered that the latest date in the dataset was **October 25, 2018**.

This could explain the decline in both sales and orders during the final quarter of 2018, because the dataset did not cover the entire quarter.

#### 6.9. Where do sales increase? (Top 10)

**Objective:** Identify the ten cities with the highest sales.

The high sales in these cities may be related to their high order volumes.

In particular, the top three cities by sales were also the top three cities by order volume.

#### 6.10. Which products have the highest sales? (Top 10)

**Objective:** Identify the top 10 products by sales.

As mentioned earlier, I observed a strong positive relationship between sales and order volume.

Therefore, higher sales may be associated with a greater number of orders. One possible explanation is that categories containing more available products offer greater variety and may receive more orders.

#### 6.11. Do higher sales always correspond to more orders? (Top 15 best-selling products)

**Objective:** Compare sales and order volumes for the 15 most frequently ordered product categories.

I observed a strong relationship between sales and order volume.

In many contexts, this relationship may be expected, so I did not consider it necessary to investigate a separate explanation for the relationship itself.

#### 6.12. What are the total sales?

**Objective:** Determine the total sales value.

This question helped me identify the overall sales figure for the dataset.

#### 6.13. Do higher order volumes correspond to higher sales?

**Objective:** Compare overall order volume with sales over time.

The relationship between sales and order volume had already been discussed in the previous analyses.

The fact that the dataset ended on October 25, 2018, was also relevant when interpreting the decline toward the end of 2018.

#### 6.14. State Slicer

I added a Slicer for State and used it to examine sales across different states.

#### Sales Dashboard

I brought together the sales-related analysis in the `Sales_Dashboard`.

---

### Orders Analysis

**Worksheet:** `Orders _ pivot table`

#### 6.15. How does the number of orders change over time?

**Objective:** Understand how order volume changed over time.

As mentioned earlier, order volume had a strong positive relationship with sales.

The dataset also ended on October 25, 2018.

These two observations may help explain the pattern shown in the chart, particularly the decline toward the end of the period.

#### 6.16. Which product categories have the highest number of orders? (Top 10)

This question was already covered in the `Product _ pivot table` worksheet.

#### 6.17. How does the average actual delivery time compare with the estimated delivery time?

**Objective:** Determine whether orders were delivered earlier than expected, on time, or later than expected.

Although some values were null, the results indicated that most orders were delivered earlier than their estimated delivery dates.

This appeared to be a positive outcome in terms of delivery performance.

#### 6.18. Order Delivery Status

**Objective:** Determine the proportion of orders that were delivered.

#### 6.19. Which cities have the highest number of orders? (Top 5)

**Objective:** Identify the five cities with the highest order volumes.

While investigating the data, I found a close relationship between the number of customers and order volume.

The high order volumes in these cities may be related to their larger geographical areas and larger populations.

#### 6.20. What is the total number of orders?

**Objective:** Determine the total number of orders in the dataset.

#### 6.21. Do higher sales always correspond to more orders? (Top 15 best-selling products)

This question was not included in the final interactive dashboard, but I had already discussed it in the earlier analysis.

#### 6.22. Do higher order volumes correspond to higher sales?

This question was also discussed earlier but was not included in the final interactive dashboard.

#### 6.23. Year Slicer

I created a Slicer for the three years represented in the data and used it to examine order volume across those years.

#### Orders Dashboard

I brought together the order-related analysis in the `Orders_Dashboard`.

---

## 7. Final Dashboards

I organized the final interactive dashboard into three main sections, in addition to a main dashboard.

### Main Dashboard
The main dashboard provided an overall view of the project.

### Products Dashboard
This dashboard focused on product categories, order volumes, sales by product, and average product prices.

### Sales Dashboard
This dashboard focused on sales trends over time, sales by city, product sales, and the relationship between sales and order volume.

### Orders Dashboard
This dashboard focused on order trends, delivery performance, order status, and the cities with the highest order volumes.

The four dashboards created for the project were:
1. Main Dashboard
2. Products Dashboard
3. Sales Dashboard
4. Orders Dashboard

---

## Conclusion

This project involved understanding, cleaning, modeling, and analyzing the Olist e-commerce dataset using Excel.

The process included working with Power Query, editing M code in the Advanced Editor, creating relationships between tables, writing DAX calculations, exploring the data with PivotTables and charts, and building interactive dashboards.

Throughout the project, I encountered data quality issues and situations where the available data did not provide a clear answer. In those cases, I documented my observations and made decisions based on the information available to me.

The exploration stage helped me become more familiar with the dataset, investigate relationships between variables, and organize the analysis around specific questions.

Although there are areas that could be improved, the project was a valuable learning experience in data preparation, data modeling, exploratory analysis, and dashboard development using Excel.

---

**Dataset:** [Brazilian E-Commerce Public Dataset by Olist](https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce) on Kaggle
