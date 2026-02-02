# Excel-datasets
# HR Dataset Assignment https://dev.to/dishon_gatambiadd_31a1/hr-dataset-assignment-de-temp-slug-4969716
Data Cleaning & Preparation
Duplicates- To remove duplicates, used the OrderID column as it the unique identifier. No duplicates were found

Formatting cells: Assigned each column a data category. Below are the columns and categories

OrderID - Text

OrderDate - Date

RequiredDate - Date

Region - Text

Country - Text

City - Text

CustomerSegment - Text

Channel - Text

Salesperson - Text

ProductCategory - Text

SKU - Text

UnitCost - Number

UnitPrice - Number

DiscountPct - Number

Quantity - Number

Missing Values were handled by visualising the columns City, Salesperson & Channel using conditional formatting. After blanks were visualised inferences were tried o be drawn up eg: A city with a country. No inference was made. The blanks were filled using the Find & Replace function with "Unknown"

Negative values in the UnitPrice were removed using the Find & Replace function. For the Discounts column, used conditional formatting to highlight discounts ">30%"

To correct the RequiredDate & OrderDate errors where RequiredDate > OrderDate, created another column LeadTimeDays. For this column, I found the difference between RequiredDate & OrderDate ie =(C2-B2). Found the mode of the column ie =Mode(E2:E632) The mode =20. For the LeadTimeDays column used an IF formula with a subtraction in it ie =IF(C2-B2<0,"20",C2-B2). This formula changed the negative values and replaced them with the mode of 20 days, as this was taken as the standard lead time. Created a new column CorrectedRequiredDate where another IF function was used ie =IF(C2<B2,C2+20,C2). This function checks if the OrderDate > RequiredDate and if it is it adds 20 days to the RequiredDate if not it returns the original RequiredDate

Created 4 calculated columns using the formulas listed beside them:

GrossRevenue - =(O2*Q2)*(1-P2)
2. CostofGoods - =(N2*Q2)
GrossProfit - =(R2-S2)
MarginPct - =IF(R2=0,"0",T2/R2)
Created a column called Month. The months were derived using the TEXT formula ie =TEXT(B2,"MMM-YYY")
Created a column Quarter and used the formulas: Month & Year which return the value of the referenced month and year and the Roundup function which rounds up numbers away from 0 to specified decimal places ie =Roundup(3.142159,3)=3.142 & =Roundup(3.2,0)=4. Used the & function as a joiner ie ="Q" & ROUNDUP(MONTH(B2)/3,0) & "-" & YEAR(B2)
Created a pivot table in a new data sheet and added the Region, Country & City in that order to create a hierarchy
Created a column PriceBand which displays the percentiles using the Percentile.Inc formula ie =PERCENTILE.INC(Q2:Q632,0.33) & =PERCENTILE.INC(Q2:Q632,0.67) which returns the 33rd & 67th percentile. Created another column ProductCategory that puts each category either low, medium or high using the quantiles ie =IF(Q2<$X$2,"Low",IF(Q2<$X$3,"Medium","High"))

To come up with the saleperson productivity:
Created a pivot table that had the salesperson as the 1st column. Against each salesperson, count of OrderID, sum of GrossRevenue & GrossProfit was found. To enable computing, copy pasted to a new cell to enable formulas. Came up with three new columns, Revenue/Order, GrossProfit/Order & Order/Month. The formulas are below

Revenue/Order - _SumofGrossRevenue/CountofOrderID_

GrossProfit/Order - _SumofGrossProfit/CountofOrderID_

Order/Month - _CountofOrderID/12_
To come up with the top and bottom 3, used multi-level sort function The hierarchy was GrossProfit/Order>Orders/Month>Revenue/Order
Top 3: C.Otieno>E.Garcia>G.Dubois
Bottom 3: I.Johnson>F.Muller>H.Kim

To come up with Service Level Proxy:
Came up with a column Target which contains the formula that assigns a value "1" less than 7 days and "0". Inserted a pivot table in a new worksheet "ServiceLvlProxy" that puts Regions>Channels in rows and average of Target. This gets the average of products that were within the target of 7 days

To come up with Price Compliance:
In the main Data worksheet, came up with a column DiscoPct that categorizes discounts >20% into a value, "1". Came up with a pivot table the rows being in a hierarchy, Africa>Salesperson & the values being the average of discounts above 20%. This pivot table was placed in a new worksheet, "PriceCompliance"

Cohort analysis???
ABC analysis by SKU???
Channel mix & Cannibalization???
The What If Scenario Modelling
Opened a new worksheet "CtrlPanelData". Came up with the baseline calculations ie

Baseline Total Revenue - =SUM(Data!W2:W632) The sum of the GrossRevenue

Baseline Total Cost - =SUM(Data!X2:X632) The sum of CostofGoods

Baseline Total Profit - =SUM(Data!Y2:Y632) The sum of GrossProfit

Baseline Margin Percent - =IF(D5=0,0,D7/D5)
The scenario have caps. The scenario calculations are below

Scenario Total Revenue - =SUM((Data!T2*Data!V2)*(1+CtrlPanelData!$B$4)*(1-MIN(Data!U2,CtrlPanelData!$B$2)))

Scenario Total Cost - =SUM(Data!S2*(1+CtrlPanelData!$B$3))*(Data!V2)*(1+CtrlPanelData!$B$4)

Scenario Total Profit - =E5-F5

Scenario Margin % - =IF(E5=0,0,G5/E5)
The scenario and baseline metrics will give rise to the chart to visualise the differences.
Came up with arbitrary scenario values that were within the caps for: Global Discount, UnitCost Inflation & Quantity Uplift
Created a pivot table with the metrics ie Cost, Revenue, Profit & Margin as the rows. The columns were the averages of baseline and scenario to compare the differences.
In a new worksheet, "WhatIfCtrlPanel". For the inputs of Global Discount Cap, UnitCost Inflation & Quantity Uplift, formatted the cells for data validation to prevent cap values from being exceeded.
Created a column chart that contains the metrics and visualises the differences between the baseline and scenario. Added a slicer that contains the metrics.

Interactive Dashboard
Created 2 new worksheets, "DashBoardCalc" for the data & "Dashboard" for the interactive dashboard.
Came up with 4 pivot tables for the visuals:

Month against the sum of revenues to visualise a line chart of Revenue by Month.
Region>Channel in rows against sum of GrossProfit to visualise a stacked column chart of Profit by Region & Channel.
SKU in rows by Sum of GrossRevenue to visualise a line chart of top 10 SKU by Revenue. Filtered the pivot table to show top 10.
Salesperson in Rows by Average of DiscoPctCat to visualise a bar chart of the discount outliers.
Came up with another pivot table for the KPIs. These were arranged in the columns:

Sum of GrossRevenue

Average of GrossRevenue

Sum of GrossProfit

AverageMargin %

Average LeadTimeDays

Count of OrderIF

To visualize the KPI cards:
To calculate the Average Order Value: Sum of GrossRevenue/Count of OrderID ie =IFERROR((GETPIVOTDATA("Sum of GrossRevenue",DashBoardCalc!$O$1)/GETPIVOTDATA("Count of OrderID",DashBoardCalc!$O$1)),0)
The rest were gotten from the pivot tables:
Total GrossRevenue - =GETPIVOTDATA("Sum of GrossRevenue",DashBoardCalc!$O$1)
Total GrossProfit - =GETPIVOTDATA("Sum of GrossProfit",DashBoardCalc!$O$1)
Margin % - =GETPIVOTDATA("Average of MarginPct",DashBoardCalc!$O$1)
Average Lead Time Days - =GETPIVOTDATA("Average of LeadTimeDays",DashBoardCalc!$O$1)
Created interactive slicers for:

Region
Country
Channel
ProductCategory
Month
Salesperson

# Jumia Dataset https://dev.to/dishon_gatambiadd_31a1/jumia-dataset-3inf-temp-slug-6825144
Data Cleaning
Removed the negative values in the 'review' column by using the find and select feature and replaced with a blank
Removed the "out of 5" in the 'rating' column by using the find and replace feature and replaced with a blank
Removed the blanks in the 'rating' column by using the find and replace feature and replaced with the average which amounted to 3.8. Found the average by using the average function ie =AVERAGE(G2:G113)
Removed the blanks in the review column using the find and select feature and replaced with '0'

In the current price and old price columns, removed the 'Ksh' using the find and select feature and replaced with a blank

Formatted the current price and old price columns into currency, KES

Formatted the discount column into percentage

Formatted the rating & reviews into numbers

Data Enrichment
Created a column Discount which was the old price minus current price divided by old price ie =(C2-B2)/C2 Dragged the formula down then formatted the cells to percentage.

Created a column called Rating Category. The column contains an IF function where if rating is <3, "poor", 3-4.4, "average", >4.5, "excellent" ie =IF(G2<3,"Poor",IF(G2<4.5,"Average","Excellent"))

Created a column called Discount Category. The column contains IF function where if discount <20%, "Low Discount", 20%-40%, "Medium Discount", >40%, "High Discount" ie =IF(D2<20%,"Low discount",IF(D2<40%,"Medium Discount","High Discount"))

Data Analysis
Found the average of the columns: Old Price, Current Price, Discount Percentage & Ratings ie Average(B2:B113)

Found the most expensive & Least expensive item by using the Max and Min function ie =MAX(B2:B113)& =MIN(B2:B113). To find the actual most and least expensive, I used the Index function ie =INDEX(A2:A113,MATCH(MAX(B2:B113),B2:B113,0))& =INDEX(A2:A113,MATCH(MIN(B2:B113),B2:B113,0))

Used a pivot table to illustrate the relationship between the average of rating and discount across the discount category. Pivot table also used to show relationship between product count and rating categories

Trends & Relationship Analysis
Made 2 charts:

1st chart is discount against the number of reviews

2nd chart is ratings against the number of reviews

Product Performance Analysis
To return 10 products of the highest discounts, I used the Large & Row functions. LARGE($D$2:$D$113, k) gives the kth largest number in that range. k must be a positive integer.
ROW() returns the row number of the cell where the formula sits. For example, if the formula is in row 2, ROW() returns 2.
ROW()–1 shifts the rank. If the formula is in row 2, ROW()–1 becomes 1, so the formula returns the largest value. In row 3 it becomes 2, returning the second largest, and so on. =LARGE($D$2:$D$113,ROW()-1)

To visualise the actual products on a different column, I used the Index & Match functions. MATCH(I2,$D$2:$D$113,0) searches for the exact value in cell I2 inside D2:D113. It returns the position number where the match occurs, for example 7 if the match is the seventh item in the range.
INDEX($A$2:$A$113, position) takes that position and returns the value from A2:A113 at the same position. If MATCH returned 7, INDEX returns the seventh item of A2:A113. =INDEX($A$2:$A$113,MATCH(I2,$D$2:$D$113,0))

The same was done for top 10 products on new columns ie `=LARGE($F$2:$F$113,ROW()-1) & =INDEX($A$2:$A$113,MATCH(K2,$F$2:$F$113,0))

Created a pivot table comparing the products with max and min ratings. Used the value filter to show a representation of top and bottom 5 products by rating.

Created a pivot table with the Discount Category column containing high discount and low discount. Added the averages of Ratings and Reviews to visualise the differences

Dashboard
Overview
Formula for the Total Products =COUNTA(Excel_jumia!A2:A113)
Formula for Average Rating =AVERAGE(Excel_jumia!G2:G113)
Formula for Average Discount percentage =AVERAGE(Excel_jumia!D2:D113)
Formula for Total number of Reviews =COUNTIF(Excel_jumia!F2:F113,">0")

Product Performance
Created 3 pivot tables

Products against the average ratings for each product

Products against the sum of reviews for each product

Products against the max discount percentage of each product
For each table, filtered the top 5. From the top 5 tables, came up with charts for each table.

Trend Analysis
Created a pivot table. Reviews on the rows and Discounts on value as average. Grouped reviews in 10s. Created a column chart from these discounts vs reviews

Created a pivot table to represent rating vs reviews; Ratings on row and count of reviews on value. Line graph of ratings vs reviews

Product Categories
Created a pivot table to show the breakdown of products against discount category & rating category. This showed the count of products against the discount category and the Rating category

DASHBOARD
Created an interactive dashboard on a new worksheet
Added slicers for product and discount category.
