# Creat Pivot Table & Chart also Dashboard in Excel
## Bike Buyers Data Analysis

This repository contains a comprehensive data analysis of bike buyers. Follow the instructions below to replicate the analysis and create visualizations for better insights.

## Instructions

### Step 1: AutoFit Columns
1. Open **Data Sheet** (Bike_Buyers).
2. Go to **View Code**.
3. Select **Worksheet**.
4. Enter the following code:
   ```vba
   Cells.EntireColumn.AutoFit
   ```
5. Cancel Code Sheet.

### Step 2: Copy Data
1. Go to the **Data Sheet** (Bike_Buyers).
2. Copy all the data.

### Step 3: Prepare Working Sheet
1. Add a new worksheet and name it **Working Sheet**.
2. Paste the copied data into **Working Sheet**.
3. Remove duplicates in **Working Sheet**.

### Step 4: Clean Data
1. Press `Control+F` to open the Find and Replace dialog.
2. In **Marital Status** column:
   - Replace `M` with `Married`.
   - Replace `S` with `Single`.
3. In **Gender** column:
   - Replace `F` with `Female`.
   - Replace `M` with `Male`.
4. Change the **Income** column format to Currency (Dollar).

### Step 5: Define Age Bracket
1. In cell **M1**, enter the following formula:
   ```excel
   =IF(L2>54, "Old", IF(L2>=31, "Middle Age", IF(L2<31, "Adolescent", "Invalid")))
   ```
2. Drag down to apply the formula to the entire column.

### Step 6: Create PivotTable
1. Add a new worksheet and name it **PivotTable**.
2. In cell **A3**:
   - Go to **Insert** > **PivotTable**.
   - Select the **Data Sheet**.
   - Press `Control+A` to select all data, then click **OK**.
3. Drag and drop fields to create the desired PivotTable.

### Step 7: Create Charts
1. Go to **Recommended Charts** and select the desired chart.
2. Add a chart title and configure the axes as needed.
3. Here, I repeat this process to create three additional charts.

### Step 8: Prepare Dashboard
1. Add a new worksheet and name it **Dashboard**.
2. Copied and paste all the charts you want to display on the **Dashboard**.
3. Go to **View** and unselect **Gridlines**.
4. Select cells **A1:O4**:
   - Merge them.
   - Change the background color to blue.
   - Add a title like "Bike Sale Dashboard" in white, bold, font size 48.
5. Arrange the charts neatly below the title.

### Step 9: Add Slicer
1. Click on a graph.
2. Go to **Insert** > **Slicer**.
3. Select the desired variable (e.g., Marital Status), then click **OK**.
4. Click on the slicer:
   - Go to **Slicer** > **Report Connections**.
   - Select all PivotTables, then click **OK**.

## Conclusion
By following these steps, you can analyze bike buyer data, clean and format it, create insightful visualizations, and compile a comprehensive dashboard for easy data interpretation.

This "Bike Sales Dashboard" provides a comprehensive overview of bike purchasing behavior based on various demographic and geographic factors. Here's a detailed interpretation of each component:

1. **Filters (Left Side)**
   - **Marital Status**: Allows filtering the data by marital status (Married or Single).
   - **Region**: Allows filtering the data by geographic region (Europe, North America, Pacific).
   - **Education**: Allows filtering the data by the level of education (Bachelors, Graduate Degree, High School, Partial College, Partial High School).

2. **Customer Commute (Top Left)**
   - **Chart**: Displays the count of purchased bikes based on the commute distance (0-1 Miles, 1-2 Miles, 2-5 Miles, 5-10 Miles, More than 10 Miles).
   - **Insights**: It appears that commute distance does not significantly impact the decision to purchase a bike, as the lines for 'Yes' and 'No' are relatively close for all distances.

3. **Customer Age Bracket (Top Right)**
   - **Chart**: Displays the count of purchased bikes by age bracket (Adolescent, Middle Age, Old).
   - **Insights**: Middle-aged customers are the most likely to purchase bikes, followed by adolescents, with older customers being the least likely.

4. **Income Per Purchase (Bottom Left)**
   - **Chart**: Displays the average income of customers based on their gender and whether they purchased a bike.
   - **Insights**: Customers who purchased bikes generally have higher incomes. Males have higher average incomes than females, and among both genders, those who purchased bikes have higher incomes compared to those who did not.

5. **Customer Age (Bottom Right)**
   - **Chart**: Displays the count of purchased bikes by individual age.
   - **Insights**: The age distribution shows that the highest number of bike purchases occur between the ages of 25 and 45, with a noticeable peak around age 35. The likelihood of purchasing a bike decreases significantly after age 45.

### Summary
The dashboard indicates that middle-aged customers, particularly those with higher incomes, are more likely to purchase bikes. Commute distance appears to have minimal impact on bike purchasing decisions. Regional and educational filters provide additional layers for more detailed analysis, allowing users to segment the data further to identify specific trends and patterns within different demographics and geographic locations.

