Introduction:
1.1	Purpose:

Rising costs of living are putting pressure on household budgets in New Zealand. To understand trends across essential spending categories, Statistics New Zealand maintains the HLPI.

This company generates useful reports by analyzing HLPI data to provide insights into price movements over time. Policymakers then actively use these reports to frame welfare measures. 

1.2	The dataset:

The Excel sheet contains the data on expenditures by different households in different types of categories like food, housing, and transportation. This dataset has collective information from the year 2008 to 2023, Along with increasing and decreasing tens in the expenditure index.

The dataset contains 11 columns with more than 50000 rows in a single excel sheet. Each of these columns are used to filter and extract data. 

•	The column hlpi_name contains a list of the different types of households and series_ref contains unique codes for these households.
•	The quarter column contains the year and quarter of the year for that saved data. 
•	The nzhec column contains the unique code for the expenditure categories stored in the nzhec_name column, and the short form of these names are stored in nzhec_short. 
•	The level column identifies whether the expenditure category is a group or a subgroup. 
•	The purchase_perFamily column contains the information on expenditure cost by a family per quarter, and change.q and change.a these increases and decreases of inflation as an index.
From the above columns hlpi_name, quarter, and nzhec_name are used to filter data from the dataset while purchase_perFamily, change.q and change.a are used to extract the data and utilize them for the report. All this data is stored in one Excel sheet with thousands of entries and unfiltered values so working with it manually is tedious.


1.3	Data Processing:

This data is processed by the analyst in the company to generate charts and reports to help make the policies. The downside of their method is that all the data is stored in one Excel sheet, and to make a report by year, category, or household type, they manually extract data from the sheet, create calculations, and finally generate a report. All this work is time-consuming and prone to errors.

Solution:
2.1	Description of the solution:
A Python script was developed using openpyxl and pyinputplus libraries to automatically generate analysis reports from this HLPI data sheet. The following are the steps that are involved in generating a report: 
a)	The script reads the data from an Excel file containing all the records using openpyxl.

b)	A function first asks the user to select a report type like only household type, only expenditure type or both types. 

c)	Functions are used to take the inputs from the employees using the pyinputplus library. These inputs include year range, household type, and expenditure category. These inputs are based on the type of report they want to create.

d)	After all the inputs are taken from the user, a new Excel file is created with the appropriate name and heading inserted into it all according to the user inputs. To this, a user-defined function is used.

e)	Thereafter, the newly created file is used by another user-defined function to insert data into the file. All the data that is inserted into the file is filtered using loops and if-else conditions.

f)	These loops also calculate average, maximum, and minimum expenditure data from the dataset and these values are stored in variables.

g)	Stored data is passed to another user-defined function which inserts these values into the new Excel sheet along with the date and time of creation of this report. Date and time retrieved using datetime module from the Python in-built libraries.

h)	After all the data has been inserted new excel sheet is save in same folder as the application and can be accessed by the analyst



2.2	Using analyzed data
All the filtered data from the data is used to construct charts and tables to analyze the trends in expenditure. 
a)	Tables are created to show statistics like average expenditure and inflation rate.
b)	Bar charts are used to show expenditure trends over time for different household groups.
c)	Tends are studied where inflations are at the highest and lowest points to plot graphs.
d)	Overlayed trendlines on charts are used to see patterns and predict future expenditure trends.
e)	Some cells are conditionally formatted to highlight the categories of high inflation impacts.

2.3	Implementing Data
After all the filtered data is analyzed then it will be used by policymakers to use this data to create welfare measures. This can be done in the following ways:

a)	By Identifying categories with the highest expenditure index impacting low-income households. This helps target subsidies and price controls.
b)	By monitoring expenditure index trends over time across expenditure categories. This can help by increasing or decreasing the availability of that product.
c)	Analyzing the impact of policy measures from the highest and lowest point of expenditure index.
d)	Publish analyses publicly for transparency and feedback from the public.

Conclusion:
All in all, the application developed by me will provide substantial benefits over manual processes previously used for household expenditure analysis. By leveraging programming techniques, it removes tedious, repetitive tasks from the analysis process. This application increases the efficiency, quality, and effectiveness of the report as well as reducing the time and effort to build it. Overall, it reduces the manual burden on the analyst and delivers higher-quality information to the policymaker on time.
