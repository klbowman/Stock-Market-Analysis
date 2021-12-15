# Stock Market Analysis 

Data analysis using VBA scripting and Excel.

## Description

This repository is designed to analyze stock market data in Excel, using VBA scripting. The script (KatlinBowman_VBAChallenge.bas)  
### Data tab
- Conditional formatting is used in the **State** column to color code projects as successful (green), failed (red), canceled (yellow), or live (grey). 
- The **Percent Funded** column uses a formula to uncover how much money a campaign made to reach its initial goal, and conditional formatting is used to color code funding amounts (0-99; red, 100-199; green, >200; blue).
- The **Average Donation** column uses a formula to uncover how much each backer for the project paid on average.
- The **Category** and **Sub-Category** columns use formulas to split the **Category and Sub-Category** column into two parts.
- The **Date Created Conversion** column uses a formula to convert the Unix timestamp data in the **launched_at** column into Excel's date format.
- The **Date Ended Conversion** column uses a formula to convert the Unix timestamp data in the **deadline** column into Excel's date format.
<p align="center">
  <img src="https://user-images.githubusercontent.com/74067302/146103340-29efbfab-be63-4fef-a516-22c16e8d376d.png" alt="Dashboard Image"/>
</p>

### Pivot 1_Category tab
- This sheet contains a pivot table that analyzes the **Data** worksheet to count how many campaigns were successful, failed, canceled, or are currently live per category. The stacked column pivot chart can be filtered by country.

### Pivot 2_Subcategory tab
- This sheet contains a pivot table that analyzes the **Data** worksheet to count how many campaigns were successful, failed, or canceled, or are currently live per sub-category. The stacked column pivot chart can be filtered by country and parent-category.
<p align="center">
  <img src="https://user-images.githubusercontent.com/74067302/146103546-87260a18-21d7-4b89-923b-cc47ab3e095c.png" alt="Dashboard Image"/>
</p>

### Pivot 3_Dates tab
- This sheet contains a pivot table with a column of state, rows of Date Created Conversion, values based on the count of state, and filters based on parent category and Years, and a pivot chart line graph that visualizes this table.
<p align="center">
  <img src="https://user-images.githubusercontent.com/74067302/146104544-17c9e503-a770-4b64-a71a-2ee1981f542e.png" alt="Dashboard Image"/>
</p>

### Bonus tab
- This sheet uses formulas to calculate the number and percentage of projects that were successful, failed, or canceled, and the total number of projects for 12 ranges of monetary goals. 
- The line chart graphs the relationship between a goal's amount and its chances at success, failure, or cancellation.
<p align="center">
  <img src="https://user-images.githubusercontent.com/74067302/146105322-519bb1d6-40da-4ad7-ba08-6bff8f1ef527.png" alt="Dashboard Image"/>
</p>

### Statistical Analysis tab
- This sheet displays the number of backers of successful and unsuccessful campaigns.
- Formulas are used to evaluate the following for successful campaigns and unsuccessful campaigns:
  - The mean number of backers.
  - The median number of backers.
  - The minimum number of backers.
  - The maximum number of backers.
  - The variance of the number of backers.
  - The standard deviation of the number of backers.

## Insights from Data Analysis

### Conclusions
- The majority of successful Kickstarter campaigns are in theater, music, film & video categories. 
- The greatest number of successful Kickstarter campaigns were launched during the month of May. 
- The following subcategories have a 100% success rate: classical music, documentary, electronic music, hardware, metal, nonfiction, pop, radio & podcasts, rock, shorts, small batch, tabletop games, television. 
- The following subcategories have a 100% failure rate: animation, children's books, drama, fiction, food trucks, gadgets, jazz, mobile games, people, places, restaurants, translation, video games, web.

### Limitations
- The data in these plots is not normalized to reflect the proportion of successful and failed accounts relative to the total number of accounts. Some categories may have a greater number of successful accounts due to a higher number of submissions, rather than a higher funding success rate.
- This data analysis does not take into consideration other factors, such as funding goal and the number of backers, that may affect the outcome of the account.

## Getting Started

### Technologies Used 

* Microsoft Excel
* VBA scripting

### Installing

* Clone this repository to your desktop.
* Navitage to the home directory and open DataAnalysis.xlsx.

### Data Sources

* Web Roots: Kickstarter Datasets [Access Data](https://webrobots.io/kickstarter-datasets/)


## Authors

Katlin Bowman, PhD

E: klbowman@ucsc.edu

[LinkedIn](https://www.linkedin.com/in/katlin-bowman/)
