# VBA-challenge Module 2 Bootcamp

## Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data

## Before You Begin
Create a new repository for this project called VBA-challenge. Do not add this assignment to an existing repository.

Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.

## Files
Download the files to help you get started: Module 2 Challenge files: alphabetical_testing.xlsx and Multiple_year_stock_data.xlsx

## Instructions
* Create a script that loops through all the stocks for one year and outputs the following information:

1- The ticker symbol

2- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

3- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

4- The total stock volume of the stock. The result should match the following image:

* Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

## NOTE
* Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

## Other Considerations
* Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

## Requirements
* Retrieval of Data (20 points)
* The script loops through one year of stock data and reads/ stores all of the following values from each row:

1-ticker symbol (5 points)

2-volume of stock (5 points)

3-open price (5 points)

4-close price (5 points)

* Column Creation (10 points)
* On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:

1- ticker symbol (2.5 points)

2- total stock volume (2.5 points)

3- yearly change ($) (2.5 points)

4- percent change (2.5 points)

5- Conditional Formatting (20 points)
Conditional formatting is applied correctly and appropriately to the yearly change column (10 points)

Conditional formatting is applied correctly and appropriately to the percent change column (10 points)

## Calculated Values (15 points)
All three of the following values are calculated correctly and displayed in the output:

1- Greatest % Increase (5 points)

2- Greatest % Decrease (5 points)

3- Greatest Total Volume (5 points)

## Looping Across Worksheet (20 points)
* The VBA script can run on all sheets successfully.

## GitHub/GitLab Submission (15 points)
* All three of the following are uploaded to GitHub/GitLab:

1- Screenshots of the results (5 points)

2- Separate VBA script files (5 points)

3- README file (5 points)

## Submission
* provide the URL of your GitHub repository for grading.

## Note: 
* I used  https://officetuts.net/excel/vba/find-the-maximum-and-minimum-value-in-the-range-in-vba/ in Script1 and Script2 to get the code to find the maximum and minimum percent change.
* I used https://excelchamps.com/vba/autofit/ to get the code to Autofit summary table and data.
