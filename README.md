# Stock Market Analysis 

Data analysis using VBA scripting and Excel.

## Description

This repository is designed to analyze stock market data in Excel, using VBA scripting. Stock market data for 2014-2016 is stored in the Resources folder (Multiple_year_stock_data.xlsl), with each year of data on a separate tab. The VBA script (KatlinBowman_VBAChallenge.bas) loops through each tab and creates new columns that store the: 
- The ticker symbol.
- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.

The script also uses conditional formatting to highlight positive change in green and negative change in red.
<p align="center">
  <img src="https://user-images.githubusercontent.com/74067302/146248616-764ad697-aa6b-4676-b297-6fb6c398115a.png" alt="Dashboard Image"/>
</p>

## Getting Started

### Technologies Used 

* Microsoft Excel
* VBA scripting

### Installing

* Clone this repository to your desktop.
* Navigate to the Resources directory and open Multiple_year_stock_data.xlsx.
* Import KatlinBowman_VBAChallenge.bas into the Visual Basics editor.

## Author

Katlin Bowman, PhD

E: klbowman@ucsc.edu

[LinkedIn](https://www.linkedin.com/in/katlin-bowman/)
