# VBA - Challenge
## Aggregate and Analyze Stock Market Data
### *2014 - 2015 - 2016*

![](ReadMe_Images/stock-image2.jpg)

**The goal of the project is to loop through all stocks in each worksheet in a workbook and output the following:**
- Unique/Distinct Ticker Symbol
- Yearly change from opening price at the beginning of the year to the closing price at the end of the year
- The percent change from opening price at the beginning of the year to closing price at the end of the year
- The total stock volume of the stock

**Each worksheet in the workbook represents a year of stock data, and originating column headers set as follows:** [^1]
|---A1---|---B1---|---C1---|---D1---|---E1---|---F1---|---G1---|
|--------|--------|--------|--------|--------|--------|--------|
| ticker | date | open | high | low | close | vol |


**For a little extra bonus to the analysis, the following summary data is extracted as well:**
- Stock with the greatest percent increase
- Stock with the greatest percent decrease
- Stock with the most volume traded 

[^1]: *Columns beyond "G" must be unused and available prior to running macro.*
