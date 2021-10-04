# Stock Analysis

## Overview of Project
The purpose of this project was to create code to help Steve analyze the stock market data from 2017 and 2018 to better inform his parents' investment decisions.

## Results
![2017 results](./resources/VBA_Challenge_2017a.png)
![2018 results](./resources/VBA_Challenge_2018a.png)

### Analysis
First, we needed to create an array of all the tickers.
```
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

```


![2017 original time](./resources/2017_original-time.png)
![2017 refactored time](./resources/2017_refactored-time.png)

![2018 original time](./resources/2018_original-time.png)
![2018 refactored time](./resources/2018_refactored-time.png)