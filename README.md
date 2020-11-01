# **Stock Analysis**  


## **Overview of Project**
in this project we are helping a client to perform analysis on some stock data  


### **Purpose**
Our client wants to analyze the stock data of 12 companies in 2017 and 2018 to find how actively each company's stock was traded in these to years and to calculate yearly return of each companys stock  


## **Results**  

We have been provided with an excel file containing the information we need for the analysis. The data is stored in two sheets one for 2017 and another for 2018 here is what the data looks like  

<img src="/scr-shots-stock/preview.png">  

We wrote a "VBA" code in order to analyze the data. to find how actively companies were trading stocks each year, we need to add up values for the number of stocks traded in each day (this is the value in olumn "H" of data sheets under the name: "Volume") for a particular Ticker during that year. Then to calculate yearly return for each Ticker we should identify the first closing price (at the begining of the year) and the last closing price (at the end of the year) of the stock (for each ticker) then we can calculate the yearly return using the formula below:  

Yearly Return = (last closing price / first closing price) - 1  

We have 12 tickers, this means that our code should go through the data find the information we need for each of these tickers, perform the desired operations, and output the result. To do so an array was define in the code to store the name of these 12 tickers:  

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



