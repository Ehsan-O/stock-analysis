# **Stock Analysis**  


## **Overview of Project**
in this project we are helping a client to perform analysis on some stock data  


### **Purpose**
Our client wants to analyze the stock data of 12 companies in 2017 and 2018 to find how actively each company's stock was traded in these to years and to calculate yearly return of each companys stock  


## **Results**  

### **VBA code**
We have been provided with an excel file containing the information we need for the analysis. The data is stored in two sheets one for 2017 and another for 2018 here is what the data looks like  

<img src="/scr-shots-stock/preview.png">  

We wrote a "VBA" code in order to analyze the data. to find how actively companies were trading stocks each year, we need to add up values for the number of stocks traded in each day (this is the value in olumn "H" of data sheets under the name: "Volume") for a particular Ticker during that year. Then to calculate yearly return for each Ticker we should identify the first closing price (at the begining of the year) and the last closing price (at the end of the year) of the stock (for each ticker) then we can calculate the yearly return using the formula below:  

Yearly Return = (last closing price / first closing price) - 1  

We have 12 tickers, this means that our code should go through the data find the information we need for each of these tickers, perform the desired operations, and output the result. To do so an array is defined in the code to store the name of these 12 tickers:  

```vb
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

Then we used nested loops and if statements to go through all the rows of the data sheet for each ticker in the array and pick the information we need from the cells if it belongs to that ticker and output the result in a new sheet we named "All Stocks Analysis", to make the code more flexible we deciced to take the desired year of the analysis from the user:  

```vb
yearValue = InputBox("What year would you like to run the analysis on?")
```


```vb
    rowstart = 2
    'rowEnd code taken from <https://stackoverflow.com/questions/18088729/row-count-where-data-exists> and finds the final row in the data sheet
    rowend = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through the tickers.
    
    For i = 0 To 11
        
        ticker = tickers(i)
        'set the initial value for the total volume of the current tecker to zero
        totalVolume = 0
    
        'Loop through rows in the data.
        Worksheets(yearValue).Activate
            
            For j = rowstart To rowend
                
                'Find the total volume for the current ticker.
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
                
                'Find the starting price for the current ticker.
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
                
                'Find the ending price for the current ticker.
                  If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
                
            Next j
            
        'Output the data for the current ticker .
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        
    Next i
```
To evaluate the performance of the code we then added a few more lines to see how long the code takes to execute. For tha,t we used "Timer" function the code will record the timer's value once right after it takes the year from the user, assining it to a "startTime" variable and another time after it shows the result in the output sheet (after the "Next i" in the above code) and assigns it to an "endTime" variable then simply by subtracting "startTime" from "endTime" we will have the runtime of the code and we can show it in a message box:  

```vb
    Dim startTime As Single
    Dim endTime As Single
    

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
```

```vb
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

```
In the images below, we can see the result of running the cod for each year

