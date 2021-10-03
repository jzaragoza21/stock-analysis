# **Stock Analysis With Excel VBA** 
Click here to view the Excel file: [VBA Challenge - Stock Aanalysis](https://github.com/jzaragoza21/stock-analysis/blob/main/VBA_challenge.xlsm)

## **Overview of Project**

### **Purpose**

The overall purpose of the challenge was to use our knowledge of VBA and use the starter code provided to refractor the Module2_VBA_Script. As we did in the module, we will do this to loop through and collect all of the 2017 and 2018 stock analysis data one time. However, our goal this time will be to see whether refractoring the code made the VBA_script run faster this time. In essence, there are many times in which we have to increase efficiency in writing code on the job and our objective here is to determine whether refractoring does in fact increase efficiency.   

### **Data**

The stock analysis data that is provided includes two sheets of data on twelve different stocks for the years 2017 and 2018. Additionally, the data is broken down by a ticker index, when the stock was issued, the stock opening and closing values, the stock high and low values, its adjusted close value and the stock total daily volume. The objective with the data is to acquire the tickers, total daily volume and ultimately the return percentage. 

## **Results**

### **Analysis**

In setting up this refractoring process, I first had to copy the original VBA script so that I could set up input box, format my output sheet, activate the righ worksheet and initialize my array of tickers. Thereafter, I copied the challenge steps into my new module and macro and began instering the new refractor code. The following is a breakdown of the refractor code:

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

## **Summary**

### **Pros and Cons of Refactoring in General**

Overall, refactoring code has the benefit of being more organized and clean. As the time ran demostrates below, it can also save you and your organization considerable time and in most cases, that will also help you save money. While we limited experience with these languages, refactoring also seems to help simplifying the code and removing some of the complexities. And simplified, cleaner code can help with debugging. As far as disadvantages go, it appears that refactoring can be prone to code defects and, at times, slow down developement.   

### **Advantages of Refactoring VBA Code**

After going through this refactoring process, the advantages are clear. In particular, you save in coding run times. The two images below democtrate this in that the run times for 2017 and 2018 were .33 and .31 of a second respectively. The run times prior to inserting the refactored code was over a half second, clocking in around .53 of a second. This is a substantial improvement of a 42 percent drop in time to run the code. I have to imagine with much larger datasets this could save considerable time and increase efficiency for an orgnization, business or a programmer.   


![This is an image](https://github.com/jzaragoza21/stock-analysis/blob/main/VBA_Resources/VBA_Challenge_2017.png)

![This is an image](https://github.com/jzaragoza21/stock-analysis/blob/main/VBA_Resources/VBA_Challenge_2018.png)
