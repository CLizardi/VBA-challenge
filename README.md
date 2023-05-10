# VBA Challenge
![aditya-vyas-7ygsBEajOG0-unsplash](https://user-images.githubusercontent.com/52866379/237005408-fb5cfee8-62e9-41b9-914e-2e47d187bff9.jpg)

# Introduction
Welcome to the VBA Challenge! In this project, I used VBA scripting to analyze stock market data. The goal was to loop through all the stocks for one year and output various information, including ticker symbol, yearly change, percentage change, and total stock volume.

# Project Overview
In this project, I used VBA scripting to automate the analysis of stock market data. The data was organized by date, ticker symbol, opening price, highest price, lowest price, closing price, and volume. The goal was to analyze the data for each stock and output the following information: ticker symbol, yearly change from opening price to closing price, percentage change from opening price to closing price, and total stock volume.

# What I Did
In this project, I performed the following tasks:

* Created a VBA script to loop through all the stocks for one year and output the required information.
* Used conditional formatting to highlight positive changes in green and negative changes in red.
* Calculated the stock with the greatest percent increase, greatest percent decrease, and greatest total volume.
* Created a summary report for each worksheet with the analyzed data and results.

# Tools Used
The following tools were used in this project:

* VBA scripting - for automating the analysis of stock market data
* Excel - for organizing and storing the data and creating the summary report
* GitHub - for version control and collaboration
* VSCode - for writing and editing markdown files

# What I Learned
Through this project, I gained experience in the following:

* VBA scripting and automation
* Data analysis and manipulation using Excel
* Conditional formatting and cell color coding
* Generating summary reports and data visualizations
* Collaborating on GitHub and using VSCode for writing and editing markdown files.

# Conclusion
In conclusion, the VBA Challenge allowed me to gain hands-on experience in automating the analysis of stock market data using VBA scripting. I was able to analyze the data for each stock and output the required information, including yearly change, percentage change, and total stock volume. This project also helped me develop my skills in data manipulation, conditional formatting, and generating summary reports.   
   
    Sub stock_analysis():
    Dim lastRow As Long
    Dim totalVolume As LongLong
    Dim openPrice As Double
    Dim closePrice As Double
    Dim ticker As String
    Dim dollarsChange As Double
    Dim percentChange As Double
    Dim summaryRow As Integer

    Dim biggestGain As Double
    Dim biggestGainTicker As String
    Dim biggestLoss As Double
    Dim biggestLossTicker As String
    Dim mostVolume As Double
    Dim mostVolumeTicker As String
    
    For Each ws In Worksheets
        ws.Activate

        summaryRow = 2
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        openPrice = Cells(2, 3).Value
        totalVolume = 0
        
        biggestGain = 0
        biggestLoss = 0
        mostVolume = 0
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        For currentRow = 2 To lastRow
            totalVolume = totalVolume + Cells(currentRow, 7)
            
            If Cells(currentRow + 1, 1).Value <> Cells(currentRow, 1).Value Then
            
                ticker = Cells(currentRow, 1).Value
                closePrice = Cells(currentRow, 6).Value
                dollarsChange = closePrice - openPrice
                percentChange = dollarsChange / openPrice
                
                Cells(summaryRow, 9).Value = ticker
                Cells(summaryRow, 10).Value = dollarsChange
                Cells(summaryRow, 11).Value = percentChange
                Cells(summaryRow, 12).Value = totalVolume
                
                If dollarsChange >= 0 Then
                    Cells(summaryRow, 10).Interior.ColorIndex = 4
                Else
                    Cells(summaryRow, 10).Interior.ColorIndex = 3
                End If
                
                If percentChange > biggestGain Then
                    biggestGain = percentChange
                    biggestGainTicker = ticker
                End If
                
                If percentChange < biggestLoss Then
                    biggestLoss = percentChange
                    biggestLossTicker = ticker
                End If
                
                If totalVolume > mostVolume Then
                    mostVolume = totalVolume
                    mostVolumeTicker = ticker
                End If
                
                summaryRow = summaryRow + 1
                
                openPrice = Cells(currentRow + 1, 3).Value
                
                totalVolume = 0
            End If
        Next currentRow
        
        Range("K2:K" & summaryRow).Style = "Percent"
        
        Range("P2").Value = biggestGainTicker
        Range("Q2").Value = biggestGain
        Range("P3").Value = biggestLossTicker
        Range("Q3").Value = biggestLoss
        Range("P4").Value = mostVolumeTicker
        Range("Q4").Value = mostVolume
    Next ws
    End Sub
