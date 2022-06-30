# VBA-challenge
Bootcamp Assignment#2 
Projext Name: The VBA of Wall Street

#Description:
	In this project, the VBA scripting is used to analyze generated stock market data from 2018 to 2020.
It generates yearly change and total stock volume of every stock. Also,the stocks with greatest %
increase, % decrease and total volume is showed for bonus information.

#Date: 06.25.2022
#Author: I Ju SU selina.iju@gmail.com
#How to start working with this?
 1. Download the Mutiple_year_stock_data.vbs.bas from https://github.com/sesu0722/VBA-challenge/blob/main/Multiple_year_stock_data.vbs.bas
 2. Open the excel workbook Mutiple_year_stock_data.xlsm (*the file is too large to upload to GitHub)
 3. run the VBA script on the excel workbook.
 4. It will automaticly generate the tickers, yearly change, percent change and total stock volume for each year/wooksheet,and highlight the       positive/negetive changes in green/red.
 5. An array shows stocks with greastest % changes and volume is created for bonus.
 6. Screenshots of the worksheets can be download from https://github.com/sesu0722/VBA-challenge/blob/main/ScreenShotForEachYear.docx
	
#Coding
 1. Ticker Symbol: For loop and If/Else are used to loop through the rows to return the ticker symbol when there is a difference
 2. Yearly Change: I use the If /Or to return the open price for dtae 2018/01/02, 2019/01/02 and 2020/01/02 and save it in column "M"(hided).
                   The close price is returned by the For loop (when there is a different ticker) and saved in column "N"(hided).
		   Yearly Change = Close price - Open price (column "N"-"M"). 
		   Use the If/Else and Interior.ColorIndex to color the postive and negetive into green and red.
 4. Total Stock volume: this is calculated by the For loop to sum the total volume for the ticker until there is a different ticker, 
                        then re-set to sum the total volumn for new ticker.
 5. Percent Change: Percent change = Yearly change / total stock volume and NumberFormat to 2 decimal and %
 6. Greatest % increase, decrease and total volume = I use WorksheetFunction Max and Min to find the value.
