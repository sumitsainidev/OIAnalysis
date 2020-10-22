# Open Interest(OI) Analysis

This python script helps in analysis of NSE Nifty Futures and Options Open Interest data.
## Nifty Future OI Analysis
### 1. Five Min Nifty Futures OI analysis
Script fetches Nifty Futures data every 5 minutes from the official NSE website (https://www1.nseindia.com/) and saves it inside OiAnalysis.xlxs in FiveMin sheet and gives BUY/SELL signal based on change in OI and change in LTP along with OI interpretaions (i.e Long Buildup, Short Buildup, Long Unwinding and Short Covering)
<img src="Images/Future5MinOI.PNG">

### 2. Fifteen Min Nifty Future OI analysis
Script fetches Nifty Futures data every 15 minutes from the official NSE website (https://www1.nseindia.com/) and saves it inside OiAnalysis.xlxs in FiftMin sheet and gives BUY/SELL signal based on change in OI and Change in LTP along with OI interpretaions (i.e Long BuildUp, Short Buildup, Long Unwinding and Short Covering)
<img src="Images/Future15MinOI.PNG">

## Nifty Option Chain OI Analysis
Script fetches Nifty Option Chain data in every 5 minutes for every strike price from official NSE website (https://www1.nseindia.com/) and saves it inside OptionChain.xlsm. This data is saved in Master sheet which is further used by Pivot Table sheet to filter out data based on strike prices. All the Call/Put data is represented on Dashboard sheet in form of OI vs LTP graph. 

### 1. Dashboard
<img src="Images/OI%20vs%20LTP%20Graph.PNG">

### 2. Pivot Table
<img src="Images/Pivot%20Table.PNG">

### 3. Master Sheet
<img src="Images/Master.PNG">
