from datetime import datetime, time
from time import sleep
import requests
import xlwings as xw

#Function to add Strike Price sheet in OptionChain.xlsx
def addStrikePriceSheet(wb,strikePrice):
    wb = xw.Book('OptionChain.xlsx')
    wb.sheets.add(name=str(strikePrice))

#Function add different option strike data in respective option strike sheet
def addOptionDataInSheet(wb,currTime,strikePrice,ceLTP,ceOI,peLTP,peOI):
    wb = xw.Book('OptionChain.xlsm')
    wb_sheet = wb.sheets['Master']
    if wb_sheet.cells(2, 1).value == None:
        lastRow = 1
    else:
        lastRow = wb_sheet.range('A1').end('down').row
        print(lastRow)

    wb_sheet.cells(lastRow + 1, 1).value = currTime
    wb_sheet.cells(lastRow + 1, 2).value = ceLTP
    wb_sheet.cells(lastRow + 1, 3).value = ceOI
    wb_sheet.cells(lastRow + 1, 4).value = strikePrice
    wb_sheet.cells(lastRow + 1, 5).value = peLTP
    wb_sheet.cells(lastRow + 1, 6).value = peOI
    #Color strike Price
    wb_sheet.range(lastRow+1, 4).color = (255, 255, 0)

#Function make excel sheet if not present already
def makeOptionChainFile(opt,currTime):
    wb = xw.Book('OptionChain.xlsm')
    optionData = opt["filtered"]["data"]

    #Create Strike Price worksheet
    for optionRow in optionData:
        strikePrice = optionRow["CE"]["strikePrice"]
        ceLTP = optionRow["CE"]["lastPrice"]
        ceOI = 75*optionRow["CE"]["openInterest"]
        peLTP = optionRow["PE"]["lastPrice"]
        peOI = 75*optionRow["PE"]["openInterest"]
        addOptionDataInSheet(wb,currTime,strikePrice,ceLTP,ceOI,peLTP,peOI)
    wb.save()

#Function to add data in optionchain.xlsm
def putOptionChainData(opt,currTime):
    wb = xw.Book('OptionChain.xlsm')
    optionData = opt["filtered"]["data"]
    for optionRow in optionData:
        strikePrice = optionRow["CE"]["strikePrice"]
        ceLTP = optionRow["CE"]["lastPrice"]
        ceOI = 75*optionRow["CE"]["openInterest"]
        peLTP = optionRow["PE"]["lastPrice"]
        peOI = 75*optionRow["PE"]["openInterest"]
        addOptionDataInSheet(wb,currTime,strikePrice,ceLTP,ceOI,peLTP,peOI)
    wb.save()

#Function to add 5 min data inside OiAnalysis.xlsx
def putInExcel5Min(currTraded, futOI, tradedVolCon, callOI, putOI,currTime):

    wb =xw.Book('OiAnalysis.xlsx')
    wb_sheet = wb.sheets[0]
    lastRow = wb_sheet.range('A1').end('down').row
    last_time = wb_sheet.cells(lastRow, 1).value
    last_timee = last_time.split('-')
    time_range = last_timee[1] + "-" + currTime
    last_traded_price = wb_sheet.cells(lastRow, 2).value
    changeInLTP = currTraded - last_traded_price
    last_Fut_OI = wb_sheet.cells(lastRow,5).value
    changeInFutOI = futOI - last_Fut_OI
    last_Call_OI = wb_sheet.cells(lastRow, 7).value
    changeInCallOI = callOI - last_Call_OI
    last_Put_OI = wb_sheet.cells(lastRow,9).value
    changeInPutOI = putOI - last_Put_OI
    OiInter = ""
    signal = ""
    redFill = (255,204,203)
    greenFill = (144,238,144)

    if changeInLTP > 0 and changeInFutOI > 0:
        OiInter = "Long Buildup"
    elif changeInLTP < 0 and changeInFutOI > 0:
        OiInter = "Short Buildup"
    elif changeInLTP < 0 and changeInFutOI < 0:
        OiInter = "Long Unwinding"
    elif changeInLTP > 0 and changeInFutOI < 0:
        OiInter = "Short Covering"

    if changeInLTP>0 and changeInFutOI>0 and changeInPutOI>0 and changeInCallOI<0:
        signal = "Buy"
    elif changeInLTP<0 and changeInFutOI>0 and changeInPutOI<0 and changeInCallOI>0:
        signal = "Sell"
    elif changeInLTP<0 and changeInFutOI<0 and changeInPutOI<0 and changeInCallOI>0:
        signal = "Sell"
    elif changeInLTP>0 and changeInFutOI<0 and changeInPutOI>0 and changeInCallOI<0:
        signal = "Buy"

    #color filling start
    if changeInLTP>0:
        wb_sheet.range(lastRow + 1, 3).color = greenFill
    elif changeInLTP<0:
        wb_sheet.range(lastRow + 1, 3).color = redFill

    if changeInFutOI>0:
        wb_sheet.range(lastRow + 1, 6).color = greenFill
    elif changeInFutOI<0:
        wb_sheet.range(lastRow + 1, 6).color = redFill

    if changeInCallOI>0:
        wb_sheet.range(lastRow + 1, 8).color = greenFill
    elif changeInCallOI<0:
        wb_sheet.range(lastRow + 1, 8).color = redFill

    if changeInPutOI>0:
        wb_sheet.range(lastRow + 1, 10).color = greenFill
    elif changeInPutOI<0:
        wb_sheet.range(lastRow + 1, 10).color = redFill

    if signal=="Buy":
        wb_sheet.range(lastRow + 1, 12).color = greenFill
    elif signal=="Sell":
        wb_sheet.range(lastRow + 1, 12).color = redFill

    wb_sheet.cells(lastRow+1, 1).value = time_range
    wb_sheet.cells(lastRow+1, 2).value = currTraded
    wb_sheet.cells(lastRow+1, 3).value = changeInLTP
    wb_sheet.cells(lastRow+1, 4).value = tradedVolCon
    wb_sheet.cells(lastRow+1, 5).value = futOI
    wb_sheet.cells(lastRow+1, 6).value = changeInFutOI
    wb_sheet.cells(lastRow+1, 7).value = callOI
    wb_sheet.cells(lastRow+1, 8).value = changeInCallOI
    wb_sheet.cells(lastRow+1, 9).value = putOI
    wb_sheet.cells(lastRow+1, 10).value = changeInPutOI
    wb_sheet.cells(lastRow + 1, 11).value = OiInter
    wb_sheet.cells(lastRow + 1, 12).value = signal
    wb.save()
    print(time_range,"    ",currTraded,"    ",changeInLTP,"     ",OiInter,"      ",signal)


#Function to add 15 min data inside OiAnalysis.xlsx
def putInExcel15Min(currTraded, futOI, tradedVolCon, callOI, putOI,currTime):
    wb = xw.Book('OiAnalysis.xlsx')
    wb_sheet = wb.sheets[1]
    lastRow = wb_sheet.range('A1').end('down').row
    last_time = wb_sheet.cells(lastRow, 1).value
    last_timee = last_time.split('-')
    time_range = last_timee[1] + "-" + currTime
    last_traded_price = wb_sheet.cells(lastRow, 2).value
    changeInLTP = currTraded - last_traded_price
    last_Fut_OI = wb_sheet.cells(lastRow, 5).value
    changeInFutOI = futOI - last_Fut_OI
    last_Call_OI = wb_sheet.cells(lastRow, 7).value
    changeInCallOI = callOI - last_Call_OI
    last_Put_OI = wb_sheet.cells(lastRow, 9).value
    changeInPutOI = putOI - last_Put_OI
    OiInter = ""
    signal = ""
    redFill = (255, 204, 203)
    greenFill = (144, 238, 144)

    if changeInLTP > 0 and changeInFutOI > 0:
        OiInter = "Long Buildup"
    elif changeInLTP < 0 and changeInFutOI > 0:
        OiInter = "Short Buildup"
    elif changeInLTP < 0 and changeInFutOI < 0:
        OiInter = "Long Unwinding"
    elif changeInLTP > 0 and changeInFutOI < 0:
        OiInter = "Short Covering"

    if changeInLTP > 0 and changeInFutOI > 0 and changeInPutOI > 0 and changeInCallOI < 0:
        signal = "Buy"
    elif changeInLTP < 0 and changeInFutOI > 0 and changeInPutOI < 0 and changeInCallOI > 0:
        signal = "Sell"
    elif changeInLTP < 0 and changeInFutOI < 0 and changeInPutOI < 0 and changeInCallOI > 0:
        signal = "Sell"
    elif changeInLTP > 0 and changeInFutOI < 0 and changeInPutOI > 0 and changeInCallOI < 0:
        signal = "Buy"

    # color filling start
    if changeInLTP > 0:
        wb_sheet.range(lastRow + 1, 3).color = greenFill
    elif changeInLTP < 0:
        wb_sheet.range(lastRow + 1, 3).color = redFill

    if changeInFutOI > 0:
        wb_sheet.range(lastRow + 1, 6).color = greenFill
    elif changeInFutOI < 0:
        wb_sheet.range(lastRow + 1, 6).color = redFill

    if changeInCallOI > 0:
        wb_sheet.range(lastRow + 1, 8).color = greenFill
    elif changeInCallOI < 0:
        wb_sheet.range(lastRow + 1, 8).color = redFill

    if changeInPutOI > 0:
        wb_sheet.range(lastRow + 1, 10).color = greenFill
    elif changeInPutOI < 0:
        wb_sheet.range(lastRow + 1, 10).color = redFill

    if signal == "Buy":
        wb_sheet.range(lastRow + 1, 12).color = greenFill
    elif signal == "Sell":
        wb_sheet.range(lastRow + 1, 12).color = redFill

    wb_sheet.cells(lastRow + 1, 1).value = time_range
    wb_sheet.cells(lastRow + 1, 2).value = currTraded
    wb_sheet.cells(lastRow + 1, 3).value = changeInLTP
    wb_sheet.cells(lastRow + 1, 4).value = tradedVolCon
    wb_sheet.cells(lastRow + 1, 5).value = futOI
    wb_sheet.cells(lastRow + 1, 6).value = changeInFutOI
    wb_sheet.cells(lastRow + 1, 7).value = callOI
    wb_sheet.cells(lastRow + 1, 8).value = changeInCallOI
    wb_sheet.cells(lastRow + 1, 9).value = putOI
    wb_sheet.cells(lastRow + 1, 10).value = changeInPutOI
    wb_sheet.cells(lastRow + 1, 11).value = OiInter
    wb_sheet.cells(lastRow + 1, 12).value = signal
    wb.save()

def enterInExcel(currTime,isFifteenMin):
    try:
        url1 = "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/ajaxFOGetQuoteJSON.jsp?underlying=NIFTY&instrument=FUTIDX&expiry=24SEP2020&type=SELECT&strike=11500.00"
        referer = "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuoteFO.jsp?underlying=NIFTY&instrument=OPTIDX&expiry=17SEP2020&type=CE&strike=11500.00"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9', 'Accept-Encoding': 'gzip, deflate, br', 'referer': referer}
        cookies = {
            'bm_sv': '808AA9DAA6889DCF6CB5A9AB14BF290F~t420SFeAK8epqchwzNEdgzJGSy4pjs8Mco4O/d7w5qW4TxShiDJYDHoeSAe60uaZMH2n5wjH3v0NnmtPWrKZosEZDfOSHts34ur4GGAPsGGL3LHYVPU/IbcIDhNf6MXF8oqWfpyJ0FXh74VTcL2NhnLG6DlsLun9OrT/+t9vHpY='}
        session = requests.session()
        for cook in cookies:
            session.cookies.set(cook, cookies[cook])

        fut = session.get(url1, headers=headers).json()
        url2 = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
        session = requests.session()

        for cook in cookies:
            session.cookies.set(cook, cookies[cook])
        opt = session.get(url2, headers=headers).json()
        # Future OI
        lastTraded = float((fut["data"][0]["lastPrice"]).replace(',', ''))
        futOI = int((fut["data"][0]["openInterest"]).replace(',', ''))
        tradedVolCon = int((fut["data"][0]["numberOfContractsTraded"]).replace(',', ''))

        # Option OI
        callOI = opt["filtered"]["CE"]["totOI"]
        putOI = opt["filtered"]["PE"]["totOI"]
        putInExcel5Min(lastTraded, futOI, tradedVolCon, callOI, putOI,currTime)
        putOptionChainData(opt,currTime)
        if isFifteenMin:
            putInExcel15Min(lastTraded, futOI, tradedVolCon, callOI, putOI,currTime)

    except Exception as e:
        print(e)

def putInExcelIni(lastTraded, futOI, tradedVolCon, callOI, putOI,currTime):
    wb = xw.Book('OiAnalysis.xlsx')
    sht5min = wb.sheets[0]
    sht15min = wb.sheets[1]
    sht5min.cells(1, 1).value = "Time"
    sht5min.cells(1, 2).value = "LTP"
    sht5min.cells(1, 3).value = "Change in LTP"
    sht5min.cells(1, 4).value = "Traded Volume(Contract)"
    sht5min.cells(1, 5).value = "Future OI"
    sht5min.cells(1, 6).value = "Change Future OI"
    sht5min.cells(1, 7).value = "Call OI"
    sht5min.cells(1, 8).value = "Change Call OI"
    sht5min.cells(1, 9).value = "Put OI"
    sht5min.cells(1, 10).value = "Change Put OI"
    sht5min.cells(1, 11).value = "OI Interpretation"
    sht5min.cells(1, 12).value = "Buy/Sell"

    sht5min.range('A1:L1').color = (255,255,0)

    sht5min.cells(2, 1).value = "00:00-"+currTime
    sht5min.cells(2, 2).value = lastTraded
    sht5min.cells(2, 3).value = 0
    sht5min.cells(2, 4).value = tradedVolCon
    sht5min.cells(2, 5).value = futOI
    sht5min.cells(2, 6).value = 0
    sht5min.cells(2, 7).value = callOI
    sht5min.cells(2, 8).value = 0
    sht5min.cells(2, 9).value = putOI
    sht5min.cells(2, 10).value = 0

    sht15min.cells(1, 1).value = "Time"
    sht15min.cells(1, 2).value = "LTP"
    sht15min.cells(1, 3).value = "Change in LTP"
    sht15min.cells(1, 4).value = "Traded Volume(Contract)"
    sht15min.cells(1, 5).value = "Future OI"
    sht15min.cells(1, 6).value = "Change Future OI"
    sht15min.cells(1, 7).value = "Call OI"
    sht15min.cells(1, 8).value = "Change Call OI"
    sht15min.cells(1, 9).value = "Put OI"
    sht15min.cells(1, 10).value = "Change Put OI"
    sht15min.cells(1, 11).value = "OI Interpretation"
    sht15min.cells(1, 12).value = "Buy/Sell"

    sht15min.range('A1:L1').color = (255, 255, 0)

    sht15min.cells(2, 1).value = "00:00-"+currTime
    sht15min.cells(2, 2).value = lastTraded
    sht15min.cells(2, 3).value = 0
    sht15min.cells(2, 4).value = tradedVolCon
    sht15min.cells(2, 5).value = futOI
    sht15min.cells(2, 6).value = 0
    sht15min.cells(2, 7).value = callOI
    sht15min.cells(2, 8).value = 0
    sht15min.cells(2, 9).value = putOI
    sht15min.cells(2, 10).value = 0
    wb.save()

def initializeFiles(currTime):
    try:
        url1 = "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/ajaxFOGetQuoteJSON.jsp?underlying=NIFTY&instrument=FUTIDX&expiry=24SEP2020&type=SELECT&strike=11500.00"
        referer = "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuoteFO.jsp?underlying=NIFTY&instrument=OPTIDX&expiry=17SEP2020&type=CE&strike=11500.00"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9', 'Accept-Encoding': 'gzip, deflate, br', 'referer': referer}
        cookies = {
            'bm_sv': '808AA9DAA6889DCF6CB5A9AB14BF290F~t420SFeAK8epqchwzNEdgzJGSy4pjs8Mco4O/d7w5qW4TxShiDJYDHoeSAe60uaZMH2n5wjH3v0NnmtPWrKZosEZDfOSHts34ur4GGAPsGGL3LHYVPU/IbcIDhNf6MXF8oqWfpyJ0FXh74VTcL2NhnLG6DlsLun9OrT/+t9vHpY='}
        session = requests.session()
        for cook in cookies:
            session.cookies.set(cook, cookies[cook])

        fut = session.get(url1, headers=headers).json()
        url2 = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
        session = requests.session()

        for cook in cookies:
            session.cookies.set(cook, cookies[cook])
        opt = session.get(url2, headers=headers).json()
        # Future OI
        lastTraded = float((fut["data"][0]["lastPrice"]).replace(',', ''))
        futOI = int((fut["data"][0]["openInterest"]).replace(',', ''))
        tradedVolCon = int((fut["data"][0]["numberOfContractsTraded"]).replace(',', ''))

        # Option OI
        callOI = opt["filtered"]["CE"]["totOI"]
        putOI = opt["filtered"]["PE"]["totOI"]

        putInExcelIni(lastTraded, futOI, tradedVolCon, callOI, putOI,currTime)
        makeOptionChainFile(opt,currTime)
        return True
    except Exception as e:
        print(e)
        return False


if __name__ == '__main__':
   success = False

   wb = xw.Book()
   wb.save('OiAnalysis.xlsx')
   wb.sheets.add(name='FiftMin')
   wb.sheets.add(name='FiveMin')
   for sheet in wb.sheets:
       if 'Sheet' in sheet.name:
           sheet.delete()
   # This loop runs from 9:15 AM to 3:30 PM till Market hours
   while(1):
        t1 = datetime.now()
        currTime = str(t1.hour)+":"+str(t1.minute)
        if(time(9,15)<=datetime.now().time()<=time(9,17,30) and not(success)):
            success = initializeFiles(currTime)

        isFifMin=1
        while(time(9,20)<=datetime.now().time()<=time(15,31)):
            t1 = datetime.now()
            currTime = str(t1.hour) + ":" + str(t1.minute)
            isFifMin+=1
            isFifteenMin = False
            if isFifMin==3:
                isFifteenMin = True
                isFifMin=0
            enterInExcel(currTime,isFifteenMin)
            t1 = datetime.now()
            while(t1.minute%5 !=0 or t1.second !=0):
                t1 = datetime.now()
                sleep(1)
        sleep(1)