import MetaTrader5 as mt5
import xlwings as xl
import time

#______Terminal Initialization_____#
mt5.initialize()
accountInfo = mt5.account_info()
print("Name: {0}\nServer: {1}\nBalance: ${2}\nEquity: ${3}\nFree Margin: ${4}\nMargin Level: {5}%"
.format(accountInfo.name, accountInfo.server, accountInfo.balance, accountInfo.equity,
accountInfo.margin_free, accountInfo.margin_level))

#______WorkBook Initialization______#
wBook = xl.Book('Market Exposure_COPY.xlsx')
wsLive = wBook.sheets('Live')       #worksheet assignment

#___________dictionaries___________#
forex = {'B6':'EURUSD', 'B7':'GBPUSD', 'B8':'AUDUSD', 'B9':'NZDUSD',
         'B11':'USDJPY', 'B12':'USDCHF', 'B13':'USDCAD', 'B14':'USDSGD',
         'B15':'USDNOK', 'B16':'USDSEK', 'B17':'USDTRY', 'B18':'USDZAR',
         'B19':'USDCNH', 'B20':'USDCZK', 'B21':'USDDKK', 'B22':'USDHKD',
         'B23':'USDHUF', 'B24':'USDMXN', 'B25':'USDPLN', 'B26':'USDRUB', 'B27':'USDTHB'}

def insertion(symbol):
    mt5.symbol_select(symbol)


#_________retrival method___________#
def closeingPrice(name):
    info = mt5.symbol_info_tick(name)
    if info is None:
        insertion(name)
        time.sleep(0.5)
        info = mt5.symbol_info_tick(name)
        return info.ask
    return info.ask


#_________account info__________#
def equity():
    info = mt5.account_info()
    return info.equity

#_____________Writing to excel file____________#
while True:
    wsLive.range('B2').value = equity()

    #________forex sheet_________#
    for i in forex:
        wsLive.range(i).value = closeingPrice(forex[i])
    
    #___resting time____#
    time.sleep(28)
    print("now!!!")
    time.sleep(2)
