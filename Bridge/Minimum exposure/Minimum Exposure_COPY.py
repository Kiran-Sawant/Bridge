"""This script writes the ask price of Forex, commodities, index, stocks, etf, bond_futures on the given Excel file on line 15"""

import MetaTrader5 as mt5
import xlwings as xl
import time

#__________Terminal Initialization_______#
mt5.initialize()
accountInfo = mt5.account_info()
print("Name: {0}\nServer: {1}\nBalance: ${2}\nEquity: ${3}\nFree Margin: ${4}\nMargin Level: {5}%"
.format(accountInfo.name, accountInfo.server, accountInfo.balance, accountInfo.equity,
accountInfo.margin_free, accountInfo.margin_level))

#_____WorkBook Initialization______#
wBook = xl.Book('Minimum Exposure_COPY.xlsx')
#_______worksheet assignment_______#
wsLive = wBook.sheets('live')
wsStocks = wBook.sheets('Stocks')

#______________________dictionaries__________________________#
#__________Fixed Assets______________#
'''Assets who's name does not change.'''

forex = ['EURUSD', 'GBPUSD', 'AUDUSD', 'NZDUSD', 'USDCHF', 'USDJPY', 'USDCAD', 'USDSGD',
        'USDCNH', 'USDCZK', 'USDDKK', 'USDHKD', 'USDHUF', 'USDMXN', 'USDNOK', 'USDPLN',
        'USDRUB', 'USDSEK', 'USDTHB', 'USDTRY', 'USDZAR']

commodities = ['XAUUSD', 'XAUEUR', 'XAUAUD', 'XAGEUR', 'XAGUSD', 'XPDUSD', 'XPTUSD',
               'XBRUSD', 'XNGUSD', 'XTIUSD']

index = ['US30', 'US500', 'USTEC', 'DE30', 'STOXX50', 'UK100', 'JP225', 'AUS200', 'F40',
         'HK50', 'IT40', 'CHINA50', 'ES35', 'US2000', 'CA60', 'NETH25', 'SE30', 'MidDE60',
         'SWI20', 'CHINAH', 'SA40', 'NOR25', 'TecDE30']

stonks_d = {'B8':'ABT.NYSE', 'B9':'ACB.NYSE', 'B10':'AGN.NYSE', 'B11':'AXP.NYSE', 'B12':'BA.NYSE', 'B13':'BAC.NYSE',
          'B14':'BLK.NYSE', 'B15':'BMY.NYSE', 'B16':'C.NYSE', 'B17':'CGC.NYSE', 'B18':'CVS.NYSE', 'B19':'CVX.NYSE',
          'B20':'DD.NYSE', 'B21':'DIS.NYSE', 'B22':'DWDP.NYSE', 'B23':'GE.NYSE', 'B24':'GS.NYSE', 'B25':'HD.NYSE',
          'B26':'HON.NYSE', 'B27':'IBM.NYSE', 'B28':'JNJ.NYSE', 'B29':'JPM.NYSE', 'B30':'KO.NYSE', 'B31':'LLY.NYSE',
          'B32':'LMT.NYSE', 'B33':'MA.NYSE', 'B34':'MCD.NYSE', 'B35':'MMM.NYSE', 'B36':'MO.NYSE', 'B37':'ABBV.NYSE',
          'B38':'MRK.NYSE', 'B39':'PINS.NYSE', 'B40':'MS.NYSE', 'B41':'NKE.NYSE', 'B42':'ORCL.NYSE', 'B43':'PFE.NYSE',
          'B44':'PG.NYSE', 'B45':'PM.NYSE', 'B46':'SLB.NYSE', 'B47':'SPOT.NYSE', 'B48':'T.NYSE', 'B49':'TWX.NYSE',
          'B50':'UNH.NYSE', 'B51':'UNP.NYSE', 'B52':'UPS.NYSE', 'B53':'USB.NYSE', 'B54':'UTX.NYSE', 'B55':'V.NYSE',
          'B56':'VZ.NYSE', 'B57':'WFC.NYSE', 'B58':'WMT.NYSE', 'B59':'XOM.NYSE', 'B60':'LIN.NYSE', 'B61':'DEO.NYSE',
          'B62':'NEE.NYSE', 'B63':'AMT.NYSE', 'B64':'BABA.NYSE', 'B65':'BRK-B.NYSE', 'B66':'TMO.NYSE', 'B67':'TM.NYSE',
          'B68':'UBER.NYSE', 'B70':'ADBE.NAS', 'B71':'ADI.NAS', 'B72':'ADP.NAS', 'B73':'AMAT.NAS', 'B74':'AMGN.NAS', 'B75':'AMZN.NAS',
          'B76':'ATVI.NAS', 'B77':'AVGO.NAS', 'B78':'BIIB.NAS', 'B79':'BKNG.NAS', 'B80':'CELG.NAS', 'B81':'CHTR.NAS', 'B82':'CMCSA.NAS',
          'B83':'CME.NAS', 'B84':'COST.NAS', 'B85':'CRON.NAS', 'B86':'CSCO.NAS', 'B87':'CSX.NAS', 'B88':'CTSH.NAS', 'B89':'DISH.NAS',
          'B90':'EA.NAS', 'B91':'EBAY.NAS', 'B92':'EQIX.NAS', 'B93':'FB.NAS', 'B94':'FOX.NAS', 'B95':'FOXA.NAS', 'B96':'AAPL.NAS',
          'B97':'GILD.NAS', 'B98':'LYFT.NAS', 'B99':'GOOG.NAS', 'B100':'INTC.NAS', 'B101':'INTU.NAS', 'B102':'ISRG.NAS', 'B103':'KHC.NAS',
          'B104':'MAR.NAS', 'B105':'MDLZ.NAS', 'B106':'MSFT.NAS', 'B107':'MU.NAS', 'B108':'NFLX.NAS',
          'B109':'NVDA.NAS', 'B110':'PEP.NAS', 'B111':'PYPL.NAS', 'B112':'QCOM.NAS', 'B113':'REGN.NAS', 'B114':'SBUX.NAS', 'B115':'TLRY.NAS',
          'B116':'TMUS.NAS', 'B117':'TSLA.NAS', 'B118':'TXN.NAS', 'B119':'VRTX.NAS', 'B120':'WBA.NAS'}

etf = [ 'DBA.NYSE', 'EEM.NYSE', 'EWH.NYSE', 'EWW.NYSE', 'EWZ.NYSE', 'FXI.NYSE', 'GDX.NYSE',
        'GDXJ.NYSE', 'IEMG.NYSE', 'LQD.NYSE', 'MJ.NYSE', 'MOO.NYSE', 'RSX.NYSE', 'SIL.NYSE',
        'UNG.NYSE', 'URA.NYSE', 'USO.NYSE', 'VXX.NYSE', 'VXXB.NYSE', 'VYM.NYSE', 'XLE.NYSE',
        'XLF.NYSE', 'XLI.NYSE', 'XLP.NYSE', 'XLU.NYSE', 'XOP.NYSE', 'QQQ.NAS', 'TLT.NAS']

#____________Variable Assets______________#
"""Assets that can change name periodically.
    Futures contracts change their names monthly"""

Future = '*_*, !Sugar*, !Corn*, !Coffee*, !Cocoa*, !Cotton*, !Wheat*, !OJ*, !Soyb*'

variable_assets = list()
all_ass_data = mt5.symbols_get(Future)

for info in all_ass_data:
    variable_assets.append(info.name)

variable_assets.sort()

def insertion(symbol):
    mt5.symbol_select(symbol)


def ask(assetName):
    info = mt5.symbol_info_tick(assetName)
    if info is None:
        insertion(assetName)
        time.sleep(0.5)
        info = mt5.symbol_info_tick(assetName)
        if info is None:
            string = "NA"
            return string
    return info.ask


#________________Writing to excel file_____________________#
wsLive.range('B1').value = accountInfo.name
wsLive.range('B2').value = accountInfo.server

while True:
    #_______________Account info__________________#
    wsLive.range('B3').value = mt5.account_info().equity
    wsLive.range('B4').value = mt5.account_info().margin_free
    wsLive.range('B5').value = mt5.account_info().margin_level
    wsLive.range('B6').value = mt5.account_info().balance

    #____________fixed name assets____________#
    for i in range(len(forex)):              #Forex Data
        wsLive.range(f"D{i+5}").value = forex[i]
        wsLive.range(f"E{i+5}").value = ask(forex[i])

    for i in range(len(commodities)):        #Commodities Data
        wsLive.range(f"G{i+5}").value = commodities[i]
        wsLive.range(f"H{i+5}").value = ask(commodities[i])

    for i in range(len(index)):             #Index Data
        wsLive.range(f"J{i+5}").value = index[i]
        wsLive.range(f"K{i+5}").value = ask(index[i])

    for i in range(len(etf)):               #ETF Data
        wsLive.range(f"M{i+5}").value = etf[i]
        wsLive.range(f"N{i+5}").value = ask(etf[i])

    #__________Variable name assets___________#
    for j in range(len(variable_assets)):
        wsLive.range(f"P{j+5}").value = variable_assets[j]
        wsLive.range(f"Q{j+5}").value = ask(variable_assets[j])

    #________stocks sheet________#
    for i in stonks_d:                      #Stocks Data
        wsStocks.range(i).value = ask(stonks_d[i])

    #___resting time____#
    time.sleep(20)
    print("now!!!")
    time.sleep(2)
