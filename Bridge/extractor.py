"""Gives a list of tradable assets"""

import MetaTrader5 as mt5

mt5.initialize()

print(f"module version: {mt5.__version__}")
print(f"Terminal Version: {mt5.version()}")


#____________________________________________________________#
# pass the strings below in .symbols_get() to extract symbols
# of specific asset classes into asset_list

usd_quote = "*USD"
usd_base = """USD*,!X*, !BTC*, !BCH*,!DSH*,
                !DSH*, !ETH*, !LTC*, !EOS*, !EMC*, !NMC*, !PPC*"""
commodities = "X*, !*.NYSE"
stock_nasdaq = "*.NAS"
stock_nyse = "*.NYSE"
Future = '*_*'
bonds = 'EURB*, EURSCHA*, ITBT*, JGB*, UKGB*, UST*, !USTEC'

all_ass_data = mt5.symbols_get()        # will return symbols of all tradable assets.
asset_list = list()

for info in all_ass_data:
    if info.custom == True:
        continue
    else:
        # print(info.name)
        asset_list.append(info.name)

print(asset_list)
# for i in asset_list:
#     print(i)
# mt5.shutdown()