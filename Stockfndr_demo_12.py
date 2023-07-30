import matplotlib as mpl
mpl.use("Agg")
import matplotlib.pyplot as plt
plt.style.use('bmh')
import talib as ta
from talib import WMA
import numpy as np
import pandas as pd
import datetime as dt
from datetime import datetime
from tkinter import *
from PIL import Image,ImageTk
import xlsxwriter
from tvDatafeed import TvDatafeed, Interval
import sys
import os

# global variables
global stock_nr
global data
global naam_tab
global lijst_foute_tickers

# regel nummers waar aandeelgegevens worden uitgevoerd in uitvoer excel
rownr_buy = rownr_sell = rownr_geen = 1
# tijdelijke waardes voor tick en portefeuille
tick = "test:test"
portefeuille = "ppppp"
# tspercent is de ruimte voor stoploss wordt getriggerd
tspercent = 0.2
# aantal weken is 62 vanwege het wma62 plus een jaar
number_of_bars = 52 + 62
# aantal aandelen in excel. Default 5. wordt aangepast na download
number_of_stocks = 5
# lijst foute tickers is leeg dataframe
lijst_foute_tickers = pd.DataFrame(columns = ['foute_ticker']) 
# definitie van values voor configure
# portef_value = StringVar()
# tick_value = StringVar()
today = dt.date.today()
signaal_monday = today + dt.timedelta(days=-today.weekday(), weeks=round((today.weekday()-8)/7))
# this_monday = today + dt.timedelta(days=-today.weekday(), weeks=0)
signaal = "Neutraal"
trend = "Neutraal"
achtergrkleur = "dimgrey"
voorgrkleur = "white"
contrastkleur = "blue"
achtergrrood = "lightcoral"
achtergrgroen = "palegreen"

# maak test png
test = plt.figure(figsize = (10,10))
plt.ylim(0,10)
plt.xlim(0,10)
plt.text(5, 5, "Geen", color = "blue", fontsize = 40, ha = 'center', weight = 'semibold')
plt.text(5, 4, "Grafiek", color = "blue", fontsize = 40, ha = 'center', weight = 'semibold')
plt.title("Geen grafiek")
test.savefig("test_test.png", bbox_inches='tight', dpi=150, transparent=False)
plt.close(test)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev abd for pyinstaller """
    try:
        # pyinstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")  
    return os.path.join(base_path, relative_path)

def normal_round(num, ndigits=0):
    """
    Rounds a float to the specified number of decimal places.
    num: the value to round
    ndigits: the number of digits to round to
    """
    if ndigits == 0:
        return int(num + 0.5)
    else:
        digit_value = 10 ** ndigits
        if num > 0:
            return int(num * digit_value + 0.5) / digit_value
        elif num < 0:
            return int(num * digit_value - 0.5) / digit_value
        else:
            return int(0)
        
def verwijder_png_files():
    teller = 0
    tellermax = len(stock_list)
    os.remove("test_test.png")
    while teller < tellermax:
        tick = stock_list.loc[stock_list.index[teller],"ticker"]
        n = tick.split(":")
        naam_plot = str(n[0])+"_"+str(n[1])+".png"
        try:
            os.remove(naam_plot)
        except:
            pass
        teller += 1

def download_fout(msg):
    global lijst_foute_tickers
    quitwindow = Tk()
    quitwindow.title("Foutmelding")
    quitwindow["bg"] = "lightblue"
    achtergrkleur = "blue"
    voorgrkleur = "white"
    quitwindow.geometry("500x200+200+200")
    fout_label = Label(quitwindow, text = msg, font = ("Arial",30), bg = achtergrkleur, fg = voorgrkleur, padx = 10, pady = 10)
    fout_button = Button(quitwindow, text = "Exit", activebackground = "blue", activeforeground = "white", bg = achtergrkleur, fg = achtergrkleur, font = ("Arial",30), command = quitwindow.destroy)
    fout_label.pack(padx = 10, pady = 10),
    fout_button.pack(padx = 10, pady = 10)
    quitwindow.mainloop()
    
     
def message(teller, totaal, msg):
    global message_label
    global row_nr
    message_label.destroy
    if teller < 0:
        message_label = Label(window, text = msg, bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        message_label.grid(column=0, row=row_nr, columnspan = 3, sticky=(W), padx=10, pady=10)
    else:
        teller += 1
        totaal += 1
        msg = str(teller) + " / " + str(totaal)    
        message_label = Label(window, text = msg, bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        message_label.grid(column=3, row=row_nr, sticky=(W), padx=10, pady=10)
    message_label.after(3000, message_label.destroy)
    return()
 

# Voeg kolommen toe aan de stock data voor berekeningen/output van TN
def voeg_kolommen_toe(__data,number_of_bars,tspercent):
    __data["wma62"] = WMA(__data.close, timeperiod=62)
    __data["wma4"] = WMA(__data.close, timeperiod=4)
    #1
    __data["vorigewma62"] = __data["wma62"].shift(1)
    __data["wma62OK"] = (__data.wma62>__data.vorigewma62)
    __data["NOTwma62OK"] = (__data.wma62<__data.vorigewma62)
    #2
    __data["vorigeclose"]=__data["close"].shift(1)
    __data["closeOK"] = (__data.close>__data.vorigeclose)
    __data["NOTcloseOK"] = (__data.close<__data.vorigeclose)
    #3
    __data["vorigewma4"] = __data["wma4"].shift(1)
    __data["vorigewma4OK"] = (__data.vorigeclose>__data.vorigewma4)
    __data["NOTvorigewma4OK"] = (__data.vorigeclose<__data.vorigewma4)
    #4
    __data["wma4OK"] = (__data.close>__data.wma4)
    __data["NOTwma4OK"] = (__data.close<__data.wma4)
    # Maaak kolommen voor enterlong en entershort:
    __data["enterlong"] = __data.wma62OK & __data.wma4OK & __data.vorigewma4OK & __data.closeOK
    __data["entershort"] = (__data.NOTwma62OK) & (__data.NOTwma4OK) & (__data.NOTvorigewma4OK) & (__data.NOTcloseOK)
    # Maak Trailtop kolom
    __data["trailtop"] = __data["close"]
    __data.loc[__data.index[0], "trailtop"] = 0
    __data.loc[__data.index[1], "trailtop"] = 0
    n = 1
    while n <= number_of_bars - 1:
        if normal_round(__data.loc[__data.index[n], "close"],2) > normal_round(__data.loc[__data.index[n-1], "trailtop"],2) or normal_round(__data.loc[__data.index[n], "close"],2) < normal_round(__data.loc[__data.index[n-1], "trailtop"]*(1-tspercent),2) or __data.loc[__data.index[n], "entershort"]:
           __data.loc[__data.index[n], "trailtop"] = __data.loc[__data.index[n], "close"]
        else:
           __data.loc[__data.index[n], "trailtop"] = __data.loc[__data.index[n-1], "trailtop"]
        n += 1
    __data.loc[__data.index[0],"trailtop"] = __data.loc[__data.index[0],"close"]
    # Maak trailbot kolom
    __data["trailbot"] = __data["close"]
    __data.loc[__data.index[0], "trailbot"] = 0
    __data.loc[__data.index[1], "trailbot"] = 0
    n = 1
    while n <= number_of_bars - 1:
        if normal_round(__data.loc[__data.index[n], "close"],2) < normal_round(__data.loc[__data.index[n-1], "trailbot"],2) or normal_round(__data.loc[__data.index[n], "close"],2) > normal_round(__data.loc[__data.index[n-1], "trailbot"]*(1+tspercent),2) or __data.loc[__data.index[n], "enterlong"]:
            __data.loc[__data.index[n], "trailbot"] = __data.loc[__data.index[n], "close"]
        else:
            __data.loc[__data.index[n], "trailbot"] = __data.loc[__data.index[n-1], "trailbot"]
        n += 1
    __data.loc[__data.index[0],"trailbot"] = __data.loc[__data.index[0],"close"] 
    # maak kolommen voor exitlong en exitshort
    __data["exitlong"] = __data["enterlong"]
    __data["exitshort"] = __data["entershort"]
    # nu de kolom exitlong vullen met juiste waarden
    n = 1
    while n <= number_of_bars - 1:
        if normal_round(__data.loc[__data.index[n], "trailtop"],2) < normal_round(__data.loc[__data.index[n-1], "trailtop"],2):
            __data.loc[__data.index[n], "exitlong"] = True
        else:
            __data.loc[__data.index[n], "exitlong"] = False
        n += 1
    # Vul kolom van exitshort
    n = 1
    while n <= number_of_bars - 1:
        if normal_round(__data.loc[__data.index[n], "trailbot"],2) > normal_round(__data.loc[__data.index[n-1], "trailbot"],2):
            __data.loc[__data.index[n], "exitshort"] = True
        else:
            __data.loc[__data.index[n], "exitshort"] = False
        n += 1
    # maak kolommen voor inlong en inshort
    __data["inlong"] = __data["enterlong"]
    __data["inshort"] = __data["entershort"]
    # vul de kolom van inlong
    n = 1
    while n <= number_of_bars - 1:
        if __data.loc[__data.index[n], "enterlong"]:
            __data.loc[__data.index[n], "inlong"] = True
        elif __data.loc[__data.index[n], "entershort"] or __data.loc[__data.index[n], "exitlong"]:
            __data.loc[__data.index[n], "inlong"] = False
        else:
            __data.loc[__data.index[n], "inlong"] = __data.loc[__data.index[n-1], "inlong"]
        n += 1
    # vul de kolom van inshort
    n = 1
    while n <= number_of_bars - 1:
        if __data.loc[__data.index[n], "entershort"]:
            __data.loc[__data.index[n], "inshort"] = True
        elif __data.loc[__data.index[n], "enterlong"] or __data.loc[__data.index[n], "exitshort"]:
            __data.loc[__data.index[n], "inshort"] = False
        else:
            __data.loc[__data.index[n], "inshort"] = __data.loc[__data.index[n-1], "inshort"]
        n += 1
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    # __data.to_excel('data.xlsx', sheet_name='sheet1', index = False)
    return(__data)

# lees lijst van beurs,aandelen vanuit het excel bestand.
# selecteer alleen de actieve (actief = 1) aandelen
def get_stock_list():
    # """gets list with active stocks from excel"""
    # """kolommen zijn ticker, portefeuille, bedrijfsnaam, actief, positie"""
    # """Alleen ticker, portefeuille en actief worden gedownload. Daarna alleen actief = 1 eruit gefilterd"""
    # try:
    #     data = pd.read_excel('SFPortefeuille_stocks.xlsx', usecols=["ticker", "portefeuille", "actief"])
    # except:
    #     msg = "Fout: Bestand \nSFPortefeuille_stocks.xlsx \nniet gevonden"
    #     download_fout(msg)
    # data = data[data.actief == 1]
    # data = data.sort_values(["portefeuille"])
    ''' nu voor de demo worden de aandelen vast ingevoerd in de dataframe data'''
    data_CRYPTO = [["CRYPTO","BINANCE:AVAXUSD"],
                ["CRYPTO","BINANCE:BNBUSD"],
                ["CRYPTO","BINANCE:BTCUSD"],
                ["CRYPTO","BINANCE:BCHUSD"],
                ["CRYPTO","BINANCE:ADAUSD"],
                ["CRYPTO","BINANCE:LINKUSD"],
                ["CRYPTO","BINANCE:DASHUSD"],
                ["CRYPTO","BINANCE:DOGEUSD"],
                ["CRYPTO","BINANCE:ETHUSD"],
                ["CRYPTO","BINANCE:ETCUSD"],
                ["CRYPTO","BINANCE:FILUSD"],
                ["CRYPTO","BINANCE:LTCUSD"],
                ["CRYPTO","BINANCE:XMRUSD"],
                ["CRYPTO","BINANCE:DOTUSD"],
                ["CRYPTO","BINANCE:XRPUSD"],
                ["CRYPTO","BINANCE:SOLUSD"],
                ["CRYPTO","BINANCE:XLMUSD"]]
    data_AEX_trend = [["AEXTrend","EURONEXT:AALB"],
                      ["AEXTrend","XETR:ACT"],
                      ["AEXTrend","EURONEXT:AD"],
                      ["AEXTrend","XETR:AEIN"],
                      ["AEXTrend","EURONEXT:AKZA"],
                      ["AEXTrend","EURONEXT:ALFRE"],
                      ["AEXTrend","EURONEXT:ALO"],
                      ["AEXTrend","EURONEXT:ASM"],
                      ["AEXTrend","EURONEXT:ASML"],
                      ["AEXTrend","EURONEXT:ASRNL"],
                      ["AEXTrend","EURONEXT:ATE"],
                      ["AEXTrend","EURONEXT:AVTX"],
                      ["AEXTrend","XETR:CSH"],
                      ["AEXTrend","EURONEXT:DSM"],
                      ["AEXTrend","EURONEXT:DSY"],
                      ["AEXTrend","EURONEXT:ENX"],
                      ["AEXTrend","EURONEXT:FGR"],
                      ["AEXTrend","EURONEXT:GLPG"],
                      ["AEXTrend","EURONEXT:IMCD"],
                      ["AEXTrend","EURONEXT:LIGHT"],
                      ["AEXTrend","EURONEXT:MT"],
                      ["AEXTrend","EURONEXT:PHIA"],
                      ["AEXTrend","EURONEXT:RAND"],
                      ["AEXTrend","EURONEXT:SOP"],
                      ["AEXTrend","XETR:FNTN"],
                      ["AEXTrend","EURONEXT:TWEKA"],
                      ["AEXTrend","EURONEXT:WKL"],
                      ["AEXTrend","XETR:CEC"],
                      ["AEXTrend","EURONEXT:SBMO"],
                      ["AEXTrend","XETR:CLIQ"]]
    data_Nasdaq = [["NASDAQ","NASDAQ:AAPL"],
                   ["NASDAQ","NASDAQ:ADBE"],
                   ["NASDAQ","NASDAQ:ADP"],
                   ["NASDAQ","NASDAQ:AMAT"],
                   ["NASDAQ","NASDAQ:AMD"],
                   ["NASDAQ","NASDAQ:AMGN"],
                   ["NASDAQ","NASDAQ:AVGO"],
                   ["NASDAQ","NASDAQ:BIDU"],
                   ["NASDAQ","NASDAQ:BIIB"],
                   ["NASDAQ","NASDAQ:CHTR"],
                   ["NASDAQ","NASDAQ:CMCSA"],
                   ["NASDAQ","NASDAQ:COST"],
                   ["NASDAQ","NASDAQ:CSCO"],
                   ["NASDAQ","NASDAQ:CSX"],
                   ["NASDAQ","NASDAQ:META"],
                   ["NASDAQ","NASDAQ:FISV"],
                   ["NASDAQ","NASDAQ:GILD"],
                   ["NASDAQ","NASDAQ:INTC"],
                   ["NASDAQ","NASDAQ:MDLZ"],
                   ["NASDAQ","NASDAQ:MSFT"],
                   ["NASDAQ","NASDAQ:NDAQ"],
                   ["NASDAQ","NASDAQ:NFLX"],
                   ["NASDAQ","NASDAQ:NVDA"],
                   ["NASDAQ","NASDAQ:PEP"],
                   ["NASDAQ","NASDAQ:PYPL"],
                   ["NASDAQ","NASDAQ:QCOM"],
                   ["NASDAQ","NASDAQ:SBUX"],
                   ["NASDAQ","NASDAQ:TMUS"],
                   ["NASDAQ","NASDAQ:TXN"],
                   ["NASDAQ","NASDAQ:WBA"]]
    data_Smallcaps = [["SmallCaps","EURONEXT:ALFEN"],
                      ["SmallCaps","EURONEXT:AMG"],
                      ["SmallCaps","EURONEXT:ACOMO"],
                      ["SmallCaps","EURONEXT:ARCAD"],
                      ["SmallCaps","EURONEXT:ASCX"],
                      ["SmallCaps","EURONEXT:AVTX"],
                      ["SmallCaps","EURONEXT:BFIT"],
                      ["SmallCaps","EURONEXT:BESI"],
                      ["SmallCaps","EURONEXT:BBED"],
                      ["SmallCaps","EURONEXT:BRNL"],
                      ["SmallCaps","EURONEXT:CMCOM"],
                      ["SmallCaps","EURONEXT:CTAC"],
                      ["SmallCaps","EURONEXT:FAGR"],
                      ["SMALLCAPS","XETR:FTK"],
                      ["SmallCaps","EURONEXT:FFARM"],
                      ["SmallCaps","EURONEXT:HEIJM"],
                      ["SmallCaps","EURONEXT:KENDR"],
                      ["SmallCaps","EURONEXT:BAMNB"],
                      ["SMALLCAPS","XETR:KRN"],
                      ["SmallCaps","EURONEXT:BOLS"],
                      ["SmallCaps","EURONEXT:MTU"],
                      ["SmallCaps","EURONEXT:NEDAP"],
                      ["SmallCaps","EURONEXT:NSI"],
                      ["SmallCaps","EURONEXT:ORDI"],
                      ["SmallCaps","EURONEXT:PHARM"],
                      ["SmallCaps","EURONEXT:SIFG"],
                      ["SmallCaps","EURONEXT:SLIGR"],
                      ["SmallCaps","EURONEXT:S30"],
                      ["SmallCaps","EURONEXT:VALUE"],
                      ["SmallCaps","EURONEXT:WDP"]]
    data_PEG = [["PEG","NYSE:ASX"],
                ["PEG","NYSE:BIO"],
                ["PEG","NYSE:CIVI"],
                ["PEG","TSX:CMMC"],
                ["PEG","NASDAQ:COOP"],
                ["PEG","NASDAQ:CRSR"],
                ["PEG","NYSE:CS"],
                ["PEG","NASDAQ:CSIQ"],
                ["PEG","NYSE:CWH"],
                ["PEG","NASDAQ:ECPG"],
                ["PEG","TSX:GGD"],
                ["PEG","NYSE:GTN"],
                ["PEG","NYSE:JEF"],
                ["PEG","NYSE:KBH"],
                ["PEG","NASDAQ:LKQ"],
                ["PEG","NYSE:LPG"],
                ["PEG","NYSE:MDC"],
                ["PEG","NASDAQ:NAVI"],
                ["PEG","NASDAQ:NBIX"],
                ["PEG","NASDAQ:OESX"],
                ["PEG","NYSE:PBA"],
                ["PEG","NYSE:PFSI"],
                ["PEG","NYSE:PHM"],
                ["PEG","NASDAQ:PNFP"],
                ["PEG","NYSE:RM"],
                ["PEG","NASDAQ:SLM"],
                ["PEG","NYSE:TROX"],
                ["PEG","NASDAQ:VCTR"],
                ["PEG","NASDAQ:QFIN"]]
    data_Indices = [["Indices","Euronext:AEX"],
                    ["Indices","EURONEXT:AMX"],
                    ["Indices","EURONEXT:ASCX"],
                    ["Indices","EURONEXT:BEL20"],
                    ["Indices","TVC:CAC40"],
                    ["Indices","XETR:DAX"],
                    ["Indices","VELOCITY:STOXX50"],
                    ["Indices","VANTAGE:FTSE100"],
                    ["Indices","TVC:HSI"],
                    ["Indices","BMFBOVESPA:IBOV"],
                    ["Indices","MIL:IMIB"],
                    ["Indices","LSIN:0MNK"],
                    ["Indices","SWB:LGQM"],
                    ["Indices","EURONEXT:FIN"],
                    ["Indices","SWB:LIRU"],
                    ["Indices","XETR:MDAX"],
                    ["Indices","NASDAQ:NQPHPHP"],
                    ["Indices","EURONEXT:PSI20"],
                    ["Indices","SP:SPX"],
                    ["Indices","SIX:SMI"],
                    ["Indices","XETR:EXXT"],
                    ["Indices","BME:LYXIB"]]
    data_Indices = pd.DataFrame(data = data_Indices, columns = ["portefeuille","ticker"])
    data_AEX_trend = pd.DataFrame(data = data_AEX_trend, columns = ["portefeuille","ticker"])
    data_Smallcaps = pd.DataFrame(data = data_Smallcaps, columns = ["portefeuille","ticker"])
    data_Nasdaq = pd.DataFrame(data = data_Nasdaq, columns = ["portefeuille","ticker"])
    data_PEG = pd.DataFrame(data = data_PEG, columns = ["portefeuille","ticker"])
    # data_CRYPTO = pd.DataFrame(data = data_CRYPTO, columns = ["portefeuille","ticker"])
    data = pd.concat([data_Indices, data_Smallcaps, data_Nasdaq, data_AEX_trend, data_PEG], axis = 0, ignore_index=True)
    # data = data_AEX_trend
    return(data)


# download aandeel gegevens vanuit Trading View
# return met df data
def get_stock_data(stock,exchange,number_of_bars):
    """download open, high, low, close, volume van een bepaald aandeel van de laatste .. weken"""
    data = tv.get_hist(stock, exchange, Interval.in_weekly, n_bars = number_of_bars)
    return (data)

def download_process_data(stock_nr):
    global tick
    global data
    global msg
    global window
    global stock_list
    tick = stock_list.loc[stock_list.index[stock_nr],"ticker"]
    portefeuille = stock_list.loc[stock_list.index[stock_nr],"portefeuille"]
    # print(portefeuille)
    # verwijder eventuele komma aan het eind
    tick = tick.strip(",")
    # print(tick)
    # splits in aparte beurs en tickercode
    n = tick.split(":")
    # print(n[0],n[1])
    # download data van trading view
    data = get_stock_data(n[1],n[0],number_of_bars)
    data = pd.DataFrame(data)
    dat_empty = data.empty
    if dat_empty:
        # msg = "Ticker " + str(tick) + " Niet gevonden. \nControleer de beurs van het aandeel \nen/of de tickercode is onjuist"
        # popup_message(msg)
        msg = "Ticker " + str(tick) + " Niet gevonden. Controleer de beurs en/of de tickercode"
        message(-1, -1, msg)
        trend = "Foutmelding"
        signaal = "Tickercode onjuist"
    elif len(data.index)<=113:
        # msg = "Ticker " + str(tick) + " niet mogelijk \nNiet genoeg weken voor dit aandeel beschikbaar \nDeactiveer ticker in het bestand \nSFPortefeuille_stocks.xlsx"
        # popup_message(msg)
        # return("leeg port", "leeg ticker", "Leeg trend", "leeg sign")
        msg = "Ticker " + str(tick) + " heeft niet genoeg weken (min 114) beschikbaar"
        message(-1, -1, msg)
        trend = "Foutmelding"
        signaal = "aantal wkn < 114"
    else:
        # breidt de dataframe uit met kolommen voor tradingnavigator
        data = voeg_kolommen_toe(data,number_of_bars,tspercent)
        # print("kolommen toegevoegd")
        # create de plot van de graph
        trend, signaal = maak_graphplot(data, tick, number_of_bars, tspercent)
    # print(trend, signaal)
    return(portefeuille, tick, trend, signaal)

def configureer_scherm(portefeuille, tick, trend, signaal):
    portef_value.set(portefeuille)
    portef_label.config(text = portef_value)
    tick_value.set(tick)
    tick_label.config(text=tick_value)
    if trend == "Stijgend":
        trend_label.config(text="Stijgend", bg = achtergrgroen, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    elif trend == "Dalend":
        trend_label.config(text="Dalend", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    elif trend == "Foutmelding":
        trend_label.config(text="FOUT", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    else:
        trend_label.config(text = "Neutraal", bg = achtergrkleur, fg = voorgrkleur,font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    # output BUY / SELL signaal
    if signaal == "BUY":
        signaal_label.config(text="BUY", bg = achtergrgroen, fg = contrastkleur,font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    elif signaal == "REentry BUY":
        signaal_label.config(text="BUY REentry", bg = achtergrgroen, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    elif signaal == "SELL":
        signaal_label.config(text="SELL", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    elif signaal == "REentry SELL":
        signaal_label.config(text="SELL REentry", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    elif signaal == "BUY EXIT":
        signaal_label.config(text="BUY EXIT",bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    elif signaal == "SELL EXIT":
        signaal_label.config(text="SELL EXIT",bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    else:
        signaal_label.config(text = signaal, bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
    # Output van graph in column 4
    t=tick.split(":")
    if trend == "Foutmelding":
        t[0] = "test"
        t[1] = "test"
    naam_plot = str(t[0])+"_"+str(t[1])+".png"
    naam_tab = str(t[0])+"_"+str(t[1])
    img = (Image.open(naam_plot))
    resized_img = img.resize((500, 500), Image.ANTIALIAS)
    new_image= ImageTk.PhotoImage(resized_img)
    image_label.config(image = new_image, bg = achtergrkleur)
    image_label.image = new_image
    exit_button = Button(window, text = "Exit", bg = "white", fg = contrastkleur, font = ("Arial",20), padx = 10, pady = 20, command = window.destroy)
    exit_button.grid(column = 4, row = row_nr, sticky=(E), padx=10, pady=10)
    return()


# functie voor creeren van de graphplot van koersgrafiek
def maak_graphplot(data, tick, number_of_bars, tspercent):
    global portef_label
    global tick_label
    global trend_label
    global signaal_label
    global image_label
    global fig
    global naam_plot
    global naam_tab
    #reset van variabelen
    today = dt.date.today()
    signaal_monday = today + dt.timedelta(days=-today.weekday(), weeks=round((today.weekday()-8)/7))
    # this_monday = today + dt.timedelta(days=-today.weekday(), weeks=0)
    signaal = "Neutraal"
    trend = "Neutraal"

    # naam voor saven van graph
    fig, ax = plt.subplots(1,1,figsize = (10,10))

    # plot de grafieken
    data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "close"].plot(linewidth = 4)
    data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "trailtop"].plot()
    data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "wma4"].plot()
    data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "wma62"].plot()

    # bepaal vertikale grenzen van de plot
    minimum_close = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "close"].min()
    minimum_trailtop = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "trailtop"].min()
    minimum_wma4 = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "wma4"].min()
    minimum_wma62 = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "wma62"].min()
    maximum_close = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "close"].max()
    maximum_trailtop = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "trailtop"].max()
    maximum_wma4 = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "wma4"].max()
    maximum_wma62 = data.loc[data.index[number_of_bars - 19]:data.index[number_of_bars - 1], "wma62"].max()
    minimum_y = round(min({minimum_close,minimum_trailtop,minimum_wma4,minimum_wma62})*0.9,4)
    maximum_y = round(max({maximum_close,maximum_trailtop,maximum_wma4,maximum_wma62})*1.1,4)
    range_y = round(maximum_y - minimum_y,4)
    # print(minimum_y,maximum_y,range_y,round(minimum_y/maximum_y,1))

    # afmetingen van graph
    # bottom,top = plt.ylim()
    # print(bottom,top)
    plt.ylim(minimum_y, maximum_y)
    # bottom,top = plt.ylim()
    # print(bottom,top)

    # voor gekleurde achtergrond
    x = data.index[95:114]
    plt.xticks(x, rotation = 70)
    groen = data.inlong[95:114]
    rood = data.inshort[95:114]

    n = number_of_bars - 19
    while n <= number_of_bars - 1:
        # bepaling van achtergrond kleur
        if (data.loc[data.index[n], "inlong"]):
            x = data.index[n:n+2]
            ax.fill_between(x,0,1, color = 'green', alpha = 0.1,transform=ax.get_xaxis_transform())
        # Plot de stoploss bolletjes
            if (data.loc[data.index[n], "trailtop"])*(1-tspercent) >= minimum_y:
                plt.text(data.index[n], data.loc[data.index[n], "trailtop"]*(1-tspercent), u"\u25CF", color = "red", ha = "center")
            trend = "Stijgend"
        # bepaling van de achtergrondkleur
        elif (data.loc[data.index[n], "inshort"]):
            x = data.index[n:n+2]
            ax.fill_between(x,0,1, color = 'red', alpha = 0.1,transform=ax.get_xaxis_transform())
        # Plot de stoploss bolletjes
            if (data.loc[data.index[n], "trailbot"])*(1+tspercent) <= maximum_y:
                plt.text(data.index[n], data.loc[data.index[n], "trailbot"]*(1+tspercent), u"\u25CF", color = "green", ha = "center")
            trend = "Dalend"
        else:
            trend = "Neutraal"

    # de opbouw van de plot met BUY, SELL, BUY EXIT, SELL EXIT, RE, RES signalen
    # --------- BUY signaal
        if (data.loc[data.index[n], "enterlong"] & (not(data.loc[data.index[n-1], "inlong"]))):
            plt.text(data.index[n], data.loc[data.index[n], "close"] - (0.1*range_y), u"\u25B2 \nBuy", color = "green", fontsize = 14, ha = 'center', weight = 'semibold')
            if pd.Timestamp(signaal_monday) == pd.to_datetime(data.index[n].date()):
                signaal = "BUY"
    # -------- SELL signaal
        if (data.loc[data.index[n], "entershort"]&(not(data.loc[data.index[n-1], "inshort"]))):
            plt.text(data.index[n], data.loc[data.index[n], "close"] + (0.1*range_y), u"Sell \n\u25BC", color = "red", fontsize = 14, ha = 'center', weight = 'semibold')
            last_sell_date = pd.to_datetime(data.index[n].date())
            if pd.Timestamp(signaal_monday) == pd.to_datetime(data.index[n].date()):
                signaal = "SELL"
    # -------- BUY EXIT signaal
        if (data.loc[data.index[n], "exitlong"] & data.loc[data.index[n-1], "inlong"] & (not(data.loc[data.index[n], "entershort"]))):
            plt.text(data.index[n], data.loc[data.index[n], "close"] + (0.1*range_y), u"Buy \nexit \n\u2193", fontsize = 14, ha = 'center', weight = 'semibold')
            if pd.Timestamp(signaal_monday) == pd.to_datetime(data.index[n].date()):
                signaal = "BUY EXIT"
    # -------- SELL EXIT signaal
        if (data.loc[data.index[n], "exitshort"] & data.loc[data.index[n-1], "inshort"] & (not(data.loc[data.index[n], "enterlong"]))):
            plt.text(data.index[n], data.loc[data.index[n], "close"] + (0.1*range_y), u"Sell \nexit \n\u2193", fontsize = 14, ha = 'center', weight = 'semibold')
            if pd.Timestamp(signaal_monday) == pd.to_datetime(data.index[n].date()):
                signaal = "SELL EXIT"
    # -------- REentry BUY signaal
        if (data.loc[data.index[n], "enterlong"] & data.loc[data.index[n], "inlong"] & data.loc[data.index[n-1], "inlong"]):
            plt.text(data.index[n], data.loc[data.index[n], "close"] - (0.1*range_y), u"\u2191 \nRe", color = "green", fontsize = 14, ha = 'center', weight = 'semibold')
            if pd.Timestamp(signaal_monday) == pd.to_datetime(data.index[n].date()):
                signaal = "REentry BUY"
    # -------- REentry SELL signaal
        if (data.loc[data.index[n], "entershort"] & data.loc[data.index[n], "inshort"] & data.loc[data.index[n-1], "inshort"]):
            plt.text(data.index[n], data.loc[data.index[n], "close"] + (0.1*range_y), u"Res \n\u2193", color = "red", fontsize = 14, ha = 'center', weight = 'semibold')
            if pd.Timestamp(signaal_monday) == pd.to_datetime(data.index[n].date()):
                signaal = "REentry SELL"
        n += 1
    # print(signaal)
    # print(tick)
    plt.legend(['Close', 'Trailtop', 'WMA4', 'WMA62'])
    plt.title(str(tick))
    t=tick.split(":")
    naam_plot = str(t[0])+"_"+str(t[1])+".png"
    naam_tab = str(t[0])+"_"+str(t[1])
    fig.savefig(naam_plot, bbox_inches='tight', dpi=150)
    plt.close(fig)
    # print("graph is gesaved")
    return(trend, signaal)


# functie voor uitvoer van aandeel gegevens naar juiste excel sheet
def uitvoer_excel(portefeuille, tick, trend, signaal):
    global datum_nu
    global book
    wsheet_buy = book.get_worksheet_by_name(wsheet_buy)
    wsheet_sell = book.get_worksheet_by_name(wsheet_sell)
    wsheet_geen = book.get_worksheet_by_name(wsheet_geen)  
    format_text= book.add_format({"align":"vcenter", "font_size":12})
    format_groot= book.add_format({"font_size":18})
    link = "book#" + naam_tab + "!A1"
    link_tick = "internal:" + naam_tab + "!A1"
    if signaal == "BUY" or signaal == "REentry BUY":
        wsheet_buy.write(stock_nr+1,0, datum_nu, format_text)
        wsheet_buy.write(stock_nr+1,1, portefeuille, format_text)
        #wsheet_buy.write(stock_nr+1,2, tick, format_text)
        wsheet_buy.write_url(stock_nr+1, 2, link_tick, format_text, tick)
        wsheet_buy.write(stock_nr+1,3, trend, format_text)
        wsheet_buy.write(stock_nr+1,4, signaal, format_text)
        wsheet_buy.insert_image(stock_nr+1, 5, naam_plot, {'x_scale': 0.3, 'y_scale': 0.3})
        wsheet_buy.cell(row = stock_nr+1, column = 5).hyperlink = link
        wsheet_image = book.add_worksheet(naam_tab)
        wsheet_image.write(1,0,tick, format_groot)
        wsheet_image.insert_image(0, 3, naam_plot, {'x_scale': 1, 'y_scale': 1})
        wsheet_image.write_url(0, 0, "internal:'buysignalen'!A1", format_text, "Home")
    elif signaal == "SELL" or signaal == "REentry SELL" or signaal == "BUY EXIT":
        wsheet_sell.write(stock_nr+1,0, datum_nu, format_text)
        wsheet_sell.write(stock_nr+1,1, portefeuille, format_text)
        wsheet_sell.write(stock_nr+1,2, tick, format_text)
        wsheet_sell.write(stock_nr+1,3, trend, format_text)
        wsheet_sell.write(stock_nr+1,4, signaal, format_text)
        wsheet_sell.insert_image(stock_nr+1, 5, naam_plot, {'x_scale': 0.3, 'y_scale': 0.3})
        wsheet_image = book.add_worksheet(naam_tab)
        wsheet_image.write(0,0,tick, format_groot)
        wsheet_image.insert_image(0, 3, naam_plot, {'x_scale': 1, 'y_scale': 1})
    else:
        wsheet_geen.write(stock_nr+1,0, datum_nu, format_text)
        wsheet_geen.write(stock_nr+1,1, portefeuille, format_text)
        wsheet_geen.write(stock_nr+1,2, tick, format_text)
        wsheet_geen.write(stock_nr+1,3, trend, format_text)
        wsheet_geen.write(stock_nr+1,4, signaal, format_text)
        wsheet_geen.insert_image(stock_nr+1, 5, naam_plot, {'x_scale': 0.3, 'y_scale': 0.3})
        wsheet_image = book.add_worksheet(naam_tab)
        wsheet_image.write(0,0,tick, format_groot)
        wsheet_image.insert_image(0, 3, naam_plot, {'x_scale': 1, 'y_scale': 1})
    return

""" Hierboven staan alle functies, voorzover niet ondergebracht in SFoutput()
Dan gaan we nu verder met het eigenllijke programma.
Overigens staan in de tkinter loop ook nog twee functies gedefinieerd"""

# LOGIN BIJ TRADINGVIEW
# username = 'FRB'
# password = 'sVVCh.MMLm8N'
# tv = TvDatafeed(username, password)
tv = TvDatafeed()

# Maak test png
# maak_test_file()

# haal lijst vvan stocks op uit excel
stock_list = get_stock_list()
number_of_stocks = len(stock_list) - 1
# print(number_of_stocks, stock_list)
stock_nr = 0

# setup en definitie van output window
window = Tk()
window.title("Overzicht van buy en sell signalen")
achtergrkleur = "dimgrey"
voorgrkleur = "white"
contrastkleur = "blue"
achtergrrood = "lightcoral"
achtergrgroen = "palegreen"
window.configure(background = achtergrkleur)
# window.geometry("1120x600+0+0")
window.geometry("1400x600+100+100")
# window.geometry("%dx%d+0+0" % (window.winfo_screenwidth(), window.winfo_screenheight()))

window.columnconfigure(0, weight=1, minsize = 160)
window.columnconfigure(1, weight=1, minsize = 160)
window.columnconfigure(2, weight=1)
window.columnconfigure(3, weight=1)
window.columnconfigure(4, weight=1, minsize = 160)

row_nr= 5

# commentaar in de tekst box, links boven
commentaar = "Resultaat per aandeel. \nHandelen op dit signaal is voor eigen verantwoordelijkheid"
Label(window, text = commentaar, bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10).grid(column=0, columnspan =2, row=0, sticky=(N, S, W, E))

# titels boven de kolommen
Label(window, text = "Portefeuille", bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10).grid(column=0, row=row_nr-4, sticky=(W, E))
Label(window, text = "Beurs:Aandeel", bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10).grid(column=0, row=row_nr-3, sticky=(W, E))
Label(window, text = "Koerstrend", bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10).grid(column=0, row=row_nr-2, sticky=(W, E))
Label(window, text = "Signaal", bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10).grid(column=0, row=row_nr-1, sticky=(W, E))
message_label = Label(window, text = " - ", bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20), padx=10, pady=10)
message_label.grid(column=0, columnspan = 4, row=row_nr, sticky=(W))

# output portefeuille
portef_value = StringVar()
portef_value.set(portefeuille)
portef_label = Label(window, textvariable=portef_value, bg = achtergrkleur, fg = voorgrkleur,font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
portef_label.grid(column=1, row=row_nr-4, sticky=(W,E))
# output exchange met ticker
tick_value = StringVar()
tick_value.set(tick)
tick_label = Label(window, textvariable=tick_value, bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
tick_label.grid(column=1, row=row_nr-3, sticky=(W,E))

# trend
trend_label = Label(window, text = "Neutraal", bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
trend_label.grid(column=1, row=row_nr-2, sticky=(W, E))
# print(trend)
# signaal
signaal_label = Label(window, text = "Neutraal", bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
signaal_label.grid(column = 1, row = row_nr-1, sticky=(W, E))
# print(signaal)


# graph plot
# Output van graph in column 4
t=tick.split(":")
naam_plot = str(t[0])+"_"+str(t[1])+".png"
naam_tab = str(t[0])+"_"+str(t[1])
img = (Image.open(naam_plot))
resized_img = img.resize((500, 500), Image.Resampling.LANCZOS)
# resized_img = img.resize((500, 500), Image.LANCZOS)
new_image= ImageTk.PhotoImage(resized_img)
image_label = Label(image = new_image, bg = "light grey")
image_label.image = new_image
image_label.grid(column = 2, columnspan = 3, row = 0, rowspan = 5, sticky = (N, S, W, E), padx = 20, pady = 10)

# function loop voor checken van alle aandelen en output naar window
def alle_aandelen():
    # definieer het excel bestand voor uitvoer van aandeel gegevens
    # creeer_excel()
    datum_nu = datetime.today().date()
    # file_name = resource_path("uitvoerexcel/"  + str(datum_nu) + ".xlsx")
    file_name = str(datum_nu) +".xlsx"
    # print(file_name)
    datum_nu = str(datum_nu)
    book = xlsxwriter.Workbook(file_name)
    format_head= book.add_format({"font_color":"white", "bg_color":"blue", "align":"center", "font_size":16})
    wsheet_buy = book.add_worksheet("buysignalen")
    wsheet_buy.set_column('A:E', 14)
    wsheet_buy.set_column("C:C", 19)
    wsheet_buy.set_column("F:F", 22)
    wsheet_buy.set_default_row(100)
    wsheet_buy.set_row(0, 24)
    wsheet_buy.write(0,0,"Datum", format_head)
    wsheet_buy.write(0,1,"Portefeuille", format_head)
    wsheet_buy.write(0,2,"Aandeel", format_head)
    wsheet_buy.write(0,3,"Trend", format_head)
    wsheet_buy.write(0,4,"Signaal", format_head)
    wsheet_buy.write(0,5,"Grafiek", format_head)
    wsheet_sell = book.add_worksheet("sell signalen")
    wsheet_sell.set_column('A:E', 14)
    wsheet_sell.set_column("C:C", 19)
    wsheet_sell.set_column("F:F", 22)
    wsheet_sell.set_default_row(100)
    wsheet_sell.set_row(0, 24)
    wsheet_sell.write(0,0,"Datum", format_head)
    wsheet_sell.write(0,1,"Portefeuille", format_head)
    wsheet_sell.write(0,2,"Aandeel", format_head)
    wsheet_sell.write(0,3,"Trend", format_head)
    wsheet_sell.write(0,4,"Signaal", format_head)
    wsheet_sell.write(0,5,"Grafiek", format_head)
    wsheet_geen = book.add_worksheet("zonder signaal")
    wsheet_geen.set_column('A:E', 14)
    wsheet_geen.set_column("C:C", 19)
    wsheet_geen.set_column("F:F", 22)
    wsheet_geen.set_default_row(100)
    wsheet_geen.set_row(0, 24)
    wsheet_geen.write(0,0,"Datum", format_head)
    wsheet_geen.write(0,1,"Portefeuille", format_head)
    wsheet_geen.write(0,2,"Aandeel", format_head)
    wsheet_geen.write(0,3,"Trend", format_head)
    wsheet_geen.write(0,4,"Signaal", format_head)
    wsheet_geen.write(0,5,"Grafiek", format_head)
   
    def doorgaan():
        global stock_nr
        global data
        global portef_label
        global tick_label
        global trend_label
        global signaal_label
        global rownr_buy
        global rownr_sell
        global rownr_geen
        global naam_plot
        global naam_tab
        
        # download de stock gegevens zoals portefeuille, trend, signaal en natuurlijk ticker
        portefeuille, tick, trend, signaal = download_process_data(stock_nr)
        # print(portefeuille, tick)
        # update het uitvoerscherm met de nieuwe stock gegevens die net zijn geprodiceerd
        # print("XXXXXX", portefeuille, tick, trend, signaal)
        portef_value.set(portefeuille)
        portef_label.config(text = portef_value)
        tick_value.set(tick)
        tick_label.config(text=tick_value)
        if trend == "Stijgend":
            trend_label.config(text="Stijgend", bg = achtergrgroen, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        elif trend == "Dalend":
            trend_label.config(text="Dalend", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        elif trend == "Foutmelding":
            trend_label.config(text="FOUT", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        else:
            trend_label.config(text = "Neutraal", bg = achtergrkleur, fg = voorgrkleur,font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        # output BUY / SELL signaal
        if signaal == "BUY":
            signaal_label.config(text="BUY", bg = achtergrgroen, fg = contrastkleur,font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        elif signaal == "REentry BUY":
            signaal_label.config(text="BUY REentry", bg = achtergrgroen, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        elif signaal == "SELL":
            signaal_label.config(text="SELL", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        elif signaal == "REentry SELL":
            signaal_label.config(text="SELL REentry", bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        elif signaal == "BUY EXIT":
            signaal_label.config(text="BUY EXIT",bg = achtergrrood, fg = contrastkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        elif signaal == "SELL EXIT":
            signaal_label.config(text="SELL EXIT",bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        else:
            signaal_label.config(text = signaal, bg = achtergrkleur, fg = voorgrkleur, font = ("Arial",20),borderwidth=2, relief="groove", padx=10, pady=10)
        # Output van graph in column 4
        t=tick.split(":")
        if trend == "Foutmelding":
            t[0] = "test"
            t[1] = "test"
        # print(tick, tick)
        naam_plot = str(t[0])+"_"+str(t[1])+".png"
        naam_tab = str(t[0])+"_"+str(t[1])
        img = (Image.open(naam_plot))
        resized_img = img.resize((500, 500), Image.Resampling.LANCZOS)
        # resized_img = img.resize((500, 500), Image.LANCZOS)
        new_image= ImageTk.PhotoImage(resized_img)
        image_label.config(image = new_image, bg = achtergrkleur)
        image_label.image = new_image
        # if stock_nr == number_of_stocks:
        #     exit_button = Button(window, text = "Exit", bg = achtergrkleur, fg = contrastkleur, font = ("Arial",20), padx = 10, pady = 20, command = window.destroy)
        #     exit_button.grid(column = 0, columnspan = 5, row = row_nr)
        # else:
        #     return
        
        # uitvoer_excel(portefeuille, tick, trend, signaal)
        format_text= book.add_format({'align': 'vcenter', "font_size": 14, 'text_wrap': True, 'border': 1})
        # format_text.set_center_across()
        # format_text = book.add_format({'text_wrap': True})
        format_groot= book.add_format({'align': 'vcenter', "font_size":18, 'text_wrap': True})
        link_tick = "internal:" + naam_tab + "!A1"
        if signaal == "BUY" or signaal == "REentry BUY":
            wsheet_buy.write(rownr_buy,0, datum_nu, format_text)
            wsheet_buy.write(rownr_buy,1, portefeuille, format_text)
            wsheet_buy.write_url(rownr_buy, 2, link_tick, format_text, "Klik hier voor:\n" + tick)
            wsheet_buy.write(rownr_buy,3, trend, format_text)
            wsheet_buy.write(rownr_buy,4, signaal, format_text)
            wsheet_buy.insert_image(rownr_buy, 5, naam_plot, {'x_scale': 0.2, 'y_scale': 0.2})
            rownr_buy += 1
            try:
                wsheet_image = book.add_worksheet(naam_tab)
            except:
                pass
            else:
                wsheet_image.set_column('A:A', 15)
                wsheet_image.write(1,0,portefeuille, format_groot)
                wsheet_image.write(2,0,tick, format_groot)
                wsheet_image.insert_image(0, 3, naam_plot, {'x_scale': 1, 'y_scale': 1})
                wsheet_image.write_url(0, 0, "internal:'buysignalen'!A1", format_groot, "Klik hier voor terug")
        elif signaal == "SELL" or signaal == "REentry SELL" or signaal == "BUY EXIT":
            wsheet_sell.write(rownr_sell,0, datum_nu, format_text)
            wsheet_sell.write(rownr_sell,1, portefeuille, format_text)
            wsheet_sell.write_url(rownr_sell, 2, link_tick, format_text, "Klik hier voor:\n" + tick)
            wsheet_sell.write(rownr_sell,3, trend, format_text)
            wsheet_sell.write(rownr_sell,4, signaal, format_text)
            wsheet_sell.insert_image(rownr_sell, 5, naam_plot, {'x_scale': 0.2, 'y_scale': 0.2})
            rownr_sell += 1
            try:
                wsheet_image = book.add_worksheet(naam_tab)
            except:
                pass
            else:
                wsheet_image.set_column('A:A', 15)
                wsheet_image.write(1,0,portefeuille, format_groot)
                wsheet_image.write(2,0,tick, format_groot)
                wsheet_image.insert_image(0, 3, naam_plot, {'x_scale': 1, 'y_scale': 1})
                wsheet_image.write_url(0, 0, "internal:'buysignalen'!A1", format_groot, "Klik hier voor terug")
        else:
            wsheet_geen.write(rownr_geen,0, datum_nu, format_text)
            wsheet_geen.write(rownr_geen,1, portefeuille, format_text)
            wsheet_geen.write(rownr_geen,2, tick, format_text)
            wsheet_geen.write_url(rownr_geen, 2, link_tick, format_text, "Klik hier voor:\n" + tick)
            wsheet_geen.write(rownr_geen,3, trend, format_text)
            wsheet_geen.write(rownr_geen,4, signaal, format_text)
            wsheet_geen.insert_image(rownr_geen, 5, naam_plot, {'x_scale': 0.2, 'y_scale': 0.2})
            rownr_geen += 1
            try:
                wsheet_image = book.add_worksheet(naam_tab)
            except:
                pass
            else:
                wsheet_image.set_column('A:A', 15)
                wsheet_image.write(1,0,portefeuille, format_groot)
                wsheet_image.write(2,0,tick, format_groot)
                wsheet_image.insert_image(0, 3, naam_plot, {'x_scale': 1, 'y_scale': 1})
                wsheet_image.write_url(0, 0, "internal:'buysignalen'!A1", format_groot, "Klik hier voor terug")
        # update_window(portefeuille, tick, trend, signaal)
        message(stock_nr, number_of_stocks, "")
        # delete de grafiek van dit aandeel
        stock_nr += 1
        if stock_nr <= number_of_stocks: 
            # print(tick, stock_nr)
            signaal_label.after(500, doorgaan)
        if stock_nr > number_of_stocks: 
            book.close()
            # os.remove png files
            verwijder_png_files()
        # print("image")
        if stock_nr > number_of_stocks:
            exit_button = Button(window, text = "Exit", bg = "white", fg = contrastkleur, font = ("Arial",20), padx = 10, pady = 10, command = window.destroy)
            exit_button.grid(column = 4, row = row_nr)
            # sluit het programma
            return
    doorgaan()
    
# nu de aanroep van de functie die door het hele lijst gaat
alle_aandelen()

window.mainloop()