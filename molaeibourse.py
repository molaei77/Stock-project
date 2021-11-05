import pytse_client as tse
import pandas as pd
import investpy

class stock():
    def __init__(self, Market, stockName):
        self.Market = Market
        self.stockName = stockName

    def saveData(self, startDate, finishDate):
        try:
            print("Please Wait...")
            if self.Market == "IRAN":
                tickers = tse.download(symbols=stockName)
                df = tickers[stockName]
                df['date'] = pd.to_datetime(df['date']).dt.date
                df = df.set_index(df['date'])
                df = df.sort_index()
                ticks = df[pd.to_datetime(startDate):pd.to_datetime(finishDate)]
            if self.Market == "US":
                search_result = investpy.search_quotes(text=stockName, products=['stocks'],
                                                       countries=['united states'], n_results=1)
                ticks = search_result.retrieve_historical_data(
                    from_date=pd.to_datetime(startDate).strftime('%d/%m/%Y'),
                    to_date=pd.to_datetime(finishDate).strftime('%d/%m/%Y'))
                ticks.reset_index(inplace=True)
                ticks['Date'] = pd.to_datetime(ticks['Date']).dt.date
                #ticks['Date'] = ticks.index
            ticks.to_excel(str(stockName) + '.xlsx', index=False, encoding='utf-8-sig')
            print(str(stockName) + " Excel file have been saved in current directory" + "\n")

        except Exception as e:
            print('There is Something Wrong')
            print(str(e))

while True:
    Market = input("Enter The Market You Want (IRAN, US) or press e to exit: ")
    if Market == 'e':
        break
    elif Market != "IRAN" and Market != "US":
        print('The Word Must Be "IRAN" or "US"')

    else:
        if Market == "IRAN":
            stockName = input("Enter Stock Name in Persian: ")
        else:
            stockName = input("Enter Stock Name Like (AAPL): ")
        Stock = stock(Market, stockName)
        startDate = input("Enter Start Date in (Y-M-D) Format: ")
        finishDate = input("Enter Finish Date in (Y-M-D) Format: ")
        Stock.saveData(startDate, finishDate)
        del Stock