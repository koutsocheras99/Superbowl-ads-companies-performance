import yfinance as yf
import pandas as pd
from datetime import date, timedelta

# because the event happens at Sundays and the stock market is closed, I have added one day
year_date_event_dictionary = {2022: "2022-02-14", 2021: "2021-02-08", 2020: "2020-02-03", 2019: "2019-02-04", \
                              2018: "2018-02-05", 2017: "2017-02-06", 2016: "2016-02-08", 2015: "2015-02-02", \
                              2014: "2014-02-03", 2013: "2013-02-04", 2012: "2022-02-06", 2011: "2022-02-07"}

# not complete yet -- follow specific_year function practices which is completed
def get_all_years_companies():
    for sheet_index in range(12):
        print("Year " + str(sheet_index+2011))   
        dataframe = pd.read_excel('companies.xlsx', sheet_name=sheet_index, usecols='A')
        print(dataframe)

def get_specific_year_companies(year):
    non_duplicate_companies_symbols = []
    all_companies_symbols = []
    
    dataframe = pd.read_excel('companies.xlsx', sheet_name=year-2011, usecols='A')
    # iterrows() is not the most efficient way to iterate but the data size is pretty small
    for _, row in dataframe.iterrows():
        all_companies_symbols.append(row['Stock Symbol'])
        non_duplicate_companies_symbols.append(row['Stock Symbol']) if (row['Stock Symbol'] not in non_duplicate_companies_symbols \
                                                                    and row['Stock Symbol'] not in ["PRIVATE", "SERIES", "DELISTED"] ) else None
    '''
    print(len(all_companies_symbols))
    print(len(non_duplicate_companies_symbols))
    print(non_duplicate_companies_symbols)
    '''
    return non_duplicate_companies_symbols

# get_specific_year_companies(2022)

# unfortunately adding months (or bigger time objects) is not supported
# https://docs.python.org/3/library/datetime.html#timedelta-objects
def add_duration_to_a_date(year, weeks_in_the_future=0, days_in_the_future=0):
    starting_date_of_specific_year = date.fromisoformat(year_date_event_dictionary[year])
    final_date = starting_date_of_specific_year+timedelta(weeks=weeks_in_the_future, days=days_in_the_future)
    final_date_plus_one_day = starting_date_of_specific_year+timedelta(weeks=weeks_in_the_future, days=days_in_the_future+1) # +1 days due to yfinance's time period retrieval procedure
    return final_date, final_date_plus_one_day

# add_duration_to_a_date(2021, weeks_in_the_future=2, days_in_the_future=3)

def get_stock_open_and_close_prices(year, weeks_in_the_future=0, days_in_the_future=0):
    stock_symbols_list = get_specific_year_companies(year)
    stock_open_and_close_prices_dictionary = {}

    start_date = year_date_event_dictionary[year]
    end_date, end_date_plus_one_day = add_duration_to_a_date(year, weeks_in_the_future, days_in_the_future)

    for stock_symbol in stock_symbols_list:
        try:
            stock_symbol = yf.Ticker(stock_symbol)
            stock_symbol_price = stock_symbol.history(start=start_date, end=end_date_plus_one_day)

            starting_price = (stock_symbol_price['Open'].get(key=start_date))
            ending_price = (stock_symbol_price['Close'].get(key=str(end_date))) 

            stock_open_and_close_prices_dictionary[stock_symbol.info['shortName']] = [ending_price, starting_price]
        except Exception as e:
            print(e)

    return stock_open_and_close_prices_dictionary

# get_stock_open_and_close_prices(2011, weeks_in_the_future=4, days_in_the_future=3)

def get_percentage_difference(ending_price, starting_price):
    try:
        return(ending_price - starting_price) / starting_price * 100.0
    except ZeroDivisionError:
        print("Value is zero!")

# duration parameter is for how many days/weeks after the event performance you are interested 
def get_stock_performance(year, weeks_in_the_future=0, days_in_the_future=0):
    stocks = get_stock_open_and_close_prices(year, weeks_in_the_future=weeks_in_the_future, days_in_the_future=0)
    # print(stocks)
    for stock in stocks:
        ending_price, starting_price = stocks[stock][0], stocks[stock][1]
        percentage_difference = get_percentage_difference(ending_price, starting_price)
        print(stock + " ~~~ " + str(percentage_difference))
    
get_stock_performance(2011, weeks_in_the_future=4, days_in_the_future=3)