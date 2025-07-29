import numpy as np
import json
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

def calculate_shares_to_buy(stocks, weights, prices, total_investment):

    if len(weights) != len(prices):
        raise ValueError("The length of weights and prices lists must be the same.")
    
    target_investments = np.array(weights) * total_investment

    shares_to_buy = target_investments / prices
    
    total_invested = 0
    shares_to_buy_rounded = np.zeros(len(shares_to_buy))
    
    for i in range(len(shares_to_buy)):
        shares_to_buy_rounded[i] = np.floor(shares_to_buy[i])
        total_invested += shares_to_buy_rounded[i] * prices[i]
        
    leeway = total_investment * 0.05
    lower_bound = total_investment - leeway
    upper_bound = total_investment + leeway
    
    if lower_bound <= total_invested <= upper_bound:
        return [(stocks[i], int(shares)) for i, shares in enumerate(shares_to_buy_rounded)]
    
    while total_invested < lower_bound or total_invested > upper_bound:
        for i, shares in enumerate(shares_to_buy_rounded):
            if total_invested > upper_bound and shares > 0:
                shares_to_buy_rounded[i] -= 1
                total_invested -= prices[i]
            elif total_invested < lower_bound:
                shares_to_buy_rounded[i] += 1
                total_invested += prices[i]
            if lower_bound <= total_invested <= upper_bound:
                break
        
        if not any(shares_to_buy_rounded):
            raise ValueError("Unable to match the investment target within the specified leeway.")

    return [(stocks[i], int(shares)) for i, shares in enumerate(shares_to_buy_rounded)]

# Example usage:
def main(file_name):

    with open(file_name, 'r') as file:
        data = json.load(file)

    portfolios = []
    for year, pf in data.items():
        portfolios.append(pf)

    portfolio_1_11_2023 = portfolios[-1]

    stock_data = pd.read_csv('Final Stock Data.csv', index_col='Date')
    stock_data.index = pd.to_datetime(stock_data.index)

    end_date = pd.Timestamp("1/1/2024")
    end_date = stock_data.index[stock_data.index.get_loc(end_date, method='nearest')]

    portfolio_stocks_data = stock_data[stock_data.index == end_date]

    weights = []
    stocks = []
    prices = []

    for stock,weight in portfolio_1_11_2023:
        weights.append(weight)
        stocks.append(stock)
        price = portfolio_stocks_data[stock].values[0]
        prices.append(price)

    total_investment = 40000  # Total investment amount

    shares_to_buy = calculate_shares_to_buy(stocks, weights, prices, total_investment)

    final_portfolio = []
    for stock, shares in shares_to_buy:
        if shares != 0:
            final_portfolio.append([stock, shares])

    return final_portfolio