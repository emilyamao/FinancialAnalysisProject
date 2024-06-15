import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

def load_and_process(filepath):
    #Load all four sheets of the excel file into pandas dataframes 
    mw = pd.read_excel(filepath, sheet_name='Manager Weights') 
    bw = pd.read_excel(filepath, sheet_name='Benchmark Weights')
    info = pd.read_excel(filepath, sheet_name='Asset Static')  #info: contains all information about the stocks (country, industry, sector etc.)
    price = pd.read_excel(filepath, sheet_name='Asset Price')  #price: contains close and return from each stock

    mw['date'] = pd.to_datetime(mw['date'])
    bw['date'] = pd.to_datetime(bw['date'])
    price['date'] = pd.to_datetime(price['date'])

    return mw, bw, info, price
    

#Clean and prepare data 
#Load all four sheets of the excel file into pandas dataframes 
mw = pd.read_excel('AA_problem_set_data.xlsx', sheet_name='Manager Weights') 
bw = pd.read_excel('AA_problem_set_data.xlsx', sheet_name='Benchmark Weights')
info = pd.read_excel('AA_problem_set_data.xlsx', sheet_name='Asset Static')  #info: contains all information about the stocks (country, industry, sector etc.)
price = pd.read_excel('AA_problem_set_data.xlsx', sheet_name='Asset Price')  #price: contains close and return from each stock

"""
QUESTION 1: 
Describe how concentrated/diversified the manager is (e.g., how many stocks they hold everyday, how they
allocate weights differently across stocks).

1. Create a series with asset_id as the index and shows the count of unique stocks they hold everyday
    a) Plot of number of stocks vs time
2. Create a weight_stats df that shows the max, min, mean, and std of the weights of stocks held each day
    a) Plot the StdDeb of weights over time, which shows how differently the manager assigns weights each day
    b) Create a table to show the overall mean of max, min, mean, and std for the whole year 
"""

#Turn 'date' column into datetime
mw['date'] = pd.to_datetime(mw['date'])

def stocks_per_day(mw):
    fig, ax = plt.subplots()
    stocks_per_day_series = mw.groupby('date')['asset_id'].nunique()
    stocks_per_day_series.plot(ax=ax, title='Number of Stocks Held Over Time')
    ax.set_xlabel('Date')
    ax.set_ylabel('Number of Stocks')
    return fig

#Analyze weight distribution each day
weight_stats = mw.groupby('date')['weight'].agg(['max', 'min', 'mean', 'std'])

#Plot the standard deviation of weights over time
weight_stats['std'].plot(title='Standard Deviation of Weights Over Time')
plt.xlabel('Date')
plt.ylabel('Standard Deviation of Weights')
#plt.show()

#Create a table for yearly stats on this manager's weight allocations
overall_stats = weight_stats.mean()

#Create a DataFrame for better visualization
overall_stats_df = pd.DataFrame(overall_stats, columns=['Yearly Average'])
#print(overall_stats_df)

"""
QUESTION 2: 
Plot the manager's net exposure (total weight) in different sectors over time.

1. Create a merged table that combines the sector and the asset_id
2. Plot the total weight in different sectors over time
"""

#Create a merged table of manager sectors to understand the diversification of this manager across sectors
manager_sectors = pd.merge(mw, info[['asset_id', 'sector']], 'left', on=['asset_id'])

#Create a series that groups 
daily_sector_weights = manager_sectors.groupby(['date', 'sector'])['weight'].sum().reset_index()

#Create a pivot table (row:dates, column: sector(each diff colored line), value: weight)
sector_pivot = daily_sector_weights.pivot(index='date', columns='sector', values='weight').fillna(0)

#Plot the sector weights over time, each line in a different color represents a different sector
ax = sector_pivot.plot(figsize=(10, 6), title='Sector Weights Over Time')
ax.set_xlabel('Date')
ax.set_ylabel('Weight')
plt.legend(title='Sector', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
#plt.show()

"""
QUESTION 3: 
Create a bar chart showing manager's average excess exposures (active weight) in different countries compared
with their benchmark.

1. Create a merged table that adds a country column to the manager table 
2. Create a merged table that adds a country column to the benchmark table
3. Create a combined table that calculates the active weight for each date and country
4. Plot the average active weight by the manager for each country
"""
#Merging the country information into both the manager and benchmark dfs
mw_with_country = pd.merge(mw, info[['asset_id', 'country']], on=['asset_id'])
bw_with_country = pd.merge(bw, info[['asset_id', 'country']], on=['asset_id'])

#Find the sum of the weight assigned for each country on each day
manager_country_weights = mw_with_country.groupby(['date','country'])['weight'].sum().reset_index()
benchmark_country_weights = bw_with_country.groupby(['date','country'])['weight'].sum().reset_index()

#Combine the manager and benchmark weights into one table by the date and country, with two weight columns, one for manager, one for bench
combined_country_weights = pd.merge(manager_country_weights, benchmark_country_weights, on=['date', 'country'], suffixes=('_m', '_b'))

combined_country_weights['active_weight'] = combined_country_weights['weight_m'] - combined_country_weights['weight_b']

#Group by the country, to get the overall average active_weight for the manager 
average_active = combined_country_weights.groupby('country')['active_weight'].mean().reset_index()

#Plot the bar chart by average active weight by the manager for each country over the year
fig, ax = plt.subplots(figsize=(14, 8))
ax.bar(average_active['country'], average_active['active_weight'], color='pink')
ax.set_xlabel('Country')
ax.set_ylabel('Average Active Weight')
ax.set_title('Manager’s Average Excess Exposures by Country')
ax.grid(True)
plt.xticks(rotation=45, ha="right", rotation_mode="anchor")
plt.tight_layout()
#plt.show()

"""
QUESTION 4: 
Compute manager's:
    (a) Annualized return: 
    (b) Annualized excess return over benchmark (Alpha):
    (c) Annualized Sharpe ratio:
    (d) Start date, end date, and magnitude of the top 3 Alpha drawdowns that do not overlap with each other"
"""

#First, add a daily returns column to price df by using the percent change formula from day to day 
price['daily_return'] = price.groupby('asset_id')['close'].pct_change()

price['date'] = pd.to_datetime(price['date'])
mw['date'] = pd.to_datetime(mw['date'])
bw['date'] = pd.to_datetime(bw['date'])

#Add ths daily_return column to the mw table, calculates the actual weighted return of the stock by weight, sums the daily return by date for the portfolio
mw_daily_returns = pd.merge(mw, price[['asset_id', 'date', 'daily_return']], on=['asset_id', 'date'])
mw_daily_returns['weighted_return'] = mw_daily_returns['weight'] * mw_daily_returns['daily_return']
portfolio_daily_return = mw_daily_returns.groupby('date')['weighted_return'].sum()

mean_daily_return = portfolio_daily_return.mean()

annualized_return = (1 + mean_daily_return) ** 252 - 1

#Find benchmark daily returns 
bw_daily_returns = pd.merge(bw, price[['asset_id', 'date', 'daily_return']], on=['asset_id', 'date'])
bw_daily_returns['weighted_return'] = bw_daily_returns['weight'] * bw_daily_returns['daily_return']
bench_portfolio_daily_return = bw_daily_returns.groupby('date')['weighted_return'].sum()

returns = pd.merge(portfolio_daily_return, bench_portfolio_daily_return, on=['date'], suffixes=('_m', '_b'))
returns['excess_return'] = returns['weighted_return_m'] - returns['weighted_return_b']
avg_excess_daily_return = returns['excess_return'].mean()
annualized_alpha = (1 + avg_excess_daily_return) ** 252 - 1

#Plot the benchmark returns vs the manager returns 
ax = returns.plot(figsize=(10, 6), title='Manager vs Benchmark Weighted Returns')
ax.set_xlabel('Date')
ax.set_ylabel('Returns')
plt.tight_layout()
#plt.show()

#Calculate the Sharpe Ratio
risk_free_rate = 0.02  # Assuming 2% risk-free rate
daily_std_dev = price['daily_return'].std()
annualized_std_dev = daily_std_dev * np.sqrt(252)
sharpe_ratio = (annualized_return - risk_free_rate) / annualized_std_dev


