import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

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

#Calculate number of stocks held per day by the manager, store in a series
stocks_per_day = mw.groupby('date')['asset_id'].nunique()

#Plot the graph of number of stocks across a time series
stocks_per_day.plot(title='Number of Stocks Held Over Time')
plt.xlabel('Date')
plt.ylabel('Number of Stocks')
#plt.show()

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

1. Create a merged table that adds a country column to the benchmark table 
2. Create a merged table that adds a country column to the manager table
3. 
"""

mw_with_country = pd.merge(mw, info[['asset_id', 'country']], on=['asset_id'])
bw_with_country = pd.merge(bw, info[['asset_id', 'country']], on=['asset_id'])

manager_country_weights = mw_with_country.groupby(['date','country'])['weight'].sum().reset_index()
benchmark_country_weights = bw_with_country.groupby(['date','country'])['weight'].sum().reset_index()

print(manager_country_weights.head())
