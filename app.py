import argparse
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter
from io import BytesIO
from data_analysis import *

def main(file_path):
    # Load data
    mw, bw, info, price = load_and_process(file_path)

    # Create a new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('Financial_Analysis_Report.xlsx')
    worksheet1 = workbook.add_worksheet('Q1')
    worksheet2 = workbook.add_worksheet('Q2')
    worksheet3 = workbook.add_worksheet('Q3')
    worksheet4 = workbook.add_worksheet('Q4')

    worksheet1.write(1, 1, 'This sheet shows how concentrated/diversified the manager is (e.g., how many stocks they hold everyday, how they allocate weights differently across stocks)')
    worksheet1.write(2, 1, 'This chart shows the overall max, min, mean, and std of the weights of stocks the manager held during this time period')
    
    # Generate and save weight distributions plot and get dataframe
    fig1, df1 = weight_distributions(mw)
    write_dataframe_to_excel(worksheet1, df1, 4, 1)
    write_plot_to_excel(worksheet1, fig1, 'B12')

    # Generate and save stocks per day plot
    fig2 = stocks_per_day(mw)
    write_plot_to_excel(worksheet1, fig2, 'B38')

    worksheet2.write(1, 1, 'This sheet shows the net exposure (total weight) of the manager in different sectors over time.')
    # Generate net exposure over different sectors figure 
    fig3 = net_exposure(mw, info)
    write_plot_to_excel(worksheet2, fig3, 'B5')

    worksheet3.write(1, 1, 'This sheet shows the average excess exposures of the manager in different countries')
    # Generate bar chart showing manager's average excess exposures (active weight) in different countries
    fig4 = excess_exposure_by_country(mw, bw, info)
    write_plot_to_excel(worksheet3, fig4, 'B5')

    worksheet4.write(1, 1, 'This sheet shows the annual metrics of the manager.')
    worksheet3.write(2, 1, 'The table below has details on annual metrics and the largest Alpha drawdowns')
    # Generate annual metrics 
    fig5, df2, df3 = annual_metrics(mw, bw, price)
    write_dataframe_to_excel(worksheet4, df2, 1, 1)

    write_dataframe_to_excel(worksheet4, df3, 8, 1)

    write_plot_to_excel(worksheet4, fig5, 'B17')

    workbook.close()

    print("Plots and data table have been saved.")

def write_dataframe_to_excel(worksheet, dataframe, startrow, startcol):
    """ Writes a pandas DataFrame to an Excel worksheet starting at a given row and column. """
    # Write headers
    worksheet.write(startrow, startcol, 'Metric')
    for col_num, header in enumerate(dataframe.columns):
        worksheet.write(startrow, startcol + col_num + 1, header)
    
    # Write index and data
    for row_num, (index, row) in enumerate(dataframe.iterrows(), start=startrow + 1):
        worksheet.write(row_num, startcol, index)  # Write the index
        for col_num, item in enumerate(row):
            worksheet.write(row_num, startcol + col_num + 1, item)

def write_plot_to_excel(worksheet, fig, startcell):
    # Save the plot to a BytesIO buffer
    image_data = BytesIO()
    fig.savefig(image_data, format='png', bbox_inches="tight")
    image_data.seek(0)  # Rewind the buffer

    # Insert the plot into the worksheet
    worksheet.insert_image(startcell, 'plot.png', {'image_data': image_data})

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process financial data.')
    parser.add_argument('file_path', type=str, help='Path to the Excel file containing the data')

    args = parser.parse_args()
    main(args.file_path)
