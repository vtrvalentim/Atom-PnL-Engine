import numpy as np
import scipy
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib as mpl
import matplotlib.cm
import matplotlib.pylab as plb
import squarify
import os

def get_file_path(file_name):

    main_path = os.getcwd()
    file_path = main_path + '\outputs\Graphs' + '\\' + file_name

    return file_path

def plot_colors():
    volume_hex = '#008484'
    revenue_hex = '#4c57a7'
    profit_hex = '#824ca7'
    fourth_hex = '#DAE8B9'

    return volume_hex, revenue_hex, profit_hex, fourth_hex

def small_pct_to_other(data_slice_temp,channel,rev_or_profit,min_pct_for_other):

    data_slice = data_slice_temp

    if rev_or_profit == 'rev':
        if channel == 'TMALL':
            col_title = 'TMALL Total Revenue (RMB)'
        elif channel == 'HYPERBULK':
            col_title = 'HYPERBULK Total Revenue (RMB)'
        elif channel == 'WECHAT':
            col_title = 'WECHAT Total Revenue (RMB)'
        elif channel == 'JD FS':
            col_title = 'JD FS Total Revenue (RMB)'
    elif rev_or_profit == 'profit':
        if channel == 'TMALL':
            col_title = 'TMALL Total Profit (RMB)'
        elif channel == 'HYPERBULK':
            col_title = 'HYPERBULK Total Profit (RMB)'
        elif channel == 'WECHAT':
            col_title = 'WECHAT Total Profit (RMB)'
        elif channel == 'JD FS':
            col_title = 'JD FS Total Profit (RMB)'

    # Sum all values where col_title < min_pct_for_other%
    rows_bypct = data_slice.loc[(data_slice[col_title] < min_pct_for_other)]
    other_pct = rows_bypct[col_title].sum()

    # Remove all rows where col_title < min_pct_for_other%
    data_slice = data_slice.loc[(data_slice[col_title] >= min_pct_for_other)]

    # Apper Other SKUs row
    if other_pct != 0:
        data_slice = data_slice.append({'SKU Name' : 'Other SKUs',
                                        'TMALL Total Revenue (RMB)' : other_pct,
                                        'HYPERBULK Total Revenue (RMB)' : other_pct,
                                        'WECHAT Total Revenue (RMB)' : other_pct,
                                        'JD FS Total Revenue (RMB)' : other_pct,
                                        'TMALL Total Profit (RMB)' : other_pct,
                                        'HYPERBULK Total Profit (RMB)' : other_pct,
                                        'WECHAT Total Profit (RMB)' : other_pct,
                                        'JD FS Total Profit (RMB)' : other_pct},ignore_index=True)

    return data_slice

def small_pct_to_other_treemaps(data_slice_temp,total_or_month,vol_rev_or_profit,min_pct_for_other):

    data_slice = data_slice_temp

    if total_or_month == 'Total':
        if vol_rev_or_profit == 'vol':
            col_title = 'Total Volume'
        elif vol_rev_or_profit == 'rev':
            col_title = 'SKU Total Revenue (RMB)'
        elif vol_rev_or_profit == 'profit':
            col_title = 'SKU Total Profit (RMB)'
    else:
        col_title = total_or_month

    # Calculate min % in absolute value
    min_value = (data_slice[col_title].sum())*min_pct_for_other
    # Sum all values where col_title < min_value
    rows_bypct = data_slice.loc[(data_slice[col_title] < min_value)]
    other_pct = rows_bypct[col_title].sum()

    # Remove all rows where col_title < min_pct_for_other%
    data_slice = data_slice.loc[(data_slice[col_title] >= min_value)]

    # Append Other SKUs row
    if other_pct != 0:
        data_slice = data_slice.append({'SKU Name' : 'Other SKUs',
                                        col_title : other_pct}, ignore_index=True)

    return data_slice

def small_pct_to_other_barchannel(data_slice_temp,vol_rev_or_profit,min_pct_for_other):

    data_slice = data_slice_temp

    if vol_rev_or_profit == 'vol':
        total_col_title = 'Total Volume'
        hyp_col_title = 'Hyperbulk Volume'
        tmall_col_title = 'TMALL Volume'
        jd_col_title = 'JD FS Volume'
        wec_col_title = 'WeChat Volume'
    elif vol_rev_or_profit == 'rev':
        total_col_title = 'SKU Total Revenue (RMB)'
        hyp_col_title = 'HYPERBULK Total Revenue (RMB)'
        tmall_col_title = 'TMALL Total Revenue (RMB)'
        jd_col_title = 'JD FS Total Revenue (RMB)'
        wec_col_title = 'WECHAT Total Revenue (RMB)'
    elif vol_rev_or_profit == 'profit':
        total_col_title = 'SKU Total Profit (RMB)'
        hyp_col_title = 'HYPERBULK Total Profit (RMB)'
        tmall_col_title = 'TMALL Total Profit (RMB)'
        jd_col_title = 'JD FS Total Profit (RMB)'
        wec_col_title = 'WECHAT Total Profit (RMB)'

    # Calculate min % in absolute value
    min_value = (data_slice[total_col_title].sum())*min_pct_for_other
    # Sum all values where col_title < min_value
    rows_bypct = data_slice.loc[(data_slice[total_col_title] < min_value)]

    total_other_pct = rows_bypct[total_col_title].sum()

    hyp_other_pct = rows_bypct[hyp_col_title].sum()
    tmall_other_pct = rows_bypct[tmall_col_title].sum()
    jd_other_pct = rows_bypct[jd_col_title].sum()
    wec_other_pct = rows_bypct[wec_col_title].sum()


    # Remove all rows where col_title < min_pct_for_other%
    data_slice = data_slice.loc[(data_slice[total_col_title] >= min_value)]

    # Append Other SKUs row
    if total_other_pct != 0:
        data_slice = data_slice.append({'SKU Name' : 'Other SKUs',
                                        total_col_title : total_other_pct,
                                        hyp_col_title : hyp_other_pct,
                                        tmall_col_title : tmall_other_pct,
                                        jd_col_title : jd_other_pct,
                                        wec_col_title : wec_other_pct
                                        }, ignore_index=True)

    return data_slice

def genereate_legend_table(ax, colors, no_of_skus_to_graph, data_slice, is_vol=False):
    # __________________________________________________________________________________________________________________
    # DESCRIPTION
    # This functions generates a legend table to be plotted with the treemap plots

    # ------------------------------------------------------------------------------------------------------------------
    # ARGUMENTS
    # ax: subplot axis where the table is to be plotted

    # colors: color list object generated by matplotlibs color library

    # no_of_skus_to_graph: how many skus are being represented in the tree map

    # data_slice: the data to be written in the table

    # is_vol: flag to determine if currency format applies to numbers
    # __________________________________________________________________________________________________________________

    # Create hex color list from normalized color object
    hex_list = []
    for n in range(len(colors)):
        hex_list.append(mpl.colors.to_hex(colors[n]))

    # Create table object on desired axis
    legend_table = ax.table(cellText=data_slice.values,
                            colLabels=data_slice.columns,
                            loc='right',
                            colLoc='right',
                            colWidths=[0.2, 0.2],
                            edges='')

    # Change table text color
    i = 0
    while i <= no_of_skus_to_graph - 1:
        legend_table[(i + 1, 0)].get_text().set_color(hex_list[i])
        legend_table[(i + 1, 1)].get_text().set_color(hex_list[i])
        if is_vol == False:
            s = '¥{:,.2f}'.format(float(legend_table[(i + 1, 1)].get_text().get_text()))
            legend_table[(i + 1, 1)].get_text().set_text(s)
        i = i + 1

    if is_vol == True:
        month = legend_table[(0, 1)].get_text().get_text()
        legend_title = month + ' Volume'
        legend_table[(0, 1)].get_text().set_text(legend_title)

    # Return table
    return legend_table

def generate_treemaps(volume_sheet, unique_month_list):
    # __________________________________________________________________________________________________________________
    # DESCRIPTION
    # This functions generates the treemap data visualizations for volume, revenue and profit per SKU.

    # ------------------------------------------------------------------------------------------------------------------
    # ARGUMENTS
    # volume_sheet: table generated by run_algorithm containing the data to be displayed

    # unique_month_list: list twith the unique month year sets in the data
    # __________________________________________________________________________________________________________________
    # VARIABLES AND STYLE SETUP

    volume_hex, revenue_hex, profit_hex, fourth_hex = plot_colors()

    # Get last month column names
    i = len(unique_month_list)
    last_month_string = str(unique_month_list[i - 1])
    rev_string = 'Total Revenue ' + last_month_string + ' (RMB)'
    profit_string = 'Total Profit ' + last_month_string + ' (RMB)'

    plt.style.use('seaborn')

    figsize = (14, 10)

    volume_hex, revenue_hex, profit_hex, fourth_hex = plot_colors()

    # __________________________________________________________________________________________________________________
    # VOLUME BY SKU (UNITS SOLD) - ENTIRE HISTORY
    # Get relevant data
    data_slice = volume_sheet[['SKU Name','Total Volume']]
    data_slice = data_slice.sort_values(by=['Total Volume'],ascending=False)
    data_slice = small_pct_to_other_treemaps(data_slice,'Total','vol',0.015)

    # Create color map
    norm = mpl.colors.LogNorm(vmin=min(data_slice['Total Volume']),
                              vmax=max(data_slice['Total Volume']))
    colors = [mpl.cm.summer_r(norm(value)) for value in data_slice['Total Volume']]

    # Create and plot graph
    fig1, (ax11, ax12) = plt.subplots(nrows=2, ncols=1)
    fig1.set_size_inches(figsize)
    plt.rcParams['text.color'] = 'w'

    squarify.plot(label=data_slice['SKU Name'],
                  sizes=data_slice['Total Volume'],
                  color=colors,
                  alpha=.6,
                  ax=ax11)
    ax11.set_title("Volume by SKU (Units Sold) - Entire History",
                   fontsize=23,
                   fontweight="bold",
                   pad=13,
                   color=volume_hex)
    ax11.axis('off')


    genereate_legend_table(ax11,
                           colors,
                           len(data_slice['SKU Name']),
                           data_slice,
                           is_vol=True)

    # ------------------------------------------------------------------------------------------------------------------
    # VOLUME BY SKU (UNITS SOLD) - LAST MONTH
    # Get relevant data
    data_slice = volume_sheet[['SKU Name',
                               last_month_string]]
    data_slice = data_slice.sort_values(by=[last_month_string],
                                        ascending=False)
    data_slice = small_pct_to_other_treemaps(data_slice,last_month_string,'vol',0.015)

    # Create color map
    norm = mpl.colors.LogNorm(vmin=min(data_slice[last_month_string]),
                              vmax=max(data_slice[last_month_string]))
    colors = [mpl.cm.summer_r(norm(value)) for value in data_slice[last_month_string]]

    # Create and plot graph
    squarify.plot(label=data_slice['SKU Name'],
                  sizes=data_slice[last_month_string],
                  color=colors,
                  alpha=.6,
                  ax=ax12)
    ax12.set_title("Volume by SKU (Units Sold) - Last Month",
                   fontsize=23,
                   fontweight="bold",
                   pad=13,
                   color=volume_hex)
    ax12.axis('off')
    genereate_legend_table(ax12,
                           colors,
                           len(data_slice['SKU Name']),
                           data_slice,
                           is_vol=True)
    plt.tight_layout()
    plt.subplots_adjust(hspace=0.22, right=0.65)

    plt.savefig(get_file_path('Volume Charts'))
    # __________________________________________________________________________________________________________________
    # REVENUE BY SKU (RMB) - ENTIRE HISTORY
    # Get relevant data
    data_slice = volume_sheet[['SKU Name',
                               'SKU Total Revenue (RMB)']]
    data_slice = data_slice.sort_values(by=['SKU Total Revenue (RMB)'],
                                        ascending=False)
    data_slice = small_pct_to_other_treemaps(data_slice,'Total','rev',0.02)

    # Create color map
    norm = mpl.colors.LogNorm(vmin=min(data_slice['SKU Total Revenue (RMB)']),
                              vmax=max(data_slice['SKU Total Revenue (RMB)']))
    colors = [mpl.cm.cividis_r(norm(value)) for value in data_slice['SKU Total Revenue (RMB)']]

    # Create and plot graph
    fig2, (ax21, ax22) = plt.subplots(nrows=2, ncols=1)
    fig2.set_size_inches(figsize)
    squarify.plot(label=data_slice['SKU Name'],
                  sizes=data_slice['SKU Total Revenue (RMB)'],
                  color=colors,
                  alpha=.6,
                  ax=ax21)
    ax21.set_title("Revenue by SKU (RMB) - Entire History",
                   fontsize=23,
                   fontweight="bold",
                   pad=13,
                   color=revenue_hex)
    ax21.axis('off')
    genereate_legend_table(ax21,
                           colors,
                           len(data_slice['SKU Name']),
                           data_slice)
    # ------------------------------------------------------------------------------------------------------------------
    # REVENUE BY SKU (RMB) - LAST MONTH
    # Get relevant data
    data_slice = volume_sheet[['SKU Name',
                               rev_string]]
    data_slice = data_slice.sort_values(by=[rev_string],
                                        ascending=False)
    data_slice = small_pct_to_other_treemaps(data_slice,rev_string,'rev',0.02)

    # Create color map
    norm = mpl.colors.LogNorm(vmin=min(data_slice[rev_string]),
                              vmax=max(data_slice[rev_string]))
    colors = [mpl.cm.cividis_r(norm(value)) for value in data_slice[rev_string]]

    # Create and plot graph
    squarify.plot(label=data_slice['SKU Name'],
                  sizes=data_slice[rev_string],
                  color=colors,
                  alpha=.6,
                  ax=ax22)
    ax22.set_title("Revenue by SKU (RMB) - Last Month",
                   fontsize=23,
                   fontweight="bold",
                   pad=13,
                   color=revenue_hex)
    ax22.axis('off')
    genereate_legend_table(ax22,
                           colors,
                           len(data_slice['SKU Name']),
                           data_slice)
    plt.tight_layout()
    plt.subplots_adjust(hspace=0.22, right=0.65)

    plt.savefig(get_file_path('Revenue Charts'))
    # __________________________________________________________________________________________________________________

    # PROFIT BY SKU (RMB) - ENTIRE HISTORY
    # Get relevant data
    data_slice = volume_sheet[['SKU Name',
                               'SKU Total Profit (RMB)']]
    data_slice = data_slice.sort_values(by=['SKU Total Profit (RMB)'],
                                        ascending=False)
    data_slice = small_pct_to_other_treemaps(data_slice,'Total','profit',0.012)

    # Create color map
    norm = mpl.colors.LogNorm(vmin=min(data_slice['SKU Total Profit (RMB)']),
                              vmax=max(data_slice['SKU Total Profit (RMB)']))
    colors = [mpl.cm.viridis_r(norm(value)) for value in data_slice['SKU Total Profit (RMB)']]

    # Create and plot graph
    fig3, (ax31, ax32) = plt.subplots(nrows=2, ncols=1)
    fig3.set_size_inches(figsize)
    squarify.plot(label=data_slice['SKU Name'],
                  sizes=data_slice['SKU Total Profit (RMB)'],
                  color=colors,
                  alpha=.6,
                  ax=ax31)
    ax31.set_title("MACO by SKU (RMB) - Entire History",
                   fontsize=23,
                   fontweight="bold",
                   pad=13,
                   color=profit_hex)
    ax31.axis('off')
    genereate_legend_table(ax31,
                           colors,
                           len(data_slice['SKU Name']),
                           data_slice)
    # ------------------------------------------------------------------------------------------------------------------
    # PROFIT BY SKU (RMB) - LAST MONTH
    # Get relevant data
    data_slice = volume_sheet[['SKU Name',
                               profit_string]]
    data_slice = data_slice.sort_values(by=[profit_string],
                                        ascending=False)
    data_slice = small_pct_to_other_treemaps(data_slice,profit_string,'profit',0.012)

    # Create color map
    norm = mpl.colors.LogNorm(vmin=min(data_slice[profit_string]),
                              vmax=max(data_slice[profit_string]))
    colors = [mpl.cm.viridis_r(norm(value)) for value in data_slice[profit_string]]

    # Create and plot graph
    squarify.plot(label=data_slice['SKU Name'],
                  sizes=data_slice[profit_string],
                  color=colors,
                  alpha=.6,
                  ax=ax32)
    ax32.set_title("MACO by SKU (RMB) - Last Month",
                   fontsize=23,
                   fontweight="bold",
                   pad=13,
                   color=profit_hex)
    ax32.axis('off')
    genereate_legend_table(ax32,
                           colors,
                           len(data_slice['SKU Name']),
                           data_slice)
    plt.tight_layout()
    plt.subplots_adjust(hspace=0.22, right=0.65)

    plt.savefig(get_file_path('Profit Charts'))

    return

def generate_month_hist_graph(volume_sheet, unique_month_list):
    # __________________________________________________________________________________________________________________
    # DESCRIPTION
    # This functions generates the sku sales history data visualizations for volume, revenue and profit.

    # ------------------------------------------------------------------------------------------------------------------
    # ARGUMENTS
    # volume_sheet: table generated by run_algorithm containing the data to be displayed

    # unique_month_list: list twith the unique month year sets in the data
    # __________________________________________________________________________________________________________________
    plt.style.use('seaborn')
    no_of_skus_to_graph = 10
    volume_hex, revenue_hex, profit_hex, fourth_hex = plot_colors()

    # CREATE DATA SLICES

    # Order volume data
    volume_sheet = volume_sheet.sort_values(by=['SKU Total Revenue (RMB)'],ascending=False)
    volume_sheet = volume_sheet.iloc[:no_of_skus_to_graph]

    # Create lists with column titles
    # Empty lists
    vol_cols = []
    rev_cols = []
    profit_cols = []
    # Append first two columns
    vol_cols.append('SKU Name')
    rev_cols.append('SKU Name')
    profit_cols.append('SKU Name')
    # Append month columns
    for k in range(len(unique_month_list)):
        vol_cols.append(str(unique_month_list[k]))
        rev_cols.append('Total Revenue ' + str(unique_month_list[k]) + ' (RMB)')
        profit_cols.append('Total Profit ' + str(unique_month_list[k]) + ' (RMB)')
    # Slice data from volume sheet
    vol_by_month_slice = volume_sheet[vol_cols]
    rev_by_month_slice = volume_sheet[rev_cols]
    profit_by_month_slice = volume_sheet[profit_cols]

    # Fix name for graph
    rev_by_month_slice.columns = vol_cols
    profit_by_month_slice.columns = vol_cols

    # Set variables
    no_of_months = len(unique_month_list)
    sku_list = volume_sheet['SKU Name']

    # __________________________________________________________________________________________________________________

    # PLOT DATA SLICES

    # Transpose dataframes
    vol_by_month_slice = vol_by_month_slice.T
    rev_by_month_slice = rev_by_month_slice.T
    profit_by_month_slice = profit_by_month_slice.T

    # Change column names and drop first row
    vol_by_month_slice.columns = vol_by_month_slice.iloc[0]
    vol_by_month_slice = vol_by_month_slice.iloc[pd.RangeIndex(len(vol_by_month_slice)).drop(0)]

    rev_by_month_slice.columns = rev_by_month_slice.iloc[0]
    rev_by_month_slice = rev_by_month_slice.iloc[pd.RangeIndex(len(rev_by_month_slice)).drop(0)]

    profit_by_month_slice.columns = profit_by_month_slice.iloc[0]
    profit_by_month_slice = profit_by_month_slice.iloc[pd.RangeIndex(len(profit_by_month_slice)).drop(0)]

    figsize = (14, 12)
    fig4, (ax41, ax42, ax43) = plt.subplots(nrows=3, ncols=1)
    fig4.set_size_inches(figsize)

    # Create color schemes
    norm = mpl.colors.LogNorm(vmin=min(volume_sheet['Total Volume']),
                              vmax=max(volume_sheet['Total Volume']))
    vol_colors = [mpl.cm.nipy_spectral_r(norm(value)) for value in volume_sheet['Total Volume']]

    vol_hex_list = []
    for n in range(len(vol_colors)):
        vol_hex_list.append(mpl.colors.to_hex(vol_colors[n]))

    # Plot data
    for n in range(len(sku_list)):
        vol_by_month_slice[sku_list.iloc[n]].astype('float').plot(legend=True, ax=ax41, color =vol_hex_list[n] )
        rev_by_month_slice[sku_list.iloc[n]].astype('float').plot(legend=True, ax=ax42, color =vol_hex_list[n])
        profit_by_month_slice[sku_list.iloc[n]].astype('float').plot(legend=True, ax=ax43, color =vol_hex_list[n])

    ax41.set_title("Top 10 SKUs Volume by Month (Units Sold)",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color=volume_hex)
    ax42.set_title("Top 10 SKUs Revenue by Month (RMB)",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color=revenue_hex)
    ax43.set_title("Top 10 SKUs MACO by Month (RMB)",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color=profit_hex)
    ax42.legend(title='SKU Labels',
                loc='center left',
                bbox_to_anchor=(1.1, 0.5),
                frameon=False)
    ax41.legend().remove()
    ax43.legend().remove()

    ax42.get_yaxis().set_major_formatter(
        mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))
    ax43.get_yaxis().set_major_formatter(
        mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))

    plt.tight_layout()
    plt.subplots_adjust(hspace=0.22)

    plt.savefig(get_file_path('Sales per Month Charts'))
    # __________________________________________________________________________________________________________________
    return

def generate_hexbin_age_rrp(volume_sheet):
    # ______________________________________________________________________________________________
    # DESCRIPTION
    # This functions generates the age RRP volume correlation chart for revenue and profit per SKU.

    # -----------------------------------------------------------------------------------------------
    # ARGUMENTS
    # volume_sheet: table generated by run_algorithm containing the data to be displayed
    # ______________________________________________________________________________________________

    plt.style.use('seaborn')

    # Slice and clean data
    data_slice = volume_sheet[['Alcohol Category',
                               'SKU Name',
                               'Age',
                               'RRP (RMB)',
                               'SKU Total Revenue (RMB)',
                               'SKU Total Profit (RMB)']]
    data_slice = data_slice[data_slice['Alcohol Category'] == 'Whisky']
    data_slice = data_slice.drop(columns='Alcohol Category')

    surface_revenue = data_slice[['SKU Total Revenue (RMB)']].astype('double').values/4
    surface_profit = data_slice[['SKU Total Profit (RMB)']].astype('double').values/3



    figsize = (18, 10)
    fig5, (ax51, ax52) = plt.subplots(nrows=1, ncols=2)
    fig5.set_size_inches(figsize)

    # Revenue Colors
    norm = mpl.colors.LogNorm(vmin=min(data_slice['SKU Total Revenue (RMB)']),
                              vmax=max(data_slice['SKU Total Revenue (RMB)']))
    rev_colors = [mpl.cm.cividis_r(norm(value)) for value in data_slice['SKU Total Revenue (RMB)']]
    rev_hex_list = []
    for n in range(len(rev_colors)):
        rev_hex_list.append(mpl.colors.to_hex(rev_colors[n]))

    # Profit Colors
    norm = mpl.colors.LogNorm(vmin=min(data_slice['SKU Total Profit (RMB)']),
                              vmax=max(data_slice['SKU Total Profit (RMB)']))
    prof_colors = [mpl.cm.viridis_r(norm(value)) for value in data_slice['SKU Total Profit (RMB)']]
    prof_hex_list = []
    for n in range(len(prof_colors)):
        prof_hex_list.append(mpl.colors.to_hex(prof_colors[n]))

    ax51.scatter(x=data_slice[['RRP (RMB)']].values,
                 y=data_slice[['Age']].values,
                 s=surface_revenue,
                 alpha=0.7,
                 label='Revenue',
                 c=rev_colors)
    ax51.set_title("Age, RRP & Revenue Correlation",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color='#4c57a7',
                   )
    ax51.get_xaxis().set_major_formatter(
        mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))
    ax51.set_xlabel('RRP (RMB)')
    ax51.set_ylabel('Whisky Age')

    ax52.scatter(x=data_slice[['RRP (RMB)']].values,
                 y=data_slice[['Age']].values,
                 s=surface_profit,
                 alpha=0.7,
                 label='Revenue',
                 c=prof_colors)
    ax52.set_title("Age, RRP & MACO Correlation",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color='#824ca7')
    ax52.get_xaxis().set_major_formatter(
        mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))
    ax52.set_xlabel('RRP (RMB)')
    ax52.set_ylabel('Whisky Age')

    plt.savefig(get_file_path('Age RRP Correlation Charts'))
    return

def generate_sales_per_channel(volume_sheet):
    # __________________________________________________________________________________________________________________
    # DESCRIPTION
    # This functions generates the volume, revenue and profit distribution bar plots per channel.

    # ------------------------------------------------------------------------------------------------------------------
    # ARGUMENTS
    # volume_sheet: table generated by run_algorithm containing the data to be displayed
    # __________________________________________________________________________________________________________________
    # Set style variables
    plt.style.use('seaborn')
    no_of_skus_to_graph = 15
    volume_hex, revenue_hex, profit_hex, fourth_hex = plot_colors()

    # Slice and clean data
    data_slice = volume_sheet[['SKU Name',
                               'RRP (RMB)',
                               'TMALL Volume',
                               'JD FS Volume',
                               'WeChat Volume',
                               'Hyperbulk Volume',
                               'TMALL Total Revenue (RMB)',
                               'JD FS Total Revenue (RMB)',
                               'WECHAT Total Revenue (RMB)',
                               'HYPERBULK Total Revenue (RMB)',
                               'TMALL Total Profit (RMB)',
                               'JD FS Total Profit (RMB)',
                               'WECHAT Total Profit (RMB)',
                               'HYPERBULK Total Profit (RMB)',
                               'Total Volume',
                               'SKU Total Profit (RMB)',
                               'SKU Total Revenue (RMB)']]
    # __________________________________________________________________________________________________________________
    # PLOT BAR CHART PERCENT PER CHANNEL PER SKU

    # Set figure & plot variables
    fig6, (ax61, ax62, ax63) = plt.subplots(nrows=3, ncols=1)
    figsize = (18, 10)
    fig6.set_size_inches(figsize)
    font_dict = {'fontsize': 7}
    # ------------------------------------------------------------------------------------------------------------------
    # Slice volume data
    vol_data_slice = data_slice.sort_values(by=['Total Volume'],ascending=False)
    vol_data_slice = small_pct_to_other_barchannel(vol_data_slice,'vol',0.01)
    vol_x_axis = list(range(len(vol_data_slice['SKU Name'])))


    # Plot volume bars
    #ax61.bar(vol_x_axis,
    #         vol_data_slice['Hyperbulk Volume'].values,
    #         color=fourth_hex)
    ax61.bar(vol_x_axis,
             vol_data_slice['JD FS Volume'].values,
             color=profit_hex)
             #,bottom=vol_data_slice['Hyperbulk Volume'].values)
    ax61.bar(vol_x_axis,
             vol_data_slice['WeChat Volume'].values,
             color=revenue_hex,
             bottom=vol_data_slice['JD FS Volume'])
             #bottom=vol_data_slice['Hyperbulk Volume']+vol_data_slice['JD FS Volume'])
    ax61.bar(vol_x_axis,
             vol_data_slice['TMALL Volume'].values,
             color=volume_hex,
             bottom=vol_data_slice['JD FS Volume'].values + vol_data_slice['WeChat Volume'].values)
             #bottom=vol_data_slice['Hyperbulk Volume'].values+vol_data_slice['JD FS Volume'].values+ vol_data_slice['WeChat Volume'].values)



    # Configure volume plot layout
    ax61.set_xticks(vol_x_axis)
    ax61.set_xticklabels(vol_data_slice['SKU Name'].values,fontdict = font_dict)
    ax61.set_title("SKU Volume per Channel (Units Sold)",
               fontsize=16,
               fontweight="bold",
               pad=10,
               color=volume_hex)

    # ------------------------------------------------------------------------------------------------------------------
    # Slice revenue data
    rev_data_slice = data_slice.sort_values(by=['SKU Total Revenue (RMB)'],ascending=False)
    rev_data_slice = small_pct_to_other_barchannel(rev_data_slice, 'rev', 0.02)
    rev_x_axis = list(range(len(rev_data_slice['SKU Name'])))

    # Plot revenue bars
    #p4 = ax62.bar(rev_x_axis,
    #              rev_data_slice['HYPERBULK Total Revenue (RMB)'].values,
    #              color=fourth_hex)
    p2 = ax62.bar(rev_x_axis,
                  rev_data_slice['JD FS Total Revenue (RMB)'].values,
                  color=profit_hex)
                  #,bottom=rev_data_slice['HYPERBULK Total Revenue (RMB)'].values)
    p3 = ax62.bar(rev_x_axis,
                  rev_data_slice['WECHAT Total Revenue (RMB)'].values,
                  color=revenue_hex,
                  bottom=rev_data_slice['JD FS Total Revenue (RMB)'].values)
    p1 = ax62.bar(rev_x_axis,
                  rev_data_slice['TMALL Total Revenue (RMB)'].values,
                  color=volume_hex,
                  bottom=(rev_data_slice['WECHAT Total Revenue (RMB)'].values+\
                          rev_data_slice['JD FS Total Revenue (RMB)'].values))

    # Configure revenue plot layout
    ax62.set_xticks(rev_x_axis)
    ax62.set_xticklabels(rev_data_slice['SKU Name'].values,fontdict = font_dict)
    ax62.set_title("SKU Revenue per Channel (RMB)",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color=revenue_hex)
    # ------------------------------------------------------------------------------------------------------------------
    # Slice profit data
    prof_data_slice = data_slice.sort_values(by=['SKU Total Profit (RMB)'],ascending=False)
    prof_data_slice = small_pct_to_other_barchannel(prof_data_slice, 'profit', 0.02)
    prof_x_axis = list(range(len(prof_data_slice['SKU Name'])))

    # Plot profit bars
    #ax63.bar(prof_x_axis,
    #         prof_data_slice['HYPERBULK Total Profit (RMB)'].values,
    #         color=fourth_hex)
    ax63.bar(prof_x_axis,
             prof_data_slice['JD FS Total Profit (RMB)'].values,
             color=profit_hex)
    #        ,bottom=prof_data_slice['HYPERBULK Total Profit (RMB)'].values)
    ax63.bar(prof_x_axis,
             prof_data_slice['WECHAT Total Profit (RMB)'].values,
             color=revenue_hex,
             bottom=(prof_data_slice['JD FS Total Profit (RMB)'].values))
                     #+prof_data_slice['HYPERBULK Total Profit (RMB)'].values))
    ax63.bar(prof_x_axis,
             prof_data_slice['TMALL Total Profit (RMB)'].values,
             color=volume_hex,
             bottom=(prof_data_slice['WECHAT Total Profit (RMB)'].values+prof_data_slice['JD FS Total Profit (RMB)'].values))
                    #+prof_data_slice['HYPERBULK Total Profit (RMB)'].values))

    # Configure profit plot layout
    ax63.set_xticks(prof_x_axis)
    ax63.set_xticklabels(prof_data_slice['SKU Name'].values,fontdict = font_dict)
    ax63.set_title("SKU MACO per Channel (RMB)",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color=profit_hex)
    # ------------------------------------------------------------------------------------------------------------------
    # Configure overal plot layout
    # Legend
    #ax62.legend((p4[0], p2[0], p3[0], p1[0]), ('Hyperbulk', 'JD FS', 'WeChat', 'TMall'),
    ax62.legend(( p1[0], p3[0],p2[0]), ( 'TMall', 'WeChat','JD FS'),
            title='Chanel Labels',
            loc='center left',
            bbox_to_anchor=(1, 0.5),
            frameon=False)
    # Y axis with money format
    ax62.get_yaxis().set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))
    ax63.get_yaxis().set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))
    # Spacing
    plt.tight_layout()
    plt.subplots_adjust(hspace=0.35)

    plt.savefig(get_file_path('Total Sales per Channel Charts'))

    return

def no_to_pct(no_col):

    hundoP = no_col.sum()
    pct_mult = hundoP/100
    no_col = (no_col/hundoP)*100

    return no_col

def generate_channel_kip_charts(volume_sheet):
    # __________________________________________________________________________________________________________________
    # DESCRIPTION
    # This functions generates the volume, revenue and profit distribution charts per channel.

    # ------------------------------------------------------------------------------------------------------------------
    # ARGUMENTS
    # volume_sheet: table generated by run_algorithm containing the data to be displayed
    # __________________________________________________________________________________________________________________
    # Set style variables
    plt.style.use('seaborn')
    volume_hex, revenue_hex, profit_hex, fourth_hex = plot_colors()

    # Slice and clean data
    data_slice = volume_sheet[['SKU Name',
                               'TMALL Total Revenue (RMB)',
                               'JD FS Total Revenue (RMB)',
                               'WECHAT Total Revenue (RMB)',
                               'HYPERBULK Total Revenue (RMB)',
                               'TMALL Total Profit (RMB)',
                               'JD FS Total Profit (RMB)',
                               'WECHAT Total Profit (RMB)',
                               'HYPERBULK Total Profit (RMB)']]

    #hyp_data_slice = data_slice[data_slice['HYPERBULK Total Revenue (RMB)'] != 0]
    #hyp_data_slice = hyp_data_slice.sort_values(by=['HYPERBULK Total Revenue (RMB)'], ascending=False)

    jd_data_slice = data_slice[data_slice['JD FS Total Revenue (RMB)'] != 0]
    jd_data_slice = jd_data_slice.sort_values(by=['JD FS Total Revenue (RMB)'], ascending=False)

    tmall_data_slice = data_slice[data_slice['TMALL Total Revenue (RMB)'] != 0]
    tmall_data_slice = tmall_data_slice.sort_values(by=['TMALL Total Revenue (RMB)'], ascending=False)

    wec_data_slice = data_slice[data_slice['WECHAT Total Revenue (RMB)'] != 0]
    wec_data_slice = wec_data_slice.sort_values(by=['WECHAT Total Revenue (RMB)'], ascending=False)

    #hyp_data_slice['HYPERBULK Total Revenue (RMB)'] = no_to_pct(hyp_data_slice['HYPERBULK Total Revenue (RMB)'])
    jd_data_slice['JD FS Total Revenue (RMB)'] = no_to_pct(jd_data_slice['JD FS Total Revenue (RMB)'])
    tmall_data_slice['TMALL Total Revenue (RMB)'] = no_to_pct(tmall_data_slice['TMALL Total Revenue (RMB)'])
    wec_data_slice['WECHAT Total Revenue (RMB)'] = no_to_pct(wec_data_slice['WECHAT Total Revenue (RMB)'])

    #hyp_data_slice['HYPERBULK Total Profit (RMB)'] = no_to_pct(hyp_data_slice['HYPERBULK Total Profit (RMB)'])
    jd_data_slice['JD FS Total Profit (RMB)'] = no_to_pct(jd_data_slice['JD FS Total Profit (RMB)'])
    tmall_data_slice['TMALL Total Profit (RMB)'] = no_to_pct(tmall_data_slice['TMALL Total Profit (RMB)'])
    wec_data_slice['WECHAT Total Profit (RMB)'] = no_to_pct(wec_data_slice['WECHAT Total Profit (RMB)'])

    #hyp_rev_data_slice = small_pct_to_other(hyp_data_slice,'HYPERBULK','rev',4)
    #hyp_profit_data_slice = small_pct_to_other(hyp_data_slice,'HYPERBULK','profit',4)

    jd_rev_data_slice = small_pct_to_other(jd_data_slice, 'JD FS', 'rev',1.5)
    jd_profit_data_slice = small_pct_to_other(jd_data_slice, 'JD FS', 'profit',1.5)

    tmall_rev_data_slice = small_pct_to_other(tmall_data_slice, 'TMALL', 'rev',4)
    tmall_profit_data_slice = small_pct_to_other(tmall_data_slice, 'TMALL', 'profit',4)

    wec_rev_data_slice = small_pct_to_other(wec_data_slice, 'WECHAT', 'rev',4)
    wec_profit_data_slice = small_pct_to_other(wec_data_slice, 'WECHAT', 'profit',4)



    #fig7,((ax71,ax73,ax75,ax77),(ax72,ax74,ax76,ax78)) = plt.subplots(nrows=2,ncols=4)
    fig7, ((ax73, ax75, ax77), (ax74, ax76, ax78)) = plt.subplots(nrows=2, ncols=3)
    figsize = (18, 10)
    fig7.set_size_inches(figsize)

    theme_rev = plt.get_cmap('cividis')
    theme_profit = plt.get_cmap('viridis')

    #ax71.set_prop_cycle("color", [theme_rev(1. * i / len(hyp_rev_data_slice['HYPERBULK Total Revenue (RMB)']))
    #                              for i in range(len(hyp_rev_data_slice['HYPERBULK Total Revenue (RMB)']))])
#
    #ax72.set_prop_cycle("color", [theme_profit(1. * i / len(hyp_rev_data_slice['HYPERBULK Total Profit (RMB)']))
    #                              for i in range(len(hyp_rev_data_slice['HYPERBULK Total Profit (RMB)']))])
#
    ax73.set_prop_cycle("color", [theme_rev(1. * i / len(jd_rev_data_slice['JD FS Total Revenue (RMB)']))
                                  for i in range(len(jd_rev_data_slice['JD FS Total Revenue (RMB)']))])

    ax74.set_prop_cycle("color", [theme_profit(1. * i / len(jd_profit_data_slice['JD FS Total Profit (RMB)']))
                                  for i in range(len(jd_profit_data_slice['JD FS Total Profit (RMB)']))])

    ax75.set_prop_cycle("color", [theme_rev(1. * i / len(tmall_rev_data_slice['TMALL Total Revenue (RMB)']))
                                  for i in range(len(tmall_rev_data_slice['TMALL Total Revenue (RMB)']))])

    ax76.set_prop_cycle("color", [theme_profit(1. * i / len(tmall_profit_data_slice['TMALL Total Profit (RMB)']))
                                  for i in range(len(tmall_profit_data_slice['TMALL Total Profit (RMB)']))])

    ax77.set_prop_cycle("color", [theme_rev(1. * i / len(wec_rev_data_slice['WECHAT Total Revenue (RMB)']))
                                  for i in range(len(wec_rev_data_slice['WECHAT Total Revenue (RMB)']))])

    ax78.set_prop_cycle("color", [theme_profit(1. * i / len(wec_profit_data_slice['WECHAT Total Profit (RMB)']))
                                  for i in range(len(wec_profit_data_slice['WECHAT Total Profit (RMB)']))])

    plt.rcParams['text.color'] = 'w'

    #wedges71, labels71, autopct71 = ax71.pie(hyp_rev_data_slice['HYPERBULK Total Revenue (RMB)'].values,
    #                                         labels=hyp_rev_data_slice['SKU Name'].values,
    #                                         autopct='%1.1f%%',
    #                                         startangle=90,
    #                                         wedgeprops={"edgecolor":"w"},
    #                                         radius=1.3)
    #wedges72, labels72, autopct72 = ax72.pie(hyp_profit_data_slice['HYPERBULK Total Profit (RMB)'].values,
    #                                         labels=hyp_profit_data_slice['SKU Name'].values,
    #                                         autopct='%1.1f%%',
    #                                         startangle=90,
    #                                         wedgeprops={"edgecolor":"w"},
    #                                         radius=1.3)


    wedges73, labels73, autopct73 = ax73.pie(jd_rev_data_slice['JD FS Total Revenue (RMB)'].values,
                                             labels=jd_rev_data_slice['SKU Name'].values,
                                             autopct='%1.1f%%',
                                             startangle=90,
                                             wedgeprops={"edgecolor":"w"},
                                             radius=1.3)
    wedges74, labels74, autopct74 = ax74.pie(jd_profit_data_slice['JD FS Total Profit (RMB)'].values,
                                             labels=jd_profit_data_slice['SKU Name'].values,
                                             autopct='%1.1f%%',
                                             startangle=90,
                                             wedgeprops={"edgecolor":"w"},
                                             radius=1.3)

    wedges75, labels75, autopct75 = ax75.pie(tmall_rev_data_slice['TMALL Total Revenue (RMB)'].values,
                                             labels=tmall_rev_data_slice['SKU Name'].values,
                                             autopct='%1.1f%%',
                                             startangle=90,
                                             wedgeprops={"edgecolor":"w"},
                                             radius=1.3)
    wedges76, labels76, autopct76 = ax76.pie(tmall_profit_data_slice['TMALL Total Profit (RMB)'].values,
                                             labels=tmall_profit_data_slice['SKU Name'].values,
                                             autopct='%1.1f%%',
                                             startangle=90,
                                             wedgeprops={"edgecolor":"w"},
                                             radius=1.3)

    wedges77, labels77, autopct77 = ax77.pie(wec_rev_data_slice['WECHAT Total Revenue (RMB)'].values,
                                             labels=wec_rev_data_slice['SKU Name'].values,
                                             autopct='%1.1f%%',
                                             startangle=90,
                                             wedgeprops={"edgecolor":"w"},
                                             radius=1.3)
    wedges78, labels78, autopct78 = ax78.pie(wec_profit_data_slice['WECHAT Total Profit (RMB)'].values,
                                             labels=wec_profit_data_slice['SKU Name'].values,
                                             autopct='%1.1f%%',
                                             startangle=90,
                                             wedgeprops={"edgecolor":"w"},
                                             radius=1.3)

    #ax71.set_title("Hyperbulk Revenue Distribution",
    #               fontsize=16,
    #               fontweight="bold",
    #               pad=35,
    #               color=revenue_hex)
    #ax72.set_title("Hyperbulk Profit Distribution",
    #               fontsize=16,
    #               fontweight="bold",
    #               pad=35,
    #               color=profit_hex)

    ax73.set_title("JD FS Revenue Distribution",
                   fontsize=16,
                   fontweight="bold",
                   pad=35,
                   color=revenue_hex)
    ax74.set_title("JD FS MACO Distribution",
                   fontsize=16,
                   fontweight="bold",
                   pad=35,
                   color=profit_hex)
    ax75.set_title("Tmall Revenue Distribution",
                   fontsize=16,
                   fontweight="bold",
                   pad=35,
                   color=revenue_hex)
    ax76.set_title("Tmall MACO Distribution",
                   fontsize=16,
                   fontweight="bold",
                   pad=35,
                   color=profit_hex)
    ax77.set_title("WeChat Revenue Distribution",
                   fontsize=16,
                   fontweight="bold",
                   pad=35,
                   color=revenue_hex)
    ax78.set_title("WeChat MACO Distribution",
                   fontsize=16,
                   fontweight="bold",
                   pad=35,
                   color=profit_hex)



    label_fontsize = 8

    #plt.setp(labels71, fontsize=label_fontsize)
    #plt.setp(labels72, fontsize=label_fontsize)
    plt.setp(labels73, fontsize=label_fontsize)
    plt.setp(labels74, fontsize=label_fontsize)
    plt.setp(labels75, fontsize=label_fontsize)
    plt.setp(labels76, fontsize=label_fontsize)
    plt.setp(labels77, fontsize=label_fontsize)
    plt.setp(labels78, fontsize=label_fontsize)

    #plt.setp(labels71, color=revenue_hex)
    #plt.setp(labels72, color=profit_hex)
    plt.setp(labels73, color=revenue_hex)
    plt.setp(labels74, color=profit_hex)
    plt.setp(labels75, color=revenue_hex)
    plt.setp(labels76, color=profit_hex)
    plt.setp(labels77, color=revenue_hex)
    plt.setp(labels78, color=profit_hex)

    plt.tight_layout()
    plt.subplots_adjust(wspace=0.5)

    plt.savefig(get_file_path('SKU per Channel Pie Charts'))
    return


def generate_hexbin_age_rrp_TEMP(volume_sheet):
    # ______________________________________________________________________________________________
    # DESCRIPTION
    # This functions generates the age RRP volume correlation chart for revenue and profit per SKU.

    # -----------------------------------------------------------------------------------------------
    # ARGUMENTS
    # volume_sheet: table generated by run_algorithm containing the data to be displayed
    # ______________________________________________________________________________________________

    plt.style.use('seaborn')

    # Slice and clean data
    data_slice = volume_sheet[['Alcohol Category',
                               'SKU Name',
                               'Age',
                               'RRP (RMB)',
                               'SKU Total Revenue (RMB)',
                               'SKU Total Profit (RMB)']]
    data_slice = data_slice[data_slice['Alcohol Category'] == 'Whisky']
    data_slice = data_slice.drop(columns='Alcohol Category')

    surface_revenue = data_slice[['SKU Total Revenue (RMB)']].astype('double').values/4
    surface_profit = data_slice[['SKU Total Profit (RMB)']].astype('double').values/3



    figsize = (18, 10)
    fig5, ax52 = plt.subplots(nrows=1, ncols=1)
    fig5.set_size_inches(figsize)

    # Revenue Colors
    norm = mpl.colors.LogNorm(vmin=min(data_slice['SKU Total Revenue (RMB)']),
                              vmax=max(data_slice['SKU Total Revenue (RMB)']))
    rev_colors = [mpl.cm.cividis_r(norm(value)) for value in data_slice['SKU Total Revenue (RMB)']]
    rev_hex_list = []
    for n in range(len(rev_colors)):
        rev_hex_list.append(mpl.colors.to_hex(rev_colors[n]))

    # Profit Colors
    norm = mpl.colors.LogNorm(vmin=min(data_slice['SKU Total Profit (RMB)']),
                              vmax=max(data_slice['SKU Total Profit (RMB)']))
    prof_colors = [mpl.cm.viridis_r(norm(value)) for value in data_slice['SKU Total Profit (RMB)']]
    prof_hex_list = []
    for n in range(len(prof_colors)):
        prof_hex_list.append(mpl.colors.to_hex(prof_colors[n]))

    #ax51.scatter(x=data_slice[['RRP (RMB)']].values,
    #             y=data_slice[['Age']].values,
    #             s=surface_revenue,
    #             alpha=0.7,
    #             label='Revenue',
    #             c=rev_colors)
    #ax51.set_title("Age, RRP & Revenue Correlation",
    #               fontsize=16,
    #               fontweight="bold",
    #               pad=10,
    #               color='#4c57a7',
    #               )
    #ax51.get_xaxis().set_major_formatter(
    #    mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))
    #ax51.set_xlabel('RRP (RMB)')
    #ax51.set_ylabel('Whisky Age')

    ax52.scatter(x=data_slice[['RRP (RMB)']].values,
                 y=data_slice[['Age']].values,
                 s=surface_profit,
                 alpha=0.7,
                 label='Revenue',
                 c=prof_colors)
    ax52.set_title("Age, RRP & MACO Correlation",
                   fontsize=16,
                   fontweight="bold",
                   pad=10,
                   color='#824ca7')
    ax52.get_xaxis().set_major_formatter(
        mpl.ticker.FuncFormatter(lambda x, p: format(float(x), ',.2f')))
    ax52.set_xlabel('RRP (RMB)')
    ax52.set_ylabel('Whisky Age')

    plt.savefig(get_file_path('Age RRP Correlation Charts'))
    return