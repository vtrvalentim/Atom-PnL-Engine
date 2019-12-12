import pandas as pd
import matplotlib.pyplot as plt
import import_export_file as file
import tkinter as tki
import data_visualization as dvis

#_______________________________________________________________________________________________________________________
# SUPPORT FUNCTIONS

def click_import_sales_hist():
    # __________________________________________________________________________________________________________________
    # DESCRIPTION

    # This function handles the button to import the sales history file
    # It will print a success message if the import works
    # __________________________________________________________________________________________________________________

    # Import file
    global in_sales_hist
    in_sales_hist = tki.filedialog.askopenfilename(initialdir = "/",
                                                   title = "Select Hyperbulk Sales History",
                                                   filetypes = (("excel files", ".xlsx"),("all files","*.*")))
    # Print success message
    print(in_sales_hist)
    if in_sales_hist != '':
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END,
                             "_____________________________________________________________________________ \n", 'GREEN')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, "Successfully imported Hyperbulk sales history from \n", 'GREEN')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, in_sales_hist + "\n")
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END,
                             "_____________________________________________________________________________ \n", 'GREEN')
        display_frame.insert(tki.END, "\n")

    return

def click_import_sku_dict():
    # __________________________________________________________________________________________________________________
    # DESCRIPTION

    # This function handles the button to import the sku dictionary file
    # It will print a success message if the import works
    # __________________________________________________________________________________________________________________

    # Import file
    global in_sku_dict
    in_sku_dict = tki.filedialog.askopenfilename(initialdir="/",
                                                 title="Select SKU Dictionary",
                                                 filetypes= (("excel files", ".xlsx"),("all files","*.*")))
    # Print success message
    if in_sku_dict != '':
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, "_____________________________________________________________________________ \n",
                             'GREEN')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, "Successfully imported SKU dictionary from \n", 'GREEN')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, in_sku_dict + "\n")
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, "_____________________________________________________________________________ \n",
                             'GREEN')
        display_frame.insert(tki.END, "\n")

    return

def click_generate_sheets():
    # __________________________________________________________________________________________________________________
    # DESCRIPTION

    # This function handles the button to generate the P&L files
    # __________________________________________________________________________________________________________________

    # Get the variable values input by user
    igbp_to_rmb = float(gbp_to_rmb_entry.get())
    ieur_to_rmb = float(eur_to_rmb_entry.get())
    itmall_margin = float(tmall_margin_entry.get())
    ijd_margin = float(jd_margin_entry.get())
    iwechat_margin = float(wechat_margin_entry.get())
    iproof_margin = float(proof_margin_entry.get())
    iop_fees = float(op_fees_entry.get())
    ifreight = float(freight_entry.get())
    iif_rate = float(if_rate_entry.get())
    iconsumption_tax = float(consumption_tax_entry.get())
    ic_tax = float(c_tax_entry.get())
    iduty = float(duty_entry.get())
    ivat_rate = float(vat_rate_entry.get())

    # Run the algorithm to generate the files
    run_algorithm(igbp_to_rmb,
                  ieur_to_rmb,
                  itmall_margin,
                  ijd_margin,
                  iwechat_margin,
                  iproof_margin,
                  iop_fees,
                  ifreight,
                  iif_rate,
                  iconsumption_tax,
                  ic_tax,
                  iduty,
                  ivat_rate)

    return

def batch_number_finder(input_string):
    # __________________________________________________________________________________________________________________
    # DESCRIPTION

    # This function finds the batch number from the SKU description
    # to call it use batch_number_finder(SKU description)
    # ------------------------------------------------------------------------------------------------------------------
    # ARGUMENTS

    # input_string: description of the SKU where the batch number will be extracted from
    # __________________________________________________________________________________________________________________

    # Find char position inside input string
    char_batch = input_string.find('Batch')

    # Return null if no char found
    if char_batch == -1:
        return 'Null'

    # Extract batch number
    if (char_batch + 7) in range(len(input_string)):
        batch_number = str(input_string[char_batch + 6]) + str(input_string[char_batch + 7])
    else:
        batch_number = str(input_string[char_batch + 6])


    return batch_number

def handle_input_error():
    # __________________________________________________________________________________________________________________
    # DESCRIPTION

    # This function handles errors in the upload of input files
    # __________________________________________________________________________________________________________________

    # Handle missing sku dictionary file
    if in_sku_dict == 'Null' or in_sku_dict == '':
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END,
                             "_____________________________________________________________________________ \n",
                             'RED')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, "ERROR: missing file SKU dictionary \n", 'RED')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END,
                             "_____________________________________________________________________________ \n",
                             'RED')
        display_frame.insert(tki.END, "\n")
        return 1

    # Handle missing sales history file
    if in_sales_hist == 'Null' or in_sales_hist == '':
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END,
                             "_____________________________________________________________________________ \n",
                             'RED')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END, "ERROR: missing file Hyperbulk sales history \n", 'RED')
        display_frame.insert(tki.END, "\n")
        display_frame.insert(tki.END,
                             "_____________________________________________________________________________ \n",
                             'RED')
        display_frame.insert(tki.END, "\n")
        return 1

    return 0

# ______________________________________________________________________________________________________________________
# MAIN FUNCTION

def run_algorithm(gbp_to_rmb,eur_to_rmb,tmall_margin ,jd_margin,wechat_margin,proof_margin,op_fees,freight,if_rate,consumption_tax,c_tax,duty,vat_rate):
    # __________________________________________________________________________________________________________________
    # DESCRIPTION

    # This function runs the algorithm to generate the P&L files from the sales history and sku dictionary inputs
    # It will print a success message if the algorithm runs correctly
    #-------------------------------------------------------------------------------------------------------------------
    # ARGUMENTS

    # gbp_to_rmb: argument passed via input in the UI
    # eur_to_rmb: argument passed via input in the UI
    # tmall_margin: argument passed via input in the UI
    # jd_margin: argument passed via input in the UI
    # wechat_margin: argument passed via input in the UI
    # proof_margin: argument passed via input in the UI
    # op_fees: argument passed via input in the UI
    # freight: argument passed via input in the UI
    # if_rate: argument passed via input in the UI
    # consumption_tax: argument passed via input in the UI
    # c_tax: argument passed via input in the UI
    # duty: argument passed via input in the UI
    # vat_rate: argument passed via input in the UI
    # __________________________________________________________________________________________________________________
    # READING THE DATA

    # Handling input error
    if handle_input_error() == 1:
        return

    # Read sales history
    hyperbulk_data = pd.read_excel(in_sales_hist)

    # Read sku dictionary
    sku_dictionary = pd.read_excel(in_sku_dict)
    # __________________________________________________________________________________________________________________
    # JOINING & FORMATTING TABLES

    # Select rows with no atom sku code
    null_atomskucode_rows = hyperbulk_data[hyperbulk_data['ATOM编码'].isnull()]

    # Export these null rows
    file.export_data(null_atomskucode_rows, 'Rows_with_null_atom_code.xlsx')

    # TODO HYPERBULK FILTER (ON/OFF)
    hyperbulk_data = hyperbulk_data[hyperbulk_data['渠道 Channel'] != 'hyperbulk']

    # Left join sales history with sku dictionary on product ID
    raw_join_table = pd.merge(hyperbulk_data,
                              sku_dictionary,
                              left_on='ATOM编码',
                              right_on='ABM EXPORT ID',
                              how='left')
    clean_join_table = raw_join_table

    # Rename index column (column 0) to Index
    clean_join_table.index.names = ['Index']

    # Reorder and remove columns
    reordered_clean_table = clean_join_table[['ATOM编码','SKU 编号','款式 SKU','Alcohol Category','SKU Name',
                                              'SKU Description','年份 Age','EXW (NR)','VILC','MACO','零售金额 ',
                                              'Margin %','日期','月份 Month','渠道 Channel']]

    # Rename columns
    reordered_clean_table.columns = ['Atom SKU Code','Proof SKU Code','Hyperbulk SKU Code','Alcohol Category',
                                     'SKU Name','SKU Description','Age','EXW (GBP)','VILC (GBP)','MACO (GBP)',
                                     'RRP (RMB)','Margin %','Date','Month','Channel']
    # __________________________________________________________________________________________________________________
    # INSERTING BATCH COLUMN

    # Algorithm
    # 1 Load SKU description string
    # 2 Find 'Batch' substring with string.find('Batch') and return char index
    # 3 Get chars index+7 and +8 into new string
    # 4 Insert this new value into column "Batch"

    # Create empty column
    batch_values = pd.Series('Null',index=reordered_clean_table.index)

    # Populate column with batch numbers
    for i in range(len(reordered_clean_table)):
        batch_values[i] = batch_number_finder(str(reordered_clean_table.iloc[i][5]))

    # Insert batch column into reordered_clean_table
    reordered_clean_table.insert(7,'Batch',batch_values)
    # __________________________________________________________________________________________________________________
    # MODIFYING ALCOHOL CATEGORY TO SEPARATE SAMPLER PACKS

    # Create empty column
    alc_cat = pd.Series('Null',index=reordered_clean_table.index)

    # Fill column with new alcohol category values
    for k in range(len(reordered_clean_table)):
        if '5 x 5cl Set' in str(reordered_clean_table.iloc[k][5]):
            alc_cat[k] = 'Whisky Sampler'
        else:
            alc_cat[k] = reordered_clean_table.iloc[k][3]

    # Replace alcohol category column with the new values
    reordered_clean_table = reordered_clean_table.drop(columns = ['Alcohol Category'])
    reordered_clean_table.insert(3,'Alcohol Category',alc_cat)
    # __________________________________________________________________________________________________________________
    # CHANGE 5X5 CL SKU NAME

    # Create empty column
    new_sku_name = pd.Series('Null',index=reordered_clean_table.index)

    # Fill column with new SKU name values
    for n in range(len(reordered_clean_table)):
        if '5 x 5cl Set' in str(reordered_clean_table.iloc[n][5]):
            new_sku_name[n] = '5x5cl Set : ¥' + str(int(reordered_clean_table.iloc[n][11]))
        else:
            new_sku_name[n] = reordered_clean_table.iloc[n][4]

    # Replace alcohol category column with the new values
    reordered_clean_table = reordered_clean_table.drop(columns=['SKU Name'])
    reordered_clean_table.insert(4,'SKU Name',new_sku_name)
    # __________________________________________________________________________________________________________________
    # HANDLING DATE COLUMNS

    reordered_clean_table = reordered_clean_table.sort_values(by=['Date'], ascending=True)

    # Remove chinese character from month column
    reordered_clean_table['Month'] = reordered_clean_table['Month'].str[:-1]

    # Create empty column
    year_col = pd.Series('Null',index=reordered_clean_table.index)

    # Fill column with year values
    for x in range(len(reordered_clean_table)):
        year_col[x] = reordered_clean_table.iloc[x][13].year

    # Insert year column
    reordered_clean_table.insert(16,'Year',year_col)

    # Insert month+year column
    reordered_clean_table['Month Year'] = reordered_clean_table['Month'] + reordered_clean_table['Year'].map(str)
    # __________________________________________________________________________________________________________________
    # REMOVING ROWS WITH INSUFFICIENT DATA

    # Select rows with no reference on sku_dictionary
    null_join_rows_merged = reordered_clean_table[(reordered_clean_table['MACO (GBP)'].isnull()) & (reordered_clean_table['Atom SKU Code'].notnull())]

    # Export these null rows
    file.export_data(null_join_rows_merged,'Rows_with_null_MACO.xlsx')

    # Remove rows with no reference on sku_dictionary
    reordered_clean_table = reordered_clean_table[reordered_clean_table['MACO (GBP)'].notnull()]
    reordered_clean_table = reordered_clean_table[reordered_clean_table['Atom SKU Code'].notnull()]

    # Save filtered joint sales history as variable
    complete_sales_history = reordered_clean_table
    # __________________________________________________________________________________________________________________

    # ADD VOLUME PER CHANNEL COLUMNS

    # Create df with unique SKU IDs
    volume_map = pd.unique(reordered_clean_table['Atom SKU Code'])

    # Create series for each vol aggregate per channel col
    total_vol = pd.Series('Null', name='Total Volume',index=volume_map)
    tmall_vol = pd.Series('Null', name='TMALL Volume', index=volume_map)
    jdfs_vol = pd.Series('Null', name='JD FS Volume', index=volume_map)
    wechat_vol = pd.Series('Null', name='WeChat Volume', index=volume_map)
    hyperbulk_vol = pd.Series('Null', name='Hyperbulk Volume', index=volume_map)

    # Populate columns with respective aggregate volumes
    for j in range(len(volume_map)):
        # Total volume
        total_vol[volume_map[j]] = len(reordered_clean_table[reordered_clean_table['Atom SKU Code'] == (volume_map[j])])

        # Tmall volume
        tmall_vol[volume_map[j]] = len(reordered_clean_table[(reordered_clean_table['Atom SKU Code'] == (volume_map[j]))
                                                             & (reordered_clean_table['Channel'] == 'TMALL')])

        # JD volume
        jdfs_vol[volume_map[j]] = len(reordered_clean_table[(reordered_clean_table['Atom SKU Code'] == (volume_map[j]))
                                                             & (reordered_clean_table['Channel'] == 'JD FS')])

        # Wechat volume
        wechat_vol[volume_map[j]] = len(reordered_clean_table[(reordered_clean_table['Atom SKU Code'] == (volume_map[j]))
                                                             & (reordered_clean_table['Channel'] == 'WeChat')])

        # Hyperbulk volume
        hyperbulk_vol[volume_map[j]] = len(reordered_clean_table[(reordered_clean_table['Atom SKU Code'] == (volume_map[j]))
                                                             & (reordered_clean_table['Channel'] == 'hyperbulk')])

    # Join volume dfs into single mapping table
    full_vol_map = total_vol.to_frame().join(tmall_vol.to_frame()).join(jdfs_vol.to_frame()).join(wechat_vol.to_frame())\
        .join(hyperbulk_vol.to_frame())

    # __________________________________________________________________________________________________________________

    # ADD VOLUME PER MONTH COLUMNS

    # Algorithm:
    # For each value in unique month list
    # Create new column with aggregate sales values for that SKU
    # Join this column to full_vol_map

    # Get unique month list
    unique_month_list = pd.unique(reordered_clean_table['Month Year'])

    # Get number of unique months
    no_of_months = len(unique_month_list)

    # Loop over month list
    for y in range(len(unique_month_list)):
        month_string = unique_month_list[y]
        # Create empty series
        unique_month_vol = pd.Series('Null',
                                     name=month_string,
                                     index=volume_map)
        # Loop over SKU ids
        for z in range(len(volume_map)):
            # Populate series
            unique_month_vol[volume_map[z]] = len(reordered_clean_table[(reordered_clean_table['Atom SKU Code']==(volume_map[z]))
                                                                        &(reordered_clean_table['Month Year']==month_string)])
        # Join column
        full_vol_map = full_vol_map.join(unique_month_vol.to_frame())
    # __________________________________________________________________________________________________________________

    # FURTHER CLEANING AND REORGANIZING COLUMNS

    # Clean irrelevant columns
    reordered_clean_table = reordered_clean_table.drop(columns = ['Date', 'Month', 'Channel'])

    # Reordering rows before dropping duplicates
    reordered_clean_table['Hyperbulk SKU Code'] = reordered_clean_table['Hyperbulk SKU Code'].replace('Null', '99999999999')
    reordered_clean_table['Hyperbulk SKU Code'] = pd.to_numeric(reordered_clean_table['Hyperbulk SKU Code'])
    reordered_clean_table = reordered_clean_table.sort_values(by='Hyperbulk SKU Code')
    reordered_clean_table = reordered_clean_table.sort_values(by='Proof SKU Code')
    reordered_clean_table['Hyperbulk SKU Code'] = reordered_clean_table['Hyperbulk SKU Code'].replace(99999999999,'Null')

    # Drop rows with repeated SKUs
    reordered_clean_table = reordered_clean_table.drop_duplicates(subset = ['Atom SKU Code'], keep = 'first')
    # __________________________________________________________________________________________________________________

    # SET P&L VARIABLES

    # Currency conversion rates derived from user inputs
    rmb_to_gbp = 1/gbp_to_rmb #multiplier
    rmb_to_eur = 1/eur_to_rmb #multiplier
    gmb_to_eur = gbp_to_rmb/eur_to_rmb #multiplier
    eur_to_gmb = 1/gmb_to_eur #multiplier
    # __________________________________________________________________________________________________________________

    # CALCULATING P&L NUMBERS PER SKU

    # EXW (GBP)
    # Imported from data

    # EXW (RMB)
    reordered_clean_table['EXW (RMB)'] = reordered_clean_table['EXW (GBP)']*gbp_to_rmb

    # EXW VILC (GBP)
    reordered_clean_table = reordered_clean_table.rename(columns = {'VILC (GBP)':'EXW VILC (GBP)'})

    # EXW VILC (RMB)
    reordered_clean_table['EXW VILC (RMB)'] = reordered_clean_table['EXW VILC (GBP)']*gbp_to_rmb

    # EXW MACO (GBP)
    reordered_clean_table = reordered_clean_table.rename(columns = {'MACO (GBP)':'EXW MACO (GBP)'})

    # EXW MACO (RMB)
    reordered_clean_table['EXW MACO (RMB)'] = reordered_clean_table['EXW MACO (GBP)']*gbp_to_rmb

    # RRP (GBP)
    reordered_clean_table['RRP (GBP)'] = reordered_clean_table['RRP (RMB)']*rmb_to_gbp

    # RRP (RMB)
    # Imported from data

    # JD FS NR (GBP)
    reordered_clean_table['JD FS NR (GBP)'] = reordered_clean_table['RRP (GBP)']*(1-jd_margin)

    # JD FS NR (RMB)
    reordered_clean_table['JD FS NR (RMB)'] = reordered_clean_table['RRP (RMB)']*(1-jd_margin)

    # TMALL NR (GBP)
    reordered_clean_table['TMALL NR (GBP)'] = reordered_clean_table['RRP (GBP)']*(1-tmall_margin)

    # TMALL NR (RMB)
    reordered_clean_table['TMALL NR (RMB)'] = reordered_clean_table['RRP (RMB)']*(1-tmall_margin)

    # WECHAT NR (GBP)
    reordered_clean_table['WECHAT NR (GBP)'] = reordered_clean_table['RRP (GBP)']*(1-wechat_margin)

    # WECHAT NR (RMB)
    reordered_clean_table['WECHAT NR (RMB)'] = reordered_clean_table['RRP (RMB)']*(1-wechat_margin)

    # HYPERBULK NR (GBP)
    reordered_clean_table['HYPERBULK NR (GBP)'] = reordered_clean_table['EXW (GBP)']

    # HYPERBULK NR (RMB)
    reordered_clean_table['HYPERBULK NR (RMB)'] = reordered_clean_table['EXW (RMB)']

    # OP FEES (RMB)
    reordered_clean_table['OP FEES (RMB)'] = (reordered_clean_table['RRP (RMB)']*0)+op_fees

    # CIF (RMB)
    reordered_clean_table['CIF (RMB)'] = (reordered_clean_table['EXW (RMB)'] + (freight*eur_to_rmb)) * (1+if_rate)

    # CONSUMPTION TAX (RMB)
    reordered_clean_table['CONSUMPTION TAX (RMB)'] = (reordered_clean_table['CIF (RMB)'] * consumption_tax) + c_tax

    # DUTY (RMB)
    reordered_clean_table['DUTY (RMB)'] = reordered_clean_table['CIF (RMB)'] * duty

    # VAT (RMB)
    reordered_clean_table['VAT (RMB)'] = reordered_clean_table['CIF (RMB)'] * vat_rate

    # COGS (RMB)
    reordered_clean_table['COGS (RMB)'] = reordered_clean_table['OP FEES (RMB)'] + reordered_clean_table['CIF (RMB)'] +\
                                          reordered_clean_table['CONSUMPTION TAX (RMB)'] + \
                                          reordered_clean_table['DUTY (RMB)'] + reordered_clean_table['VAT (RMB)']

    # IMPORTER MARGIN (RMB)
    reordered_clean_table['IMPORTER MARGIN (RMB)'] = reordered_clean_table['COGS (RMB)'] * proof_margin

    # TOTAL COST (RMB)
    reordered_clean_table['TOTAL COST (RMB)'] = reordered_clean_table['COGS (RMB)'] + \
                                                reordered_clean_table['IMPORTER MARGIN (RMB)']

    # TOTAL COST (GBP)
    reordered_clean_table['TOTAL COST (GBP)'] = reordered_clean_table['TOTAL COST (RMB)'] * rmb_to_gbp

    # JD FS MACO (GBP)
    reordered_clean_table['JD FS MACO (GBP)'] = reordered_clean_table['JD FS NR (GBP)']\
                                                - reordered_clean_table['TOTAL COST (GBP)']

    # JD FS MACO (RMB)
    reordered_clean_table['JD FS MACO (RMB)'] = reordered_clean_table['JD FS NR (RMB)']\
                                                - reordered_clean_table['TOTAL COST (RMB)']

    # TMALL MACO (GBP)
    reordered_clean_table['TMALL MACO (GBP)'] = reordered_clean_table['TMALL NR (GBP)']\
                                                - reordered_clean_table['TOTAL COST (GBP)']

    # TMALL MACO (RMB)
    reordered_clean_table['TMALL MACO (RMB)'] = reordered_clean_table['TMALL NR (RMB)']\
                                                - reordered_clean_table['TOTAL COST (RMB)']

    # WECHAT MACO (GBP)
    reordered_clean_table['WECHAT MACO (GBP)'] = reordered_clean_table['WECHAT NR (GBP)']\
                                                 - reordered_clean_table['TOTAL COST (GBP)']

    # WECHAT MACO (RMB)
    reordered_clean_table['WECHAT MACO (RMB)'] = reordered_clean_table['WECHAT NR (RMB)']\
                                                 - reordered_clean_table['TOTAL COST (RMB)']

    # HYPERBULK MACO (GBP)
    reordered_clean_table['HYPERBULK MACO (GBP)'] = reordered_clean_table['EXW MACO (GBP)']

    # HYPERBULK MACO (RMB)
    reordered_clean_table['HYPERBULK MACO (RMB)'] = reordered_clean_table['EXW MACO (RMB)']
    # __________________________________________________________________________________________________________________
    # FIXING TABLE FORMAT AND CREATE SHEETS FOR OUTPUT

    # Reorder columns and create finance sheet
    finance_sheet = reordered_clean_table[['Atom SKU Code','Proof SKU Code','Hyperbulk SKU Code','Alcohol Category',
                                           'SKU Name','SKU Description','Age','Batch','EXW (GBP)','EXW (RMB)',
                                           'EXW VILC (GBP)','EXW VILC (RMB)','EXW MACO (GBP)','EXW MACO (RMB)',
                                           'Margin %','RRP (GBP)','RRP (RMB)','JD FS NR (GBP)','JD FS NR (RMB)',
                                           'TMALL NR (GBP)','TMALL NR (RMB)','WECHAT NR (GBP)','WECHAT NR (RMB)',
                                           'HYPERBULK NR (GBP)','HYPERBULK NR (RMB)','OP FEES (RMB)','CIF (RMB)',
                                           'CONSUMPTION TAX (RMB)','DUTY (RMB)','VAT (RMB)','COGS (RMB)',
                                           'IMPORTER MARGIN (RMB)','TOTAL COST (GBP)','TOTAL COST (RMB)',
                                           'JD FS MACO (GBP)','JD FS MACO (RMB)','TMALL MACO (GBP)','TMALL MACO (RMB)',
                                           'WECHAT MACO (GBP)','WECHAT MACO (RMB)','HYPERBULK MACO (GBP)', 'HYPERBULK MACO (RMB)']]

    # Create volume sheet
    volume_sheet = pd.merge(finance_sheet, full_vol_map, left_on='Atom SKU Code',right_index=True,how='left')
    volume_sheet=volume_sheet.drop(columns = ['Proof SKU Code','Hyperbulk SKU Code','SKU Description','Batch',
                                              'EXW (GBP)','EXW (RMB)','EXW VILC (GBP)','EXW VILC (RMB)',
                                              'EXW MACO (GBP)','EXW MACO (RMB)','Margin %',#'RRP (GBP)','RRP (RMB)',
                                              'OP FEES (RMB)','CIF (RMB)','CONSUMPTION TAX (RMB)','DUTY (RMB)',
                                              'VAT (RMB)','COGS (RMB)','IMPORTER MARGIN (RMB)','TOTAL COST (GBP)','TOTAL COST (RMB)'])
    # __________________________________________________________________________________________________________________
    # ADD REVENUE & PROFIT COLUMNS TO VOLUME SHEET

    # Revenue
    # By channel
    volume_sheet['JD FS Total Revenue (GBP)'] = volume_sheet['JD FS Volume'] * volume_sheet['JD FS NR (GBP)']
    volume_sheet['JD FS Total Revenue (RMB)'] = volume_sheet['JD FS Volume'] * volume_sheet['JD FS NR (RMB)']

    volume_sheet['TMALL Total Revenue (GBP)'] = volume_sheet['TMALL Volume'] * volume_sheet['TMALL NR (GBP)']
    volume_sheet['TMALL Total Revenue (RMB)'] = volume_sheet['TMALL Volume'] * volume_sheet['TMALL NR (RMB)']

    volume_sheet['WECHAT Total Revenue (GBP)'] = volume_sheet['WeChat Volume'] * volume_sheet['WECHAT NR (GBP)']
    volume_sheet['WECHAT Total Revenue (RMB)'] = volume_sheet['WeChat Volume'] * volume_sheet['WECHAT NR (RMB)']

    volume_sheet['HYPERBULK Total Revenue (GBP)'] = volume_sheet['Hyperbulk Volume'] * volume_sheet['HYPERBULK NR (GBP)']
    volume_sheet['HYPERBULK Total Revenue (RMB)'] = volume_sheet['Hyperbulk Volume'] * volume_sheet['HYPERBULK NR (RMB)']
    # By SKU
    volume_sheet['SKU Total Revenue (GBP)'] = volume_sheet['JD FS Total Revenue (GBP)'] + \
                                              volume_sheet['TMALL Total Revenue (GBP)'] + \
                                              volume_sheet['WECHAT Total Revenue (GBP)'] + \
                                              volume_sheet['HYPERBULK Total Revenue (GBP)']
    volume_sheet['SKU Total Revenue (RMB)'] = volume_sheet['JD FS Total Revenue (RMB)'] + \
                                              volume_sheet['TMALL Total Revenue (RMB)'] + \
                                              volume_sheet['WECHAT Total Revenue (RMB)'] + \
                                              volume_sheet['HYPERBULK Total Revenue (RMB)']

    # Profit
    # By channel
    volume_sheet['JD FS Total Profit (GBP)'] = volume_sheet['JD FS Volume'] * volume_sheet['JD FS MACO (GBP)']
    volume_sheet['JD FS Total Profit (RMB)'] = volume_sheet['JD FS Volume'] * volume_sheet['JD FS MACO (RMB)']

    volume_sheet['TMALL Total Profit (GBP)'] = volume_sheet['TMALL Volume'] * volume_sheet['TMALL MACO (GBP)']
    volume_sheet['TMALL Total Profit (RMB)'] = volume_sheet['TMALL Volume'] * volume_sheet['TMALL MACO (RMB)']

    volume_sheet['WECHAT Total Profit (GBP)'] = volume_sheet['WeChat Volume'] * volume_sheet['WECHAT MACO (GBP)']
    volume_sheet['WECHAT Total Profit (RMB)'] = volume_sheet['WeChat Volume'] * volume_sheet['WECHAT MACO (RMB)']

    volume_sheet['HYPERBULK Total Profit (GBP)'] = volume_sheet['Hyperbulk Volume'] * volume_sheet['HYPERBULK MACO (GBP)']
    volume_sheet['HYPERBULK Total Profit (RMB)'] = volume_sheet['Hyperbulk Volume'] * volume_sheet['HYPERBULK MACO (RMB)']
    # By SKU
    volume_sheet['SKU Total Profit (GBP)'] = volume_sheet['JD FS Total Profit (GBP)'] + \
                                             volume_sheet['TMALL Total Profit (GBP)'] + \
                                             volume_sheet['WECHAT Total Profit (GBP)'] + \
                                             volume_sheet['HYPERBULK Total Profit (GBP)']
    volume_sheet['SKU Total Profit (RMB)'] = volume_sheet['JD FS Total Profit (RMB)'] + \
                                             volume_sheet['TMALL Total Profit (RMB)'] + \
                                             volume_sheet['WECHAT Total Profit (RMB)'] + \
                                             volume_sheet['HYPERBULK Total Profit (RMB)']

    # By month
    for a in range(len(unique_month_list)):
        month_string = unique_month_list[a]
        column_titlegbp_string = 'Total Revenue ' + month_string + ' (GBP)'
        column_titlermb_string = 'Total Revenue ' + month_string + ' (RMB)'

        volume_sheet[column_titlegbp_string] = (volume_sheet[month_string] / volume_sheet['Total Volume']) * \
                                               volume_sheet['SKU Total Revenue (GBP)']
        volume_sheet[column_titlermb_string] = (volume_sheet[month_string] / volume_sheet['Total Volume']) * \
                                               volume_sheet['SKU Total Revenue (RMB)']

    for a in range(len(unique_month_list)):
        month_string = unique_month_list[a]
        column_titlegbp_string = 'Total Profit ' + month_string + ' (GBP)'
        column_titlermb_string = 'Total Profit ' + month_string + ' (RMB)'

        volume_sheet[column_titlegbp_string] = (volume_sheet[month_string] / volume_sheet['Total Volume']) * \
                                               volume_sheet['SKU Total Profit (GBP)']
        volume_sheet[column_titlermb_string] = (volume_sheet[month_string] / volume_sheet['Total Volume']) * \
                                               volume_sheet['SKU Total Profit (RMB)']
    # __________________________________________________________________________________________________________________
    # MAKE TABLE LOOK NICE

    # Export excel file
    file_path_format = file.export_nice_excel(volume_sheet, finance_sheet, complete_sales_history, no_of_months)

    # Print success message
    display_frame.insert(tki.END, "\n")
    display_frame.insert(tki.END, "_____________________________________________________________________________ \n",
                         'GREEN')
    display_frame.insert(tki.END, "\n")
    display_frame.insert(tki.END, "Successfully exported P&L data to \n", 'GREEN')
    display_frame.insert(tki.END, "\n")
    display_frame.insert(tki.END, file_path_format + "\n")
    display_frame.insert(tki.END, "\n")
    display_frame.insert(tki.END, "_____________________________________________________________________________ \n",
                         'GREEN')
    display_frame.insert(tki.END, "\n")
    # __________________________________________________________________________________________________________________
    # DATA VISUALIZATION

    dvis.generate_treemaps(volume_sheet,unique_month_list)

    dvis.generate_month_hist_graph(volume_sheet,unique_month_list)

    dvis.generate_hexbin_age_rrp(volume_sheet)

    dvis.generate_sales_per_channel(volume_sheet)

    dvis.generate_channel_kip_charts(volume_sheet)

    plt.show()
    # __________________________________________________________________________________________________________________
    # GENERATE PPT

    file.export_ppt()
    return

# ______________________________________________________________________________________________________________________
# USER INTERFACE

if __name__ == "__main__":
    
    # Set global variables
    global in_sales_hist
    global in_sku_dict

    in_sales_hist = 'Null'
    in_sku_dict = 'Null'

    # Main window config
    top = tki.Tk()
    # Window title
    top.title("P&L Generator")
    # Window background
    top.configure(background='white')

    # Frames Config
    # Variables frame
    variables_frame = tki.Frame(top,bg='white',padx=20,pady=10)
    variables_frame.grid(row = 0, column = 0)
    # Buttons frame
    buttons_frame = tki.Frame(top,bg='white')
    buttons_frame.grid(row = 1, column = 0)
    # Display frame
    global display_frame
    display_frame = tki.Text(top, bg='white',relief = 'solid')#, height=30, width=55)
    display_frame.grid(row=0, column=1, rowspan=2, pady = 10, padx = 10)
    display_frame.tag_config('RED', foreground='red')
    display_frame.tag_config('GREEN', foreground='#008484')

    # Defining variable input boxes and labels
    # Currency conversion
    gbp_to_rmb_label = tki.Label(variables_frame, text='Pound to Yuan ratio',background='white', anchor = 's', height = 2,width = 20)
    gbp_to_rmb_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')
    gbp_to_rmb_entry.insert(0,string='8.97')
    gbp_to_rmb_label.grid(row = 0, column = 0)
    gbp_to_rmb_entry.grid(row = 1, column = 0)

    eur_to_rmb_label = tki.Label(variables_frame, text='Euro to Yuan ratio',background='white', anchor = 's', height = 2)
    eur_to_rmb_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')
    eur_to_rmb_entry.insert(0,string='7.73')
    eur_to_rmb_label.grid(row = 2, column = 0)
    eur_to_rmb_entry.grid(row = 3, column = 0)

    # Margins
    tmall_margin_label = tki.Label(variables_frame,text='Tmall margin (%)',background='white',anchor='s',height=2,width=20)
    tmall_margin_entry = tki.Entry(variables_frame, foreground='#008484', justify='center') 
    tmall_margin_entry.insert(0,string='0.21')
    tmall_margin_label.grid(row = 0, column = 1)
    tmall_margin_entry.grid(row = 1, column = 1)

    jd_margin_label = tki.Label(variables_frame, text='JD FS margin (%)',background='white', anchor = 's', height = 2)
    jd_margin_entry = tki.Entry(variables_frame, foreground='#008484', justify='center') 
    jd_margin_entry.insert(0,string='0.22')
    jd_margin_label.grid(row = 2, column = 1)
    jd_margin_entry.grid(row = 3, column = 1)

    wechat_margin_label = tki.Label(variables_frame, text='WeChat margin (%)',background='white',anchor='s',height = 2)
    wechat_margin_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  
    wechat_margin_entry.insert(0,string='0.33')
    wechat_margin_label.grid(row = 4, column = 1)
    wechat_margin_entry.grid(row = 5, column = 1)

    proof_margin_label = tki.Label(variables_frame,text='Importer margin (%)',background='white',anchor='s',height = 2)
    proof_margin_entry = tki.Entry(variables_frame, foreground='#008484', justify='center') 
    proof_margin_entry.insert(0,string='0.18')
    proof_margin_label.grid(row = 6, column = 1)
    proof_margin_entry.grid(row = 7, column = 1)
#
    # Costs
    op_fees_label = tki.Label(variables_frame, text='Operating fees (RMB)',background='white', anchor = 's', height = 2,width = 20)
    op_fees_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  # 18.5
    op_fees_entry.insert(0,string='18.5')
    op_fees_label.grid(row = 0, column = 2)
    op_fees_entry.grid(row = 1, column = 2)

    freight_label = tki.Label(variables_frame, text='Freight (EUR)',background='white', anchor = 's', height = 2)
    freight_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  # 0.5
    freight_entry.insert(0,string='0.5')
    freight_label.grid(row = 2, column = 2)
    freight_entry.grid(row = 3, column = 2)

    c_tax_label = tki.Label(variables_frame, text='C tax (RMB)', background='white', anchor='s', height=2)
    c_tax_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  # 0.7104
    c_tax_entry.insert(0, string='0.7104')
    c_tax_label.grid(row=4, column=2)
    c_tax_entry.grid(row=5, column=2)

    if_rate_label = tki.Label(variables_frame, text='IF rate (%)',background='white', anchor = 's', height = 2,width = 20)
    if_rate_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  #0.003
    if_rate_entry.insert(0,string='0.003')
    if_rate_label.grid(row = 0, column = 3)
    if_rate_entry.grid(row = 1, column = 3)

    consumption_tax_label = tki.Label(variables_frame, text='Consumption tax rate (%)',background='white', anchor = 's', height = 2)
    consumption_tax_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  # 0.2625
    consumption_tax_entry.insert(0,string='0.2625')
    consumption_tax_label.grid(row = 2, column = 3)
    consumption_tax_entry.grid(row = 3, column = 3)

    duty_label = tki.Label(variables_frame, text='Duty rate (%)',background='white', anchor = 's', height = 2)
    duty_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  # 0.05
    duty_entry.insert(0,string='0.05')
    duty_label.grid(row = 4, column = 3)
    duty_entry.grid(row = 5, column = 3)

    vat_rate_label = tki.Label(variables_frame, text='VAT rate (%)',background='white', anchor = 's', height = 2)
    vat_rate_entry = tki.Entry(variables_frame, foreground='#008484', justify='center')  # 0.1706
    vat_rate_entry.insert(0,string='0.1706')
    vat_rate_label.grid(row = 6, column = 3)
    vat_rate_entry.grid(row = 7, column = 3)


    # Defining Buttons
    # Upload sku dictionary
    upload_sku_dict = tki.Button(buttons_frame, text='UPLOAD SKU DICTIONARY', bg='#008484', fg='#ffffff',
                                 activebackground='#ffffff',activeforeground='#008484', padx=10,relief = 'solid'
                                 ,command=click_import_sku_dict)
    upload_sku_dict.grid(row = 0, column = 0,padx=10)

    # Upload Hyperbulk sales history
    upload_sales_hist = tki.Button(buttons_frame, text='UPLOAD SALES HISTORY', bg='#008484', fg='#ffffff',
                                 activebackground='#ffffff',activeforeground='#008484',padx=10,relief = 'solid'
                                   ,command = click_import_sales_hist)
    upload_sales_hist.grid(row = 0, column = 1,padx=10)

    # Generate Sheets Button
    generate_sheets = tki.Button(buttons_frame, text='GENERATE SHEETS',bg='#008484',fg='#ffffff'
                                 ,activebackground='#ffffff',activeforeground='#008484',padx=10,relief = 'solid'
                                 ,command = click_generate_sheets)
    generate_sheets.grid(row = 0, column = 2,padx=10)


    top.mainloop()
# ______________________________________________________________________________________________________________________
