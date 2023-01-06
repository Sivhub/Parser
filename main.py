import csv
import re

import xlsxwriter
import sys, getopt
import os.path

#
################################ Global App Data ###############################
#
PARSED = int(0)
NOT_PARSED = int(-1)
BOLD = int(1)
NOT_BOLD = int(0)
INCOMING = int(0)
OUTGOING = int(1)

inputfile = ''
outputfile = ''
filetype = ''
row= int(0)
GrandTotalIncoming = []
GrandTotalOutgoing = []
CheckSum = []

#
################################ IG Data ####################################
#

#
# IG INPUT.csv fields
#

IG_Date = int(0)
IG_Summary = int(1)
IG_MarketName = int(2)
IG_Period = int(3)
IG_PandL = int(4)
IG_Trans = int(5)
IG_Ref = int(6)
IG_Open = int(7)
IG_Close = int(8)
IG_Size = int(9)
IG_Currency = int(10)
IG_PL_Amount = int(11)
IG_Cash = int(12)
IG_Close_Date = int(13)
IG_Open_Date = int(14)
IG_ISO_Currency = int(15)

#
# IG OUTPUT.xlsx -> Transactions Workbook
#

IG_Trans_Date = int(0)
IG_Desc = int(1)
IG_Open_Price = int(2)
IG_Close_Price = int(3)
IG_Trans_Size = int(4)
IG_Total_Invested = int(5)
IG_Profit_Loss = int(6)
IG_Percent = int(7)
IG_Trans_Open_Date = int(8)
IG_Days = int(9)
IG_Gains = int(10)
IG_Loss = int(11)
IG_Gain_Percentage = int(12)
IG_Loss_Percentage = int(13)
IG_Gains_Sterling = int(14)
IG_Loss_Sterling = int(15)
IG_Days_Gain = int(16)
IG_Days_Loss = int(17)

#
# IG OUTPUT.xlsx -> Cost workbook
#
IG_Trans_Date = int(0)
IG_Desc = int(1)
IG_Costs_Trans = int(2)
IG_Costs_Amount = int(3)

# IG Data stores
#
IG_Deals = []
IG_Costs = []


#
################################ NW Data ####################################
#

# Transaction Fields
Date = int(0)
Transaction = int(1)
Desc = int(2)
Pout = int(3)
Pin = int(4)
Balance = int(5)

# NW Incoming Transactions
NW_Header = []
NW_Incoming_Trans = []
NW_Incoming_Trans_Total = []
NW_Incoming_Transfers = []
NW_Incoming_Transfers_Total = []

# NW Outgoing Transactions
NW_Bills = []
NW_Groceries = []
NW_Household = []
NW_General = []
NW_Food_Drink = []
NW_Personal_Care = []
NW_Experiences = []
NW_Shopping = []
NW_Transport = []
NW_OG_Transfers = []
NW_Other = []
NW_CheckSum = []
IC_TRANSFER = int(0)

#
# Dictionary created when IC_HEADER.cvs & OG_HEADER.cvs loaded
#
ic_types_dict = {}
og_types_dict = {}

#
############################ GLOBAL FUNCTIONS #####################################
#


def write_row_of_text_to_excel(worksheet, row, col, text, shade, border):

    for item in text:
        #print("Text: ",item)
        if(shade == 'BOLD'):
            worksheet.write(row, col, item, bold)
        elif(border == 'UNDERLINE'):
            worksheet.write(row, col, item, underline)
        else:
            worksheet.write(row, col, item)
        col+=1
    row+=1
    return(row)
#
############################ IG FUNCTIONS #####################################
#
def build_IG_excel_costs(worksheet,row,col,IG_Costs):

    IG_Costs.reverse()
    row = write_row_of_text_to_excel(worksheet, row, col,['Date', 'Description', 'Transaction', 'Charge'], 'NO_BOLD', 'UNDERLINE')

    for line in IG_Costs:
        #print('line: ',line)
        # Date
        worksheet.write(row, col, line[IG_Date], date_format)
        col += 1

        # Desc
        worksheet.write(row, col, line[IG_MarketName])
        col += 1

        # Trans
        worksheet.write(row, col, line[IG_Trans])
        col += 1

        # P&L
        #
        # 05/01/23. This took forever to work out
        # line[IG_PL_Amount] is a string. Cast as a float to cope with decimal places but first need to
        # eliminate the ',' from 2,000 etc
        #
        worksheet.write(row, col, float(re.sub(',','',line[IG_PL_Amount])),currency_format)
        col += 1

        row += 1
        col = 0

    # Sum costs
#    formula = 'SUM(D2:D' + str(row) + ')'
    formula = build_SUM_formula('D', 1, row)
    col = 3
    row +=1
    worksheet.write_formula(row,col,formula,currency_format)


# row = write_row_of_text_to_excel(row, col,[IG_Deals], 'NO_BOLD', 'UNDERLINE')

def format_date(UTC_date):

    x= str(UTC_date).split("T")
    y=str(x[0]).split("-")
    correct_date = y[2] + '-' + y[1] + '-' +y[0]
    #print('correct_date =',correct_date)
    return(correct_date)

def build_SUM_formula(col,start_row,end_row):

    formula = 'SUM(' + str(col) + str(start_row) + ':' + str(col) + str(end_row) + ')'
    #print('Formula = ',formula)
    return(formula)
def build_COUNT_formula(col,start_row,end_row):

    formula = 'COUNT(' + str(col) + str(start_row) + ':' + str(col) + str(end_row) + ')'
    #print('Formula = ', formula)
    return (formula)

def build_MAX_formula(col,start_row,end_row):
    formula = 'MAX(' + str(col) + str(start_row) + ':' + str(col) + str(end_row) + ')'
    #print('Formula = ',formula)
    return (formula)


def build_MIN_formula(col,start_row,end_row):
    formula = 'MIN(' + str(col) + str(start_row) + ':' + str(col) + str(end_row) + ')'
    # print('Formula = ',formula)
    return (formula)

def manipulate_UTC_dates(UTC_Open_Date,UTC_Close_Date):

    from datetime import date

    #print('UTC_Open_Date  = ', UTC_Open_Date)
    #print('UTC_Close_Date = ', UTC_Close_Date)

    op_date = str(UTC_Open_Date).split("T")
    open_date = str(op_date[0]).split("-")

    cl_date = str(UTC_Close_Date).split("T")
    close_date = str(cl_date[0]).split("-")

    #print('Open_Date  = ',open_date)
    #print('Close_Date = ', close_date)

    d1 = date(int(open_date[0]), int(open_date[1]), int(open_date[2]))
    d2 = date(int(close_date[0]), int(close_date[1]), int(close_date[2]))

    diff = d2 - d1

    #print('diff=', diff)

    return(diff.days)

def build_checksum_formula(checkSumRow,start_column,end_row,end_column):

    formula = 'IF((' + str(start_column) + str(checkSumRow) + '=' + str(end_column) + str(end_row) + '),"PASS","FAIL")'
#    print('formula = ',formula)
    return(formula)

def build_IG_Deals_Summary_row2(worksheet, row, col, IG_Deal,checkSumRow):

    gains_sterling = []
    gains_percent = []
    gains_days = []
    losses_sterling = []
    losses_percent = []
    losses_days = []

    for line in IG_Deal:

        days= manipulate_UTC_dates(line[IG_Open_Date], line[IG_Close_Date])
        # calcuate P&L amount remove ',' from 2,000 etc
        float_PL_Amount = float(re.sub(',','',line[IG_PL_Amount]))
        if(float(line[IG_Size])>0):
           total_invested = float(line[IG_Open]) * float(line[IG_Size])
        else:
           total_invested = float(line[IG_Open]) * (int(-1)*float(line[IG_Size]))

        if(float_PL_Amount>0):
           gains_sterling.append(float_PL_Amount)
           gains_percent.append(float_PL_Amount / total_invested)
           gains_days.append(days)
        else:
           losses_sterling.append(float_PL_Amount)
           losses_percent.append(float_PL_Amount / total_invested)
           losses_days.append(days)

        #print('total invested =', total_invested)

    number_of_trades = len(IG_Deal)
    #print('No of Trades = ', number_of_trades)

    print('*************************************************')
    print('*************** SUMMARY LISTS *******************')
    print('*************************************************')

    print('Gains_sterling = ', gains_sterling)
    print('Gains_percent = ', gains_percent)
    print('Gains_Days = ', gains_days)

    print('Losses_sterling = ', losses_sterling)
    print('Losses_percent = ', losses_percent)
    print('Losses_days = ', losses_days)


    no_of_gains = len(gains_sterling)
    no_of_losses = len(losses_sterling)

    sum_of_gains_sterling = sum(gains_sterling)
    sum_of_losses_sterling = sum(losses_sterling)

    sum_of_gains_percent = sum(gains_percent)
    sum_of_losses_percent = sum(losses_percent)

    sum_of_gains_days = sum(gains_days)
    sum_of_losses_days = sum(losses_days)

    print('*************************************************')
    print('*************** SUMMARY TOTALS ******************')
    print('*************************************************')

    print('No. of Gains = ', no_of_gains)
    print('No. of Losses = ', no_of_losses)

    print('Sum of Gains £= ', sum_of_gains_sterling)
    print('Sum of Losses £= ', sum_of_losses_sterling)

    print('Sum of Gains %= ', sum_of_gains_percent)
    print('Sum of Losses %= ', sum_of_losses_percent)

    print('Sum of Gains in Days =', sum_of_gains_days)
    print('Sum of Losses in Days =', sum_of_losses_days)



    if(no_of_gains>0):
        avgGainSterling = sum_of_gains_sterling / no_of_gains
    else:
        avgGainSterling = 0

    if(no_of_losses>0):
        avgLossSterling = sum_of_losses_sterling / no_of_losses
    else:
        avgLossSterling = 0


    if(no_of_gains>0):
        avgGain =  sum_of_gains_percent / no_of_gains
    else:
        avgGain = 0

    if(no_of_losses>0):
        avgLoss =  sum_of_losses_percent / no_of_losses
    else:
        avgLoss = 0

    if(no_of_gains>0):
        max_gain = max(gains_percent)
    else:
        max_gain = 0

    if(no_of_losses>0):
        max_loss = min(losses_percent)
    else:
        max_loss = 0

    if(no_of_gains>0):
        avgDaysGain = sum_of_gains_days / no_of_gains
    else:
        avgDaysGain = 0

    if(no_of_losses>0):
        avgDaysLosses = sum_of_losses_days / no_of_losses
    else:
        avgDaysLosses = 0

    if((no_of_gains>0) and (number_of_trades>0)):
        winPercent = no_of_gains / number_of_trades
    else:
        winPercent = 0

    total_rtn = sum(gains_sterling) + sum(losses_sterling)


    print('*************************************************')
    print('*************** SUMMARY RESULTS *****************')
    print('*************************************************')

    print('AVG GAIN = ', avgGain)
    print('AVG LOSS = ', avgLoss)
    print('WIN % = ', winPercent)
    print('No of Trades = ', number_of_trades)
    print('Max GAIN %=', max_gain)
    print('Max LOSS %=', max_loss)
    print('Avg G =', avgDaysGain)
    print('Avg L =', avgDaysLosses)
    print('AVG Profit = ', avgGainSterling)
    print('AVG Loss = ', avgLossSterling)
    print('P/L = ', total_rtn)

    #
    # Define columns for summary row #2
    #

    AVG_GAIN = int(1)
    AVG_LOSS = int(2)
    WIN_PERCENT = int(3)
    TOTAL_TRADES = int(4)
    MAX_GAIN = int(5)
    MAX_LOSS = int(6)
    AVG_DAYS_GAIN = int(7)
    AVG_DAYS_LOSS = int(8)
    AVG_GAIN_STERLING = int(9)
    AVG_LOSS_STERLING = int(10)
    PANDL = int(11)

    #
    # Create Summary row #2
    #

    write_row_of_text_to_excel(worksheet, row, AVG_GAIN, ['AVG GAIN'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, AVG_LOSS, ['AVG LOSS'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, WIN_PERCENT, ['WIN %'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, TOTAL_TRADES, ['TRADES'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, MAX_GAIN, ['MAX GAIN'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, MAX_LOSS, ['MAX LOSS'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, AVG_DAYS_GAIN, ['AVG G'], 'BOLD', 'NO_UNDERLINE')
    worksheet.write_comment(row, AVG_DAYS_GAIN, 'The Average number of days a Gain is held for')
    write_row_of_text_to_excel(worksheet, row, AVG_DAYS_LOSS, ['AVG L'], 'BOLD', 'NO_UNDERLINE')
    worksheet.write_comment(row, AVG_DAYS_LOSS, 'The Average number of days a Loss is held for')
    write_row_of_text_to_excel(worksheet, row, AVG_GAIN_STERLING, ['AVG Profit'], 'BOLD', 'NO_UNDERLINE')
    worksheet.write_comment(row, AVG_GAIN_STERLING, 'The Average Gain £')
    write_row_of_text_to_excel(worksheet, row, AVG_LOSS_STERLING, ['AVG loss'], 'BOLD', 'NO_UNDERLINE')
    worksheet.write_comment(row, AVG_LOSS_STERLING, 'The Average Loss £')
    write_row_of_text_to_excel(worksheet, row, PANDL, ['P&L'], 'BOLD', 'NO_UNDERLINE')

    #
    # Write Summary row #2 details
    #
    row +=1

    # AVG GAIN %
    worksheet.write(row,AVG_GAIN,avgGain,percentage_format)
    # AVG LOSS %
    worksheet.write(row,AVG_LOSS, avgLoss,percentage_format)
    # WIN %
    worksheet.write(row, WIN_PERCENT, winPercent, percentage_format)
    # TOTAL TRADES
    worksheet.write(row, TOTAL_TRADES, number_of_trades)
    # MAX GAIN
    worksheet.write(row, MAX_GAIN, max_gain,percentage_format)
    # MAX LOSS
    worksheet.write(row, MAX_LOSS, max_loss, percentage_format)
    # Average Days Gain %
    worksheet.write(row, AVG_DAYS_GAIN, avgDaysGain,no_decimals_format)
    # Average Days Losses %
    worksheet.write(row, AVG_DAYS_LOSS, avgDaysLosses,no_decimals_format)
    # Average Gain £
    worksheet.write(row, AVG_GAIN_STERLING, avgGainSterling, currency_format)
    # Average Loss £
    worksheet.write(row, AVG_LOSS_STERLING, avgLossSterling, currency_format)
    # P&L
    worksheet.write(row, PANDL, total_rtn, currency_format)

    #
    # Perform Checksum
    #
    row += 1
    #formula = build_checksum_formula(checkSumRow, 'E', row, 'F')

    # AVG_GAIN
    worksheet.write_formula(row, AVG_GAIN, build_checksum_formula(checkSumRow, 'M', row, 'B'), percentage_format)
    worksheet.write_formula(row, AVG_LOSS, build_checksum_formula(checkSumRow, 'N', row, 'C'), percentage_format)
    worksheet.write_formula(row, WIN_PERCENT, build_checksum_formula(checkSumRow, 'H', row, 'D'), percentage_format)
    worksheet.write_formula(row, TOTAL_TRADES, build_checksum_formula(checkSumRow, 'B', row, 'E'), percentage_format)
    worksheet.write_formula(row, MAX_GAIN, build_checksum_formula(checkSumRow, 'E', row, 'F'), percentage_format)
    worksheet.write_formula(row, MAX_LOSS, build_checksum_formula(checkSumRow, 'F', row, 'G'), percentage_format)
    worksheet.write_formula(row, AVG_DAYS_GAIN, build_checksum_formula(checkSumRow, 'Q', row, 'H'), percentage_format)
    worksheet.write_formula(row, AVG_DAYS_LOSS, build_checksum_formula(checkSumRow, 'R', row, 'I'), percentage_format)
    worksheet.write_formula(row, AVG_GAIN_STERLING, build_checksum_formula(checkSumRow, 'O', row, 'J'), percentage_format)
    worksheet.write_formula(row, AVG_LOSS_STERLING, build_checksum_formula(checkSumRow, 'P', row, 'K'),percentage_format)
    worksheet.write_formula(row, PANDL, build_checksum_formula(checkSumRow, 'G', row, 'L'), percentage_format)

    # Add Charts
    title_row = row - 1
    data_row = row
    row += 3

    # Create AVG Win / AVG Loss pie chart

    catStr = '=Transactions!$B$' + str(title_row) + ':' + '$C$' + str(title_row)
    valStr = '=Transactions!$B$' + str(data_row) + ':' + '$C$' + str(data_row)

    chart = workbook.add_chart({'type': 'pie'})

    # Add a series to the chart.
    chart.add_series({
        'name': 'Average Gain & Loss %',
        'categories': catStr,
        'values': valStr,
        'data_labels': {
            'value': True,
            'leader_lines': True
            },
        'points': [
            {'fill': {'color': 'green'}},
            {'fill': {'color': 'red'}},
        ],
    })

    # Insert the chart into the worksheet.
    worksheet.insert_chart('A'+str(row), chart)

    #
    # Create Max Gain / Max Loss pie chart
    #

    catStr = '=Transactions!$F$' + str(title_row) + ':' + '$G$' + str(title_row)
    valStr = '=Transactions!$F$' + str(data_row) + ':' + '$G$' + str(data_row)

    chart1 = workbook.add_chart({'type': 'pie'})

    # Add a series to the chart.
    chart1.add_series({
        'name': 'Max Gain & Loss %',
        'categories': catStr,
        'values': valStr,
        'data_labels': {
            'value': True,
            'leader_lines': True
        },
        'points': [
            {'fill': {'color': 'green'}},
            {'fill': {'color': 'red'}},
        ],
    })

    # Insert the chart into the worksheet.
    worksheet.insert_chart('E' + str(row), chart1)

    #
    # Create Average time Gains & Losses held in Days
    #
    catStr = '=Transactions!$H$' + str(title_row) + ':' + '$I$' + str(title_row)
    valStr = '=Transactions!$H$' + str(data_row) + ':' + '$I$' + str(data_row)

    chart2 = workbook.add_chart({'type': 'pie'})

    # Add a series to the chart.
    chart2.add_series({
        'name': 'Avg Days held for Gain & Loss',
        'categories': catStr,
        'values': valStr,
        'data_labels': {
            'value': True,
            'leader_lines': True
        },
        'points': [
            {'fill': {'color': 'green'}},
            {'fill': {'color': 'red'}},
        ],
    })

    # Insert the chart into the worksheet.
    worksheet.insert_chart('L' + str(row), chart2)


def build_IG_Deals_Header_row(worksheet,row,col,IG_Deals):

    write_row_of_text_to_excel(worksheet, row, IG_Trans_Date, ['Date'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Desc, ['Description'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Open_Price, ['Open Price'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Close_Price, ['Close Price'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Trans_Size, ['Size'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Total_Invested, ['Total Invested'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Profit_Loss, ['P&L'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Percent, ['%'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Trans_Open_Date, ['Open Date'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Days, ['Days'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Gains, ['Gains'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Loss, ['Loss'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Gain_Percentage, ['Gains %'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Loss_Percentage, ['Loss %'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Gains_Sterling, ['Gain £'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Loss_Sterling, ['Loss £'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Days_Gain, ['Days Gain'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Days_Loss, ['Days Loss'], 'BOLD', 'NO_UNDERLINE')





def build_IG_Deals_Summary_row(worksheet,row,col,IG_Deals):

    # Title
    write_row_of_text_to_excel(worksheet, row,IG_Trans_Date,['TOTALS'], 'BOLD','NO_UNDERLINE')
    # Count no of deals
    worksheet.write_formula(row, IG_Desc, build_COUNT_formula('C', 1, row))
    # Max Gain £
    worksheet.write_formula(row, IG_Open_Price, build_MAX_formula('O', 1, row),currency_format)
    # Max Loss £
    worksheet.write_formula(row, IG_Close_Price, build_MIN_formula('P', 1, row),currency_format)
    # Max Gain %
    worksheet.write_formula(row, IG_Trans_Size, build_MAX_formula('M', 1, row),percentage_format)
    # Max Loss %
    worksheet.write_formula(row, IG_Total_Invested, build_MIN_formula('N', 1, row),percentage_format)
    # Total P&L
    worksheet.write_formula(row, IG_Profit_Loss, build_SUM_formula('G', 1, row), currency_format)
    # Win %
    worksheet.write_formula(row, IG_Percent, ('K' + str(row+1) + '/B' + str(row+1)), percentage_format)
    # Total Gains
    worksheet.write_formula(row, IG_Gains, build_SUM_formula('K', 1, row))
    # Total Losses
    worksheet.write_formula(row, IG_Loss, build_SUM_formula('L', 1, row))
    # Avg Total Gains
    worksheet.write_formula(row, IG_Gain_Percentage, (build_SUM_formula('M', 1, row) +'/K' +str(row+1)), percentage_format)
    # Avg Total Losses
    worksheet.write_formula(row, IG_Loss_Percentage, (build_SUM_formula('N', 1, row) + '/L' + str(row + 1)),percentage_format)
    # Avg Profit
    worksheet.write_formula(row, IG_Gains_Sterling, (build_SUM_formula('O', 1, row) + '/K' + str(row + 1)),currency_format)
    # Avg £ Loss
    worksheet.write_formula(row, IG_Loss_Sterling, (build_SUM_formula('P', 1, row) + '/L' + str(row + 1)),currency_format)
    # Avg Days Gain
    worksheet.write_formula(row, IG_Days_Gain, (build_SUM_formula('Q', 1, row) + '/K' + str(row + 1)),no_decimals_format)
    # Avg Daily Losses
    worksheet.write_formula(row, IG_Days_Loss, (build_SUM_formula('R', 1, row) + '/L' + str(row + 1)),no_decimals_format)


    # Summary Details
    row += 1
    write_row_of_text_to_excel(worksheet, row, IG_Desc, ['No. Trades'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Open_Price, ['Max Gain'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Close_Price, ['Max Loss'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Trans_Size, ['Max Gain'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Total_Invested, ['Max Loss'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Profit_Loss, ['P & L'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Percent, ['Win %'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Gains, ['Gains'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Loss, ['Losses'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Gain_Percentage, ['Avg Gain'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Loss_Percentage, ['Avg Loss'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Gains_Sterling, ['Avg Gain'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Loss_Sterling, ['Avg Loss'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Days_Gain, ['Avg D Gain'], 'BOLD', 'NO_UNDERLINE')
    write_row_of_text_to_excel(worksheet, row, IG_Days_Loss, ['Avg D Loss'], 'BOLD', 'NO_UNDERLINE')




def build_IG_excel_deals(worksheet,row,col,IG_Deals):

    IG_Deals.reverse()
    build_IG_Deals_Header_row(worksheet, row, col, IG_Deals)
    row +=1

    #Blank row
#    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD','NO_UNDERLINE')

    for line in IG_Deals:

        # Date
        worksheet.write(row,col,line[IG_Date],date_format)
        col+=1

        # Desc
        worksheet.write(row, col, line[IG_MarketName])
        col += 1

        # Open Price
        worksheet.write(row, col, float(line[IG_Open]))
        col += 1

        # ClosePrice
        worksheet.write(row, col, float(line[IG_Close]))
        col += 1

        # Size
        worksheet.write(row, col, float(line[IG_Size]))
        col += 1

        # Total Invested
        TotalInvested =  (float(line[IG_Open])*float(line[IG_Size]))
        if(TotalInvested<0):
            TotalInvested *= -1
        worksheet.write(row, col, TotalInvested,currency_format)
        col += 1

        # P&L
        worksheet.write(row, col, float(re.sub(',','',line[IG_PL_Amount])),currency_format)
        col += 1

        # %
        if(row!=0):
            jump_row = row + 1
            formula = 'G' + str(jump_row) + '/' + 'F' + str(jump_row)
            worksheet.write_formula(row,col,formula,percentage_format)
        col += 1

        # Open Date
        worksheet.write(row, col, format_date(line[IG_Open_Date]),date_format)
        # Close Date
        #col += 1
        #worksheet.write(row, col, line[IG_Close_Date])
        col += 1

        # Calculate number of Days trade open
        if(row!=0):
            jump_row = row + 1
            worksheet.write_formula(row, col, str('A' + str(jump_row) + '-' + 'I' + str(jump_row)),)
        col += 1

        # Calculate if row a gain
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((G' + str(jump_row) + ')>0,1,0)'
            worksheet.write_formula(row, col, formula)
        col += 1

        # Calculate if row a loss
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((G' + str(jump_row) + ')<0,1,0)'
            worksheet.write_formula(row, col, formula)

        col += 1
        # Calculate GAIN %
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((H' + str(jump_row) + ')>0,' + 'H' + str(jump_row) + ',0)'
            worksheet.write_formula(row, col, formula,percentage_format)

        col += 1
        # Calculate LOSS %
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((H' + str(jump_row) + ')<0,' + 'H' + str(jump_row) + ',0)'
            worksheet.write_formula(row, col, formula, percentage_format)

        col +=1
        # Calculate Gains £
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((G' + str(jump_row) + ')>0,' + 'G' + str(jump_row) + ',0)'
            worksheet.write_formula(row, col, formula, currency_format)

        col += 1
        # Calculate Loss £
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((G' + str(jump_row) + ')<0,' + 'G' + str(jump_row) + ',0)'
            worksheet.write_formula(row, col, formula, currency_format)

        col += 1
        # Calculate Days of Gains
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((G' + str(jump_row) + ')>0,' + 'J' + str(jump_row) + ',0)'
            worksheet.write_formula(row, col, formula)

        col += 1
        # Calculate Days of Loss
        if (row != 0):
            jump_row = row + 1
            formula = 'IF((G' + str(jump_row) + ')<0,' + 'J' + str(jump_row) + ',0)'
            worksheet.write_formula(row, col, formula)

        row+=1
        col=0

    #
    # Build Summary Row
    #
    row += 1
    CheckSum = row+1
    build_IG_Deals_Summary_row(worksheet, row, col, IG_Deals)

    row += 3
    build_IG_Deals_Summary_row2(worksheet, row, col, IG_Deals,CheckSum)



def process_IG_csv_row(row):

    if(row[IG_Trans] == 'DEAL'):
        IG_Deals.append(row)
        #print(row)
    elif(row[IG_Trans] == 'DEPO' or row[IG_Trans] == 'WITH' or row[IG_Trans] == 'DIVIDEND'):
        IG_Costs.append(row)
        #print(row)
    else:
        print('Unrecognised Transaction Type:', row[IG_Trans])

#
########################## NW FUNCTIONS #######################################
#



def write_NW_trans_to_excel(worksheet,row, col, type, trans):

    SubTotal = int(0)
    start_row = str(row)
    for date, transaction, desc, pout, pin, balance in (trans):
        #print("Transaction",date,trans,desc,pout,pin,balance)
        col = 0
        worksheet.write(row, col, date)
     #  col+=1
     #  worksheet.write(row, col, transaction)
        col+=1
        worksheet.write(row, col, desc)
        col+=1
        if(type == 'INCOMING'):
            worksheet.write(row, col, float(re.sub('£','',pin)), currency_format)
            SubTotal += float(re.sub('£','',pin))
        else:
            worksheet.write(row, col, float(re.sub('£', '', pout)), currency_format)
            SubTotal += float(re.sub('£', '', pout))
        row += 1

    end_row = str(row)

    # Calculate Subtotal

    formula = 'SUM(C' + start_row + ':C' + end_row + ')'
    worksheet.write_formula(row, col+1, formula, currency_format)

    # Keep a track of all subtotals

    if(type == 'INCOMING'):
        GrandTotalIncoming.append(SubTotal)
        #print('GrandTotalIncoming', GrandTotalIncoming)
    else:
        GrandTotalOutgoing.append(SubTotal)
        #print('GrandTotalOutgoing', GrandTotalOutgoing)

    return(row)

def write_formula_to_excel(worksheet,row,col,formula,format):

    worksheet.write_formula(row, col, formula, format)
    row+=1
    return(row)

def build_NW_excel(worksheet):
    # Add a bold format to use to highlight cells.

    row = 0
    col = 0
    row = write_row_of_text_to_excel(worksheet, row, col, ['Date', 'Description', 'Paid In', 'Sub Totals','Grand Totals'], 'NO_BOLD', 'UNDERLINE')
    #Blank row
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD','NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet, row, col, ['Income'], 'BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD','NO_UNDERLINE')
    start_row = row
    row  = write_NW_trans_to_excel(worksheet, row, col, 'INCOMING', NW_Incoming_Trans)
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet, row, col, ['Transfer Payments'],'BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet, row, col, 'INCOMING', NW_Incoming_Transfers)
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD','NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')

    #
    # Write Total Income  formula
    #
    write_row_of_text_to_excel(worksheet,row, 1, ['Total Incoming'], 'BOLD', 'NO_UNDERLINE')
    formula = 'SUM(D' + str(start_row) + ':D' + str(row) + ')'
    row = write_formula_to_excel(worksheet, row, 4, formula, currency_format)

    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Outgoing Payments'], 'BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Date', 'Description', 'Paid Out', 'Sub Totals','Grand Totals'], 'NO_BOLD', 'UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Bills'], 'BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    start_row = row
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Bills)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Groceries'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Groceries)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Household'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Household)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['General'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_General)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Food & Drink'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Food_Drink)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Personal Care'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Personal_Care)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Experiences'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Experiences)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Shopping'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Shopping)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Transport'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Transport)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Other'], 'BOLD', 'NO_UNDERLINE')
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_Other)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')

    #
    # Write Total Outgoings
    #

    write_row_of_text_to_excel(worksheet,row, 1, ['Total Monthly Outgoings'], 'BOLD', 'NO_UNDERLINE')
    formula = 'SUM(D' + str(start_row) + ':D' + str(row) + ')'
    row = write_formula_to_excel(worksheet,row, 4, formula, currency_format)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')

    # Write Outgoing Transfers

    row = write_row_of_text_to_excel(worksheet,row, col, ['Transfers'], 'BOLD', 'NO_UNDERLINE')
    start_row = row
    row = write_NW_trans_to_excel(worksheet,row, col, 'OUTGOING', NW_OG_Transfers)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')

    #
    # Write Total Outgoing Transfers
    #
    write_row_of_text_to_excel(worksheet,row, 1, ['Total Outgoing Transfers'], 'BOLD', 'NO_UNDERLINE')
    formula = 'SUM(D' + str(start_row) + ':D' + str(row) + ')'
    row = write_formula_to_excel(worksheet,row, 4, formula, currency_format)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')

    #
    # Build Checksum section
    #
    rtn = process_NW_CheckSum(CheckSumBalance)
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['CHECKSUM'], 'BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['NW Account Balance',rtn[0]], 'NOT_BOLD', 'NO_UNDERLINE')
    row = write_row_of_text_to_excel(worksheet,row, col, ['Checksum   Balance',rtn[1]], 'NOT_BOLD', 'NO_UNDERLINE')

def check_dictionary(row, dictionary):
    try:
        category = dictionary[row[Desc]]
        #print('Dictionary Match', category)
        return(category)
    except:
        #print('Dictionary Lookup Failed')
        return('NOT_MATCHED')

def process_NW_incoming(row):

    category=check_dictionary(row,ic_types_dict)

    if(category == 'IC_TRANSFER'):
        #print('Transfer',row)
        NW_Incoming_Transfers.append(row)
    else:
        #print('I/C',row)
        NW_Incoming_Trans.append(row)
    #print("process_NW_incoming NW_Incoming_Trans",NW_Incoming_Trans)

def process_NW_outgoing(row):
    category = check_dictionary(row, og_types_dict)

    if(category == 'BILLS'):
        NW_Bills.append(row)
    elif(category == 'GROCERIES'):
        NW_Groceries.append(row)
    elif(category == 'HOUSEHOLD'):
        NW_Household.append(row)
    elif(category == 'GENERAL'):
        NW_General.append(row)
    elif(category == 'FOOD_DRINK'):
        NW_Food_Drink.append(row)
    elif(category == 'PERSONAL_CARE'):
        NW_Personal_Care.append(row)
    elif(category == 'EXPERIENCES'):
        NW_Experiences.append(row)
    elif(category == 'SHOPPING'):
        NW_Shopping.append(row)
    elif(category == 'TRANSPORT'):
        NW_Transport.append(row)
    elif(category == 'TRANSFER'):
        NW_OG_Transfers.append(row)
    else:
        print("OTHER: ",row)
        NW_Other.append(row)

def process_NW_csv_row(row):

    # Skip if a blank row
    if(len(row)>0):
        # print(row[0])
        #Check if Header row
        if(row[0] == 'Account Name:' or row[0] == 'Account Balance:' or row[0] == 'Available Balance: ' or row[0] == 'Date'):
            NW_Header.append(row)
            if(row[0] == 'Date'):
                return('ON')
        elif(len(row[Pin])>0):
            # Process Incoming
            process_NW_incoming(row)

        elif(len(row[Pout])>0):
            # Process Outgoing
            #print(row)
            process_NW_outgoing(row)

    return('OFF')
def process_NW_CheckSum(StartingBalance):

    TotalIncome = sum(GrandTotalIncoming)
    TotalOutgoing = sum(GrandTotalOutgoing)
    print('Starting Balance', StartingBalance)
    print('Total Income', TotalIncome)
    print('Total Outgoing', TotalOutgoing)

    # Retrieve the Account Balance from the Header

    for header in NW_Header:
        if(header[0] == 'Account Balance:'):
            AccountBalance = float(re.sub('£','',header[1]))
            print('Account Balance', AccountBalance)

    # Perform the CheckSum
    FinalBalance = round((StartingBalance + TotalIncome - TotalOutgoing),2)
    print('Final Balance', FinalBalance)

    if(FinalBalance == AccountBalance):
        print('********************* CHECKSUM OK   ***********************')
    else:
        print('********************* CHECKSUM FAIL ***********************')

    rtnTuple = (AccountBalance,FinalBalance)
    return (rtnTuple)


#####################################################################################
#####################################  MAIN #########################################
#####################################################################################

#def main(argv):

#
# Read in ARGV
#

opts, args = getopt.getopt(sys.argv[1:],"hi:o:t:",["ifile=","ofile=","ftype="])
for opt, arg in opts:
   if opt == '-h':
      print ('test.py -i <inputfile> -o <outputfile> -t <filetype>')
      sys.exit()
   elif opt in ("-i", "--ifile"):
      inputfile = arg
   elif opt in ("-o", "--ofile"):
      outputfile = arg
   elif opt in ("-t", "--ftype"):
      filetype = arg


"""
if __name__ == "__main__":
   main(sys.argv[1:])
"""

#
# Construct Directory Paths for input and output files
#
finalInputPath = os.path.join("c:",os.sep, "Users","sdinn","Downloads",inputfile)
finalOutputPath = os.path.join("c:",os.sep, "Users","sdinn","Downloads",outputfile)

finalIPHeaderPath = os.path.join("c:",os.sep, "Users","sdinn","OneDrive","Documents","Personal","Python","ParserFiles","Nationwide","IC_HEADER.csv")
finalOPHeaderPath = os.path.join("c:",os.sep, "Users","sdinn","OneDrive","Documents","Personal","Python","ParserFiles","Nationwide","OG_HEADER.csv")



print('FINAL Input PATH', finalInputPath)
print('FINAL Output PATH',finalOutputPath)
print('FINAL IC Header PATH', finalIPHeaderPath)
print('FINAL OG Header PATH', finalOPHeaderPath)

#
# If Nationwide Open IC_HEADER.csv & OG_HEADER.csv & build ic_types_dict Dictionaries
#

if(filetype == 'NW'):

    with open(finalIPHeaderPath, 'r') as f:

        csv_reader = csv.reader(f)

        for line in csv_reader:
            #print('IC_HEADER_FILE:', line)
            ic_types_dict[line[0]] = line[1]

        # print('IC_Dictionary', ic_types_dict)

    #
    # Open OP_HEADER.csv & build ic_ty
    #

    with open(finalOPHeaderPath, 'r') as f:

        csv_reader = csv.reader(f)

        for line in csv_reader:
            #print('OG_HEADER_FILE:', line)
            og_types_dict[line[0]] = line[1]

       # print('OG_Dictionary', og_types_dict)



#
# Open CSV file & process
#
with open(finalInputPath, 'r') as f:
    CheckSumFlag = 'OFF'
    CheckSumBalance = 0

    csv_reader = csv.reader(f)
    for line in csv_reader:
        # process each line
        #print(line)
        if CheckSumFlag == 'ON':
           CheckSumBalance = float(re.sub('£', '', line[Balance]))
           #print('CheckSum Balance', CheckSumBalance)
           if(len(line[Pin])>0):
              # Adjust CheckSumBalance to reflect the Incoming payment
              PaidIn = float(re.sub('£','',line[Pin]))
              CheckSumBalance -= PaidIn
              print('Adjusted Checksum', CheckSumBalance)
           elif(len(line[Pout])>0):
              # Adjust CheckSumBalance to reflect the Outgoing Payment
              PaidOut = float(re.sub('£','',line[Pout]))
              CheckSumBalance += PaidOut
              print('Opening Balance', CheckSumBalance)
        CheckSumFlag = 'OFF'
        if filetype == 'NW':
           CheckSumFlag = process_NW_csv_row(line)
        elif filetype == 'IG':
            process_IG_csv_row(line)



########################################################################################
################################   BUILD EXCEL #########################################
########################################################################################

workbook = xlsxwriter.Workbook(finalOutputPath)

# By default worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet("Transactions")
if(filetype == 'IG'):
    worksheet1 = workbook.add_worksheet("Costs")

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': 1})

# Add a number format for cells with money.
currency_format = workbook.add_format({'num_format': '£#,##0.00'})

# Add border underline

underline = workbook.add_format()
underline.set_bottom(1)

# date Format

date_format_str = 'dd/mm/yy'
date_format = workbook.add_format({'num_format': date_format_str,'align': 'left'})

# % format

percentage_format = workbook.add_format({'num_format': '0.00%'})

# Remove decimal figures

no_decimals_format = workbook.add_format({'num_format': '0'})

if filetype == 'NW':
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 50)
    worksheet.set_column(2, 2, 12)
    worksheet.set_column(3, 2, 12)
    worksheet.set_column(4, 2, 12)
    build_NW_excel(worksheet)
elif filetype == 'IG':
    # Build OUTPUT.xlsx -> Transactions
    worksheet.set_column(IG_Trans_Date, 0, 11)
    worksheet.set_column(IG_Desc, 1, 40)
    worksheet.set_column(IG_Open_Price, 2, 10)
    worksheet.set_column(IG_Close_Price, 3, 10)
    worksheet.set_column(IG_Trans_Size, 4, 7)
    worksheet.set_column(IG_Total_Invested, 5, 12)
    worksheet.set_column(IG_Profit_Loss, 6, 10)
    worksheet.set_column(IG_Percent, 7, 10)
    worksheet.set_column(IG_Trans_Open_Date, 8, 12)
    worksheet.set_column(IG_Days, 9, 10)
    worksheet.set_column(IG_Gains, 10, 8)
    worksheet.set_column(IG_Loss, 11, 10)
    worksheet.set_column(IG_Gain_Percentage, 12, 8)
    worksheet.set_column(IG_Loss_Percentage, 13, 8)
    worksheet.set_column(IG_Gains_Sterling, 14, 8)
    worksheet.set_column(IG_Loss_Sterling, 15, 8)
    worksheet.set_column(IG_Days_Gain, 16, 8)
    worksheet.set_column(IG_Days_Loss, 17, 10)
    build_IG_excel_deals(worksheet,0,0,IG_Deals)

    # Build OUTPUT.xlsx -> Costs
    worksheet1.set_column(IG_Trans_Date, 0, 11)
    worksheet1.set_column(IG_Desc, 1, 50)
    worksheet1.set_column(IG_Costs_Trans, 2, 10)
    worksheet1.set_column(IG_Costs_Amount, 3, 12)
    build_IG_excel_costs(worksheet1,0,0,IG_Costs)


workbook.close()
