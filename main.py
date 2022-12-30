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

#
################################ IG Data ####################################
#

#
#Transaction Fields
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

# NW Outgong Transactions
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
def build_IG_excel_costs(worksheet):
    row = 0
    col = 0
    row = write_row_of_text_to_excel(worksheet, row, col,
                                     ['Date', 'Description', 'Buy Price', 'Size', 'Total Invested', 'P&L', 'Close Date',
                                      'Days'], 'NO_BOLD', 'UNDERLINE')
    # Blank row
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD', 'NO_UNDERLINE')


# row = write_row_of_text_to_excel(row, col,[IG_Deals], 'NO_BOLD', 'UNDERLINE')

def build_IG_excel_deals(worksheet):

    row = 0
    col = 0
    row = write_row_of_text_to_excel(worksheet, row, col, ['Date', 'Description', 'Buy Price', 'Size','Total Invested','P&L','Close Date','Days'], 'NO_BOLD', 'UNDERLINE')
    #Blank row
    row = write_row_of_text_to_excel(worksheet, row, col, '', 'NOT_BOLD','NO_UNDERLINE')


def process_IG_csv_row(row):

    if(row[IG_Trans] == 'DEAL'):
        IG_Deals.append(row)
        #print(row)
    else:
        IG_Costs.append(row)
        print(row)

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
            worksheet.write(row, col, eval(re.sub('£','',pin)), currency_format)
            SubTotal += eval(re.sub('£','',pin))
        else:
            worksheet.write(row, col, eval(re.sub('£', '', pout)), currency_format)
            SubTotal += eval(re.sub('£', '', pout))
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
            AccountBalance = eval(re.sub('£','',header[1]))
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

#worksheet.write(row, col, eval(re.sub('£','',pin)), currency_format)
 #           SubTotal += eval(re.sub('£','',pin))
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
           CheckSumBalance = eval(re.sub('£', '', line[Balance]))
           #print('CheckSum Balance', CheckSumBalance)
           if(len(line[Pin])>0):
              # Adjust CheckSumBalance to reflect the Incoming payment
              PaidIn = eval(re.sub('£','',line[Pin]))
              CheckSumBalance -= PaidIn
              print('Adjusted Checksum', CheckSumBalance)
           elif(len(line[Pout])>0):
              # Adjust CheckSumBalance to reflect the Outgoing Payment
              PaidOut = eval(re.sub('£','',line[Pout]))
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





if filetype == 'NW':
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 50)
    worksheet.set_column(2, 2, 12)
    worksheet.set_column(3, 2, 12)
    worksheet.set_column(4, 2, 12)
    build_NW_excel(worksheet)
elif filetype == 'IG':
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 50)
    worksheet.set_column(2, 2, 12)
    worksheet.set_column(3, 2, 12)
    worksheet.set_column(4, 2, 12)
    worksheet.set_column(5, 0, 20)
    worksheet.set_column(6, 1, 50)
    worksheet.set_column(7, 2, 12)
    worksheet.set_column(8, 2, 12)
    worksheet.set_column(9, 2, 12)
    build_IG_excel_deals(worksheet)
    build_IG_excel_costs(worksheet1)


workbook.close()