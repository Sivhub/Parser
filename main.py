import csv
import xlsxwriter
import sys, getopt
import os.path

#
################################ Global App Data ###############################
#
PARSED = int(0)
NOT_PARSED = int(-1)
#NOT_MATCHED = str('MATCH')
# Excel worksheet row


inputfile = ''
outputfile = ''
filetype = ''
row= int(0)
#
################################ NW Data ####################################
#

# Transaction Fields
Date = int(0)
Trans = int(1)
Desc = int(2)
Pout = int(3)
Pin = int(4)
Balance = int(5)

# NW Incoming Transactions
NW_Incoming_Trans = []
NW_Incoming_Transfers = []

IC_TRANSFER = int(0)

ic_types_dict = {
    '070040 18408110':'IC_TRANSFER',
    'Bank credit SCO GBP CMT AC':'IC_TRANSFER'
}

#
############################ IG FUNCTIONS #####################################
#

def build_IG_excel():
    pass
def process_IG_csv_row(row):
    print(row)

#
########################## NW FUNCTIONS #######################################
#

def write_row_of_text_to_excel(row, col, text):
    for item in text:
        #print("Text: ",item)
        worksheet.write(row, col, item, bold)
        col+=1
    row+=1
    return(row)
def write_NW_trans_to_excel(row, col, trans):

    for date, trans, desc, pout, pin, balance in (trans):
        #print("Transaction",date,trans,desc,pout,pin,balance)
        col = 0
        worksheet.write(row, col, date)
     #  col+=1
     #  worksheet.write(row, col, trans)
        col+=1
        worksheet.write(row, col, desc)
        col+=1
        worksheet.write(row, col, pin)
        row += 1
    return(row)
def build_NW_excel():
    # Add a bold format to use to highlight cells.

    row = 0
    col = 0
    row = write_row_of_text_to_excel(row, col,['Incoming Payments'])
    row = write_row_of_text_to_excel(row,col,['Date','Description','Paid In'])
    row  = write_NW_trans_to_excel(row, col, NW_Incoming_Trans)
    row = write_row_of_text_to_excel(row, col, ['Transfer Payments'])
    row = write_NW_trans_to_excel(row, col, NW_Incoming_Transfers)

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
        print('Transfer',row)
        NW_Incoming_Transfers.append(row)
    else:
        print('I/C',row)
        NW_Incoming_Trans.append(row)
    #print("process_NW_incoming NW_Incoming_Trans",NW_Incoming_Trans)

def process_NW_outgoing(row):
    pass
def process_NW_csv_row(row):

    #
    # Eliminate the header
    #
    if(len(row)>0):
       # print(row[0])
        if(row[0] == 'Account Name:' or row[0] == 'Account Balance:' or row[0] == 'Available Balance: ' or row[0] == 'Date'):
            return(NOT_PARSED)

        if(len(row[Pin])>0):
            # Process Incoming
            process_NW_incoming(row)

        if(len(row[Pout])>0):
            # Process Outgoing
            #print(row)
            process_NW_outgoing(row)



    #if(row == 'Account Name:'):
        #print(row)

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

#print ('inputfile is ', inputfile)
#print ('outputfile is ', outputfile)
#print('Type is', filetype)


"""
if __name__ == "__main__":
   main(sys.argv[1:])
"""

#
# Construct Directory Paths for input and output files
#
finalInputPath = os.path.join("c:",os.sep, "Users","sdinn","Downloads",inputfile)
finalOutputPath = os.path.join("c:",os.sep, "Users","sdinn","Downloads",outputfile)

print('FINAL Input PATH =', finalInputPath)
print('FINAL Output PATH',finalOutputPath)

#
# Open CSV file & process
#
with open(finalInputPath, 'r') as f:

    csv_reader = csv.reader(f)
    for line in csv_reader:
        # process each line
        #print(line)
        if filetype == 'NW':
           if(process_NW_csv_row(line) == NOT_PARSED):
               print('Not Parsed:', line)
        elif filetype == 'IG':
            process_IG_csv_row(line)


########################################################################################
################################   BUILD EXCEL #########################################
########################################################################################

workbook = xlsxwriter.Workbook(finalOutputPath)

# By default worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet("My sheet")

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': 1})

if filetype == 'NW':
    build_NW_excel()
elif filetype == 'IG':
    build_IG_excel()


workbook.close()