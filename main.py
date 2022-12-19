import csv
import xlsxwriter
import sys, getopt
import os.path

inputfile = ''
outputfile = ''
filetype = ''

#
# Start of main program
#

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

print ('inputfile is ', inputfile)
print ('outputfile is ', outputfile)
print('Type is', filetype)


"""
if __name__ == "__main__":
   main(sys.argv[1:])
"""

#
# Construct Directory Paths for input and output files
#
finalInputPath = os.path.join("c:",os.sep, "Users","sdinn","OneDrive",inputfile)
finalOutputPath = os.path.join("c:",os.sep, "Users","sdinn","Downloads",outputfile)

print('FINAL Input PATH =', finalInputPath)
print('FINAL Output PATH',finalOutputPath)


#f = open(finalInputPath,'r')

#f.close()




#with open(r'c:\Users\sdinn\Downloads\MoneyDB181222.csv', 'r') as f:
with open(finalInputPath, 'r') as f:

    csv_reader = csv.reader(f)
    for line in csv_reader:
        # process each line
        print(line)

print(line[0])


#workbook = xlsxwriter.Workbook(r'c:\Users\sdinn\Downloads\Example3.xlsx')

workbook = xlsxwriter.Workbook(finalOutputPath)

# By default worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet("My sheet")
# Some data we want to write to the worksheet.
scores = (
   ['blah', 1000],
   ['rahul', 100],
   ['priya', 300],
   ['harshita', 50],
)

# Start from the first cell. Rows and
# columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for name, score in (scores):
   worksheet.write(row, col, name)
   worksheet.write(row, col + 1, score)
   row += 1

workbook.close()