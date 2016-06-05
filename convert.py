import os, subprocess, glob
import sys



"""
# set path to folder containing xls files
#path = sys.path
#os.chdir(path)

# find all files with extension .xls
xls = glob.glob('*.xls')

# create output filenames with extension .csv
csvs = [x.replace('.xls','.csv') for x in xls]

# zip into a list of tuples
in_out = zip(xls,csvs)

# loop through each file, calling the in2csv utility from subprocess
for xl,csv in in_out:
#   out = open(csv,'w')
#   command = 'in2csv --format xls %s > %s' % (xl, csv)
#   proc = subprocess.Popen(command,stdout=out)
#   proc.wait()
#   out.close()
   with open(csv, 'w') as of:
      subprocess.check_call(["in2csv", "--format",  "xls", xl],
                            stdout = of)
"""
print("yo")
for name in glob.glob('*.xls'):
   # suppose the first file name is "data.xls"
   # we would want to call "in2csv data.xls > data.csv" if we were at the command line
   # here is the equivalent with subprocess.call:
   print('in2csv ' + name + ' > ' + name + '.csv')

print("a")
subprocess.call(["ls", "-l"])

print("b")
subprocess.check_call(["ls", "-l"])

print("c")
subprocess.check_call(["in2csv", "--format", "xls", ""], stdout = "" )
