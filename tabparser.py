# .xls files are not actually xls
# they appear to be tab-separated text files
import glob

for files in glob.glob("*.xls"):
    print(files)
    with open(files, 'r') as f:
        for line in f:
            print line
