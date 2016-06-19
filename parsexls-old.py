import glob
import csv
import re
#import pandas
from bs4 import BeautifulSoup
import statistics

def distance(allmean, allsd, onemean):
    return ((onemean - allmean)/allsd)

_nsre = re.compile('([0-9]+)')
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(_nsre, s)] 

allresponses = []
allcourses = set()
for files in glob.glob("*.xls"):
    coursename = files.rstrip(".xls")
    allcourses.add(coursename)
    with open(files, 'rb') as f:
        soup = BeautifulSoup(f, "html.parser")
        lines = str(soup).split('\n')
        for index, line in enumerate(lines):
            if index == 0:
                header = ["Coursename"]+line.split('\t')
                header = [x.replace('"','') for x in header]
            else:
                courseline = line.split('\t')
                courseline.insert(0,coursename)
                courseline = [x.replace('"','') for x in courseline]
                courseline = [re.sub('<(?:"[^"]*"[\'"]*|\'[^\']*\'[\'"]*|[^\'">])+>', '', x) for x in courseline]
                allresponses.append(courseline)

## write to csv with all data
g = open('allspring.csv', 'w')
gcsvwriter = csv.writer(g, delimiter=',', dialect='excel', lineterminator='\n')
gcsvwriter.writerow(header)
for response in allresponses:
    gcsvwriter.writerow(response)
g.close()

mydict = {0:'<unanswered>',
                1:'Strongly Disagree',
                2:'Disagree',
                3:'Neither Agree nor Disagree',
                4:'Agree',
                5:'Strongly Agree'}
answerindex = {} # key = index: ['Answer XX', {'course':[[values],mean,sd fromoverall mean}, overall mean, sd]
for index, item in enumerate(header):
    m = re.search(r'Answer (\d+)', item)
    if m != None:
        answerindex[index] = [m.group(0),{}]

for response in allresponses:
   if len(response) == len(header):
       for key in answerindex.keys():
           if response[key] in mydict.values():
               rating = list(mydict.keys())[list(mydict.values()).index(response[key])]
               #if response[0] in answerindex[key][1].keys(): #check if coursename has been added to dictionary
               answerindex[key][1].setdefault(response[0], [[]])
               answerindex[key][1][response[0]][0].append(rating)
                
for answer in answerindex.keys():
    allresponses = []
    for course in answerindex[answer][1].keys():
        c =(answerindex[answer][1][course][0])
        answerindex[answer][1][course].append(sum(c,0.0)/len(c))  
        allresponses.extend(c)
#    print(allresponses)
    if len(allresponses) > 0:
         answermean = sum(allresponses, 0.0)/(len(allresponses))
         answerindex[answer].append(answermean)
         answersd = statistics.stdev(allresponses)
         answerindex[answer].append(answersd)
         [answerindex[answer][1][course].append(distance(answermean, answersd, answerindex[answer][1][course][1]))
          for course in answerindex[answer][1].keys()]
#print(answerindex)
courses = list(allcourses)
courses.sort(key=natural_sort_key)

g = open('coursesByAnswers.csv', 'w')
gcsvwriter = csv.writer(g, delimiter=',', dialect='excel', lineterminator='\n')
courseslong = []
[courseslong.extend([x,'Course mean', 'SD from Overall Mean']) for x in courses]
print(courseslong)
answerHeader = ["Answer number","Overall Mean","Overall SD"]
answerHeader.extend(courseslong)
gcsvwriter.writerow(answerHeader)

for answer in answerindex.keys():
    answerlist = answerindex[answer]
    #answerout = [answerlist[0],'','']
    if len(answerlist) ==4:
        answerout = [answerlist[0],answerlist[2],answerlist[3]]
        for course in courses:
            coursemean = ''
            coursesd = ''
            if course in answerlist[1].keys():
                coursemean = answerlist[1][course][1]
                coursesd = answerlist[1][course][2]
            answerout.extend([course,coursemean,coursesd])
        gcsvwriter.writerow(answerout)
g.close()
            
                
