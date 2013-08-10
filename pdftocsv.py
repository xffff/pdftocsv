# file:   pdftocsv.py
# desc:   example of how to grab all data from a 
#         number of PDF forms using Python and VBS
#         and output to something readable (CSV)
# author: Michael Murphy 2013

import os, csv, fileinput, codecs, re, sys
from bs4 import BeautifulSoup

fileDicts = []

def runVBS(d):
    ''' for some reason I can't use AcroJS from Python, so use VBS '''
    print "Converting all PDF files to XML..."
    os.system("pdftoxml.vbs " + d + "\\")
    print "Done converting..."

def parseFiles(d):
    for root, dirs, files in os.walk(d):
        fn = 0
        for file in files:
            if file.endswith(".xml"):
                print 'Parsing: ', os.path.join(root, file)
                fileDicts.append(dict())
                for line in open(os.path.join(d,file), "r"):                    
                    if "<TD>" in line:
                        data = BeautifulSoup(line).get_text().strip().encode('ascii','ignore')
                        fileDicts[fn][keys[count]] = data                        
                fn += 1

def writeOut(d):
    masterDict = {k: [] for k in keys}

    for f in fileDicts:
        for k, v in f.items():
            if k in masterDict.keys():
                masterDict[k].append(v)
    try:
        print "Writing %s offlinefiles.csv" % d
        writer = csv.writer(open(os.path.join(d,'pdfoutput.csv'),'wb'))
        writer.writerow([k for k in masterDict.keys()])
        for i in range(len(fileDicts)):
            writer.writerow([v[i] for k, v in masterDict.items() if i < len(v)])
    except IOError:
        print 'ERROR: cannot open file to write CSV'
        return 1

def main(arg):
    if len(arg) < 2:
        return 'Usage: %s path-to-folder' % arg[0]
    if not os.path.exists(arg[1]):
        return 'ERROR: Folder %s was not found' % arg[1]

    d  = os.path.normpath(arg[1])

    if not os.listdir(directory): 
        return "%s is empty, giving up" % d
    else: 
        parseFiles(d)
        writeOut(d)

if __name__ == "__main__":
    sys.exit(main(sys.argv))
