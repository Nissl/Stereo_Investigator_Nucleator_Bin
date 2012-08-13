##This Python program produces binned nucleator data from Stereo Investigator outputs.
##You can investigate either area or volume,
##select any number of markers, any bin sizes, and any inquiry length you want.
##The output is in .xls format. 
##The comment sections below cover Python setup, file preparation, and variable sets.

###################################################################
##Section 1: setting up your computer to run this program
##This program is written in Python 2.7. It also uses the xlwt addon library to make Excel spreadsheets.

##Step 1: Install Python 2.7.3 from here: http://www.python.org/download/releases/2.7.3/
##On the page, select the Windows MSI installer (x86 if you have 32-bit Windows installed,
##x86-64 if you have 64-bit Windows installed.)
##I suggest using the default option, which will install Python to c:/Python27

##Step 2: Install the xlwt library from here: http://pypi.python.org/pypi/xlwt/
##Use the program WinRAR to unzip the files to a directory
##Go to "run" in the start menu and type cmd
##Type cd c:\directory_where_xlwt_was_unzipped_to
##Type setup.py install

##Step 3: Copy this program into the c:/Python27 directory
##You can also put it into another directory that is added to the correct PATH.

###################################################################
##Section 2: file preparation for this program

##Create input files
##Stereo Investigator doesn't mark nucleator data by region, but you can sort by modified 
##Select all of the nucleator data for a region in the "display probe run list"
##For each of your markers, highlight the marker, select "copy to clipboard"
##Paste the data into an Excel spreadsheet. 
##Save the spreadsheet a as tab delimited .txt file
##Save the file using the naming convention "casename region celltype.txt"
##If you have multiple runs to track (e.g., for reliability) save as "casename run region celltype.txt"
##(Don't mess around with the headers, this program expects default Stereo Investigator output)

##Optional:
##This program can return either raw binned data, or data corrected by estimated cell number
##To correct data by cell number, create an excel file with the format:
##Column 1: Case
##Column 2: Run number (or leave blank)
##Column 3: Region name
##Column 4: Cell marker name
##Column 5: Number of cells

###################################################################
##Section 3: set your parameters for this program

##The directory where your saved .txt files are stored
directory = "C:\Documents and Settings\Administrator\Desktop\Nucleator Outputs"

#The names of each case you want to look at
caselist = ["Case B", "Case X"]

#Run name, you can set this to "" if you have a single run per case
runs = ["Nuc 2"]

#Regions you're examining
regions = ["Acc Basal", "Amyg Other", "Basal", "Central", "Lateral"]

#Glial types you're counting, as filenames
celltypes = ["Astrocyte", "Endothelial", "Oligodendrocyte"]

#Marker names in file
markertypes = ["AstroMicro", "Endothelial", "Marker 2"]

#Set whether you want to bin "Area" or "Volume"
datatype = "Volume"

#Binning parameters here
#Set the binsize:
binsize = 10
#Set the start point:
binmin = 0
#Set the max point
binmax = 2000

#Set your output filename
outputfile = "Refactor test compare.xls"

#If you want to correct by cell number, put the cell number file name here.
#The file should be in the same directory as your raw nucleator files and saved as a text file
#If you *don't* want to correct by stereologically estimated cell number, input ""
numberfile = "Nissl pilot raw numbers"

###################################################################
##Section 4: save and run the program!
##The output file will be saved to the same directory as your input data

###################################################################
##Program begins here

#Import modules to handle a tab-delimited text file and produce .xls output
import csv
import xlwt

#Program-specific modules here

##This opens each nucleator file and creates a list, nucleatordata,
##with all nucleator data
def nucleatordata():
    nucleatordata = []
    linetracker = 1
    for case in caselist:
        for run in runs:
            for region in regions:
                for celltype in celltypes:
                    if run == "":
                        nextfile = case + " " + region + " " + celltype + ".txt"
                    else:
                        nextfile = case + " " + run + " " + region + " " + celltype + ".txt"
                    path = directory + '/' + nextfile
                    myfileobj = open(path,"r") 
                    csv_read = csv.reader(myfileobj,dialect=csv.excel_tab)
                    newinput = []
                    for line in csv_read:
                        newinput.append(line[0:8])
                    newinput = newinput[1:-6]
                    linenumber = 0
                    for line in newinput:
                        addline = [case,run,region,newinput[linenumber][0],newinput[linenumber][1],newinput[linenumber][2],
                                newinput[linenumber][3]] 
                        nucleatordata.append(addline)
                        linenumber += 1
                        linetracker += 1
    return nucleatordata

#This opens the number file and gets it loaded into list numberdata if you want to correct 
#for raw numbers
def numberdata():
    if len(numberfile) > 0:
        numberdata = []
        path=directory + '/' + numberfile + ".txt"
        myfileobj = open(path,"r") 
        csv_read = csv.reader(myfileobj,dialect=csv.excel_tab)
        for line in csv_read:
            numberdata.append(line[0:5])
        numberdata = numberdata[1:]
        return numberdata

##the binning algorithm for an individual case and celltype                    
def bins(nucleatordata,case,celltype,region,binsize,binmin,binmax):
    binlist = [0] * ((binmax - binmin) / binsize)

#fill bins with data
    for cell in nucleatordata:
        if cell[0] == case:
            if cell[3] == celltype:
                if cell[2] == region:
                    try:
                        binselect = int(float(cell[datatype]) / binsize - binmin)
                        binlist[binselect] = binlist[binselect] + 1
                    except:
                        pass
    return binlist

def bins_nocelltype(nucleatordata,case,region,binsize,binmin,binmax):
    binlist = [0] * ((binmax - binmin) / binsize)

#fill bins with data
    for cell in nucleatordata:
        if cell[0] == case:
            if cell[2] == region:
                try:
                    binselect = int(float(cell[datatype])/binsize - binmin)
                    binlist[binselect] = binlist[binselect]+1
                except:
                    pass
    return binlist

##calculate the total number of nucleator marked cells for a cell type
def bintotal(binoutput):
    nuccellnumber = 0
    binnumber = len(binoutput)
    bintracker = 0
    while bintracker < binnumber:
        nuccellnumber = nuccellnumber + binoutput[bintracker]
        bintracker = bintracker + 1
    return nuccellnumber

##divide number of calculated cells by number of nucleator cells to get correction
def numcorrect(case,celltype,region,binoutput,nuccellnumber,numberdata):
    numcorrect = 0
    for row in numberdata:
        if row[0] == case:
            if row[3] == celltype:
                if row[2] == region:
                    numcorrect = float(row[4]) / float(nuccellnumber)
    return numcorrect

##the same numcorrect except that all cell types are merged together 
def numcorrect_nocelltype(case,region,binoutput,nuccellnumber,numberdata):
    numcorrect = 0
    for row in numberdata:
        if row[0] == case:
            if row[2] == region:
                if row[3] == "All":
                    numcorrect = float(row[4]) / float(nuccellnumber)
    return numcorrect

##apply the number correction factor by bin
def bincorrect(binoutput,correction):
    bintrack = 0
    while bintrack < len(binoutput):
        binoutput[bintrack] = int(binoutput[bintrack] * correction)
        bintrack += 1
    return binoutput

##summation program for whole amygdala
def amygregionsum(case,celltype,binoutputcorrect,wholeamyg):
    bintrack = 0
    while bintrack < len(binoutput):
        wholeamyg[bintrack] = wholeamyg[bintrack] + binoutputcorrect[bintrack]
        bintrack += 1
    return wholeamyg

def wholecasesum(case,wholeamyg,wholecase):
    bintrack = 0
    while bintrack < len(binoutput):
        wholecase[bintrack] = wholecase[bintrack] + wholeamyg[bintrack]
        bintrack += 1
    return wholecase

##

#Excel sheet prep
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Python Sheet 1")

#convert data type into values
if datatype == "Area":
    datatype = 4
if datatype == "Volume":
    datatype = 5

#generate list headers,first row, bin is "x and smaller"
bintrack = binsize
while bintrack < binmax:
    sheet1.write(0,bintrack/binsize+2,bintrack)
    bintrack = bintrack + binsize
sheet1.write(0,0,"Case")
sheet1.write(0,1,"Cell Type")
sheet1.write(0,2,"Region")

#fill in the sheet with each output, starting in the second row for headers
analysistrack = 1
nucleatordata = nucleatordata()
numberdata = numberdata()
for case in caselist:
    wholecase = [0] * ((binmax-binmin)/binsize)
    for celltype in markertypes:
        wholeamyg = [0] * ((binmax-binmin)/binsize)
        for region in regions:
            sheet1.write(analysistrack,0,case)
            sheet1.write(analysistrack,1,celltype)
            sheet1.write(analysistrack,2,region)
            
            binoutput = bins(nucleatordata,case,celltype,region,binsize,binmin,binmax)
            nuccellnumber = bintotal(binoutput)
            if len(numberfile) > 0:
                correction = numcorrect(case,celltype,region,binoutput,nuccellnumber,numberdata)
                binoutputcorrect = bincorrect(binoutput,correction)      
            else:
                binoutputcorrect = binoutput
            wholeamyg = amygregionsum(case,celltype,binoutputcorrect,wholeamyg)
            bintrack = 0
            while bintrack < (binmax-binsize):
                sheet1.write(analysistrack,bintrack/binsize+3,binoutputcorrect[bintrack/binsize])
                bintrack = bintrack + binsize                
            analysistrack += 1
        #tack on the whole amygdala summation
        sheet1.write(analysistrack,0,case)
        sheet1.write(analysistrack,1,celltype)
        sheet1.write(analysistrack,2,"Whole")
        bintrack = 0
        while bintrack < (binmax - binsize):
            sheet1.write(analysistrack,bintrack/binsize+3,wholeamyg[bintrack/binsize])
            bintrack = bintrack + binsize
        wholecase = wholecasesum(case,wholeamyg,wholecase)
        analysistrack += 1
    #tack on a whole case summation
    sheet1.write(analysistrack,0,case)
    sheet1.write(analysistrack,1,"All")
    sheet1.write(analysistrack,2,"Whole")
    bintrack = 0
    while bintrack < (binmax - binsize):
        sheet1.write(analysistrack,bintrack/binsize+3,wholecase[bintrack/binsize])
        bintrack = bintrack + binsize 
    analysistrack += 1

#do a whole cell analysis by region
for case in caselist:
    for region in regions:
        binoutput = bins_nocelltype(nucleatordata,case,region,binsize,binmin,binmax)
        nuccellnumber = bintotal(binoutput)
        if len(numberfile) > 0:
            correction = numcorrect_nocelltype(case,region,binoutput,nuccellnumber,numberdata)
            binoutputcorrect = bincorrect(binoutput,correction)
        else:
            binoutputcorrect = binoutput
        sheet1.write(analysistrack,0,case)
        sheet1.write(analysistrack,1,"All")
        sheet1.write(analysistrack,2,region)
        bintrack = 0
        while bintrack < (binmax - binsize):
            sheet1.write(analysistrack,bintrack/binsize+3,binoutputcorrect[bintrack/binsize])
            bintrack = bintrack + binsize                
        analysistrack += 1
    
               
savepath = directory + "\\" + outputfile
book.save(savepath)

