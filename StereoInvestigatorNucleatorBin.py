# This Python program produces binned nucleator data from Stereo Investigator 
# outputs. You can investigate either area or volume, select any number of 
# markers, any bin sizes, and any inquiry length you want.
# The comment sections below cover Python setup, file preparation, and 
# variable sets.

##############################################################################
# Section 1: setting up your computer to run this program
# This program is written in Python 2.7. 

# Step 1: Install Python 2.7.3 from here: 
# http://www.python.org/download/releases/2.7.3/
# On the page, select the Windows MSI installer (x86 if you have 32-bit 
# Windows installed, x86-64 if you have 64-bit Windows installed.)
# I suggest using the default option, which will install Python to c:/Python27

# Step 2: Copy this program into the c:/Python27 directory
# You can also put it into another directory that is added to the correct 
# PATH.

##############################################################################
# Section 2: file preparation for this program

# Create input files:
# Stereo Investigator doesn't mark nucleator data by region, but you can sort 
# by modified. Select all of the nucleator data for a region in the "display 
# probe run list". For each of your markers, highlight the marker, select 
# "copy to clipboard." Paste the data into an Excel spreadsheet. Save the 
# spreadsheet a as tab delimited .txt file. Save the file using the naming 
# convention "casename region celltype.txt". If you have multiple runs to 
# track (e.g., for reliability) save as "casename run region celltype.txt"
# (Don't mess around with the headers, this program expects default Stereo 
# Investigator output)

# Optional:
# This program can return either raw binned data, or data corrected by 
# estimated cell number.
# To correct data by cell number, create an excel file with the format:
# Column 1: Case
# Column 2: Run number (or leave blank)
# Column 3: Region name
# Column 4: Cell marker name
# Column 5: Number of cells
# Save as a tab-delimited text file.

##############################################################################
# Section 3: set your parameters for this program

# The directory where your saved .txt files are stored
directory = r"C:\Documents and Settings\Administrator\Desktop\Nucleator Outputs"

# The names of each case you want to look at
caselist = ["Case B", "Case X"]

# Run name, you can set this to "" if you have a single run per case
runs = ["Nuc 2"]

# Regions you're examining
regions = ["Acc Basal", "Amyg Other", "Basal", "Central", "Lateral"]

# Glial types you're counting, as filenames
celltypes = ["Astrocyte", "Endothelial", "Oligodendrocyte"]

# Marker names in file
markertypes = ["AstroMicro", "Endothelial", "Marker 2"]

# Set whether you want to bin "Area" or "Volume"
datatype = "Volume"

# Binning parameters here
# Set the binsize:
binsize = 10

# Set the starting distance:
binmin = 0

# Set the maximum distance:
binmax = 2000

# Set your output filename
output_file = r"Format test compare CSV 3.txt"

# If you want to correct by cell number, put the cell number file name here.
# The file should be in the same directory as your raw nucleator files and 
# saved as a text file. If you *don't* want to correct by stereologically 
# estimated cell number, input ""
number_file = "Nissl pilot raw numbers"

##############################################################################
# Section 4: save and run the program!
# The output file will be saved to the same directory as your input data.

##############################################################################
# program begins here

# Import modules to handle tab-delimited text file input and output
import csv


def nucleatordata():
    """This opens each nucleator file and creates a list, nucleatordata,
    with all nucleator data"""
    nucleatordata = []
    linetracker = 1
    for case in caselist:
        for run in runs:
            for region in regions:
                for celltype in celltypes:
                    if run == "":
                        nextfile = (case + " " + region + " " + celltype +
                                    ".txt")
                    else:
                        nextfile = (case + " " + run + " " + region + " " + 
                                    celltype + ".txt")
                    path = directory + "\\" + nextfile
                    myfileobj = open(path, "r") 
                    csv_read = csv.reader(myfileobj, dialect=csv.excel_tab)
                    newinput = []
                    for line in csv_read:
                        newinput.append(line[0:8])
                    newinput = newinput[1:-6]
                    linenumber = 0
                    for line in newinput:
                        addline = [case, run, region, newinput[linenumber][0],
                                   newinput[linenumber][1], 
                                   newinput[linenumber][2],
                                   newinput[linenumber][3]] 
                        nucleatordata.append(addline)
                        linenumber += 1
                        linetracker += 1
    return nucleatordata


def numberdata():
    """This opens the number file and gets it loaded into list numberdata if 
    you want to correct for raw numbers"""
    if len(number_file) > 0:
        numberdata = []
        path=directory + '/' + number_file + ".txt"
        myfileobj = open(path, "r") 
        csv_read = csv.reader(myfileobj, dialect=csv.excel_tab)
        for line in csv_read:
            numberdata.append(line[0:5])
        numberdata = numberdata[1:]
        return numberdata

                 
def bins(nucleatordata, case, celltype, region, binsize, binmin, binmax):
    """the binning algorithm for an individual case and celltype"""
    binlist = [0] * ((binmax - binmin) / binsize)
    # fill bins with data
    for cell in nucleatordata:
        if cell[0] == case:
            if cell[3] == celltype:
                if cell[2] == region:
                    try:
                        binselect = int(float(cell[datatype]) / 
                                        binsize - binmin)
                        binlist[binselect] = binlist[binselect] + 1
                    except:
                        pass
    return binlist


def bins_nocelltype(nucleatordata, case, region, binsize, binmin, binmax):
    binlist = [0] * ((binmax - binmin) / binsize)
    for cell in nucleatordata:
        if cell[0] == case:
            if cell[2] == region:
                try:
                    binselect = int(float(cell[datatype])/binsize - binmin)
                    binlist[binselect] = binlist[binselect]+1
                except:
                    pass
    return binlist


def bintotal(binoutput):
    """calculate the total number of nucleator marked cells for a cell type"""
    nuccellnumber = 0
    binnumber = len(binoutput)
    bintracker = 0
    while bintracker < binnumber:
        nuccellnumber = nuccellnumber + binoutput[bintracker]
        bintracker = bintracker + 1
    return nuccellnumber


def numcorrect(case, celltype, region, binoutput, nuccellnumber, numberdata):
    """divide number of calculated cells by number of nucleator cells to get 
    correction"""
    numcorrect = 0
    for row in numberdata:
        if row[0] == case:
            if row[3] == celltype:
                if row[2] == region:
                    numcorrect = float(row[4]) / float(nuccellnumber)
    return numcorrect


def numcorrect_nocelltype(case, region, binoutput, nuccellnumber, numberdata):
    """the same numcorrect except that all cell types are merged together"""
    numcorrect = 0
    for row in numberdata:
        if row[0] == case:
            if row[2] == region:
                if row[3] == "All":
                    numcorrect = float(row[4]) / float(nuccellnumber)
    return numcorrect


def bincorrect(binoutput, correction):
    """apply the number correction factor by bin"""
    bintrack = 0
    while bintrack < len(binoutput):
        binoutput[bintrack] = int(binoutput[bintrack] * correction)
        bintrack += 1
    return binoutput


def amygregionsum(case, celltype, binoutputcorrect, wholeamyg):
    """summation program for whole amygdala"""
    bintrack = 0
    while bintrack < len(binoutput):
        wholeamyg[bintrack] = wholeamyg[bintrack] + binoutputcorrect[bintrack]
        bintrack += 1
    return wholeamyg


def wholecasesum(case, wholeamyg, wholecase):
    bintrack = 0
    while bintrack < len(binoutput):
        wholecase[bintrack] = wholecase[bintrack] + wholeamyg[bintrack]
        bintrack += 1
    return wholecase


# prep file to write data into
out_path = directory + "\\" + output_file
output_writer = csv.writer(open(out_path, 'w'), delimiter='\t', quotechar='|',
                           quoting=csv.QUOTE_MINIMAL)

# convert data type into values
if datatype == "Area":
    datatype = 4
if datatype == "Volume":
    datatype = 5

# generate list headers,first row, bin is "x and smaller"
bintrack = binsize
header_row = ["Case", "Cell Type", "Region"]
while bintrack < binmax:
    header_row.append(str(bintrack) + " um")
    bintrack = bintrack + binsize
output_writer.writerow(header_row)

# fill in the sheet with each output, starting in the second row for headers
analysistrack = 1
nucleatordata = nucleatordata()
numberdata = numberdata()
for case in caselist:
    wholecase = [0] * ((binmax - binmin) / binsize)
    for celltype in markertypes:
        wholeamyg = [0] * ((binmax - binmin) / binsize)
        for region in regions:
            region_row = [case, celltype, region]        
            binoutput = bins(nucleatordata, case, celltype, region, binsize, 
                             binmin, binmax)
            nuccellnumber = bintotal(binoutput)
            if len(number_file) > 0:
                correction = numcorrect(case, celltype, region, binoutput, 
                                        nuccellnumber, numberdata)
                binoutputcorrect = bincorrect(binoutput,correction)      
            else:
                binoutputcorrect = binoutput
            wholeamyg = amygregionsum(case, celltype, binoutputcorrect, 
                                      wholeamyg)
            bintrack = 0
            while bintrack < (binmax - binsize):
                region_row.append(binoutputcorrect[bintrack / binsize])
                bintrack = bintrack + binsize                
            output_writer.writerow(region_row)
        #  tack on the whole amygdala summation
        wholeamy_row = [case, celltype, "Whole"]
        bintrack = 0
        while bintrack < (binmax - binsize):
            wholeamy_row.append(wholeamyg[bintrack / binsize])
            bintrack = bintrack + binsize
        wholecase = wholecasesum(case, wholeamyg, wholecase)
        output_writer.writerow(wholeamy_row)
    # tack on a whole case summation
    all_whole_row = [case, "All", "Whole"]
    bintrack = 0
    while bintrack < (binmax - binsize):
        all_whole_row.append(wholecase[bintrack / binsize])
        bintrack = bintrack + binsize
    output_writer.writerow(all_whole_row) 

# do an all cell analysis by region
for case in caselist:
    for region in regions:
        binoutput = bins_nocelltype(nucleatordata, case, region, binsize, 
                                    binmin, binmax)
        nuccellnumber = bintotal(binoutput)
        if len(number_file) > 0:
            correction = numcorrect_nocelltype(case, region, binoutput, 
                                               nuccellnumber, numberdata)
            binoutputcorrect = bincorrect(binoutput, correction)
        else:
            binoutputcorrect = binoutput
        all_reg_row = [case, "All", region]
        bintrack = 0
        while bintrack < (binmax - binsize):
            all_reg_row.append(binoutputcorrect[bintrack / binsize])
            bintrack = bintrack + binsize                
        output_writer.writerow(all_reg_row) 

