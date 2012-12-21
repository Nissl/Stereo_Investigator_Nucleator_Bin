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
# spreadsheet as a tab delimited .txt file. Leave one line of headers.

# Alternately, choose the "export to Excel" option, open the raw nucleator 
# data, and delete any rows with average calculations. Delete all tabs other 
# than the one containing the raw data. Save the spreadsheet as a tab delimited
# .txt file. Caution: the export function is extremely slow in current 
# versions of Stereo Investigator.

# Optional:
# This program can return either raw binned data, or data corrected by 
# estimated cell number.
# To correct data by cell number, create an excel file with the format:
# Column 1: Case: use the same names as in the case list.
# Column 2: Cell marker name
# Column 3: Number of cells
# Save as a tab-delimited text file.

##############################################################################
# Section 3: set your parameters for this program

# The directory where your saved .txt files are stored
directory = r"C:\Documents and Settings\Administrator\My Documents"

# The names of each case you want to look at
case_list = ["Nucleator Case A reliability 1 12-21-12", 
            "Nucleator Case A reliability 2 12-21-12",
            "Nucleator Case A reliability 3 12-21-12"]

# Multiple markers?
multiple_marks = True

# Set whether you want to bin "Area" or "Volume"
data_type = "Volume"

# Binning parameters here
# Set the bin_size:
bin_size = 10

# Set the starting distance:
bin_min = 0

# Set the maximum distance:
bin_max = 2000

# Set your output filename
output_file = r"Nucleator oligo-astro reliability bin10.txt"

# If you want to correct by cell number, put the cell number file name here.
# The file should be in the same directory as your raw nucleator files and 
# saved as a text file. If you *don't* want to correct by stereologically 
# estimated cell number, input ""
number_file = ""

##############################################################################
# Section 4: save and run the program!
# The output file will be saved to the same directory as your input data.

##############################################################################
# program begins here

# Import module to handle tab-delimited text file
import csv


def nucleator_data(case_list):
    """This opens each nucleator file and creates a list, nuc_data,
    with all nucleator data"""
    nuc_data = [["Run", "Cell Type", "Area (um^2)", "Volume (um^3)", 
                     "Length (um)"]]
    for case in case_list:
        nextfile = (case + ".txt")
        path = directory + "\\" + nextfile
        myfileobj = open(path, "r") 
        csv_read = csv.reader(myfileobj, dialect=csv.excel_tab)
        raw_data = []
        if not multiple_marks:
            for line in csv_read:
                raw_data.append(line[0:4])
            raw_data = raw_data[1:]
            linenumber = 0
            for _ in raw_data:
                addline = [case, raw_data[linenumber][0],
                           raw_data[linenumber][1], 
                           raw_data[linenumber][2],
                           raw_data[linenumber][3]] 
                nuc_data.append(addline)
                linenumber += 1
        if multiple_marks: 
            for line in csv_read:
                raw_data.append(line[0:7])
            raw_data = raw_data[1:]
            linenumber = 0
            for _ in raw_data:
                addline = [case, raw_data[linenumber][0],
                           raw_data[linenumber][1], 
                           raw_data[linenumber][2],
                           raw_data[linenumber][3], 
                           raw_data[linenumber][4],
                           raw_data[linenumber][5],
                           raw_data[linenumber][6]] 
                nuc_data.append(addline)
                linenumber += 1
    return nuc_data


def celltypes(nuc_data):
    cell_type_list = []
    for cell in nuc_data:
        if (cell[1] not in cell_type_list and cell[1] != "Cell Type" and 
            cell[1] != " "):
            cell_type_list.append(cell[1])
    return cell_type_list


def number_data():
    """This opens the number file and gets it loaded into list number_data if 
    you want to correct for raw numbers"""
    number_data = []
    path=directory + '/' + number_file + ".txt"
    myfileobj = open(path, "r") 
    csv_read = csv.reader(myfileobj, dialect=csv.excel_tab)
    for line in csv_read:
        number_data.append(line[0:5])
    number_data = number_data[1:]
    return number_data

                 
def bins(nuc_data, case, cell_type, data_type, bin_size, bin_min, bin_max):
    """the binning algorithm for an individual case and celltype"""
    bin_list = [0] * ((bin_max - bin_min) / bin_size)
    # fill bins with data
    for cell in nuc_data:
        if cell[0] == case:
            if cell[1] == cell_type:
                try:
                    binselect = int(float(cell[data_type]) / 
                                    bin_size - bin_min)
                    bin_list[binselect] = bin_list[binselect] + 1
                except:
                    pass
    return bin_list


def bintotal(bin_output):
    """Calculate the total number of nucleator marked cells for a cell type."""
    nuc_cell_num = 0
    bin_num = len(bin_output)
    bin_tracker = 0
    while bin_tracker < bin_num:
        nuc_cell_num = nuc_cell_num + bin_output[bin_tracker]
        bin_tracker += 1
    return nuc_cell_num


def num_correct(case, cell_type, bin_output, nuc_cell_num, number_data):
    """Divide the number of calculated cells by number of nucleator cells to 
    get a correction factor"""
    num_correct = 0
    for row in number_data:
        if row[0] == case and row[1] == cell_type:
            num_correct = float(row[2]) / float(nuc_cell_num)
    return num_correct


def bin_correct(bin_output, correction):
    """Apply the number correction factor by bin."""
    bin_track = 0
    while bin_track < len(bin_output):
        bin_output[bin_track] = int(bin_output[bin_track] * correction)
        bin_track += 1
    return bin_output


##############################################################################
# Convert data type into location value.
if multiple_marks:
    if data_type == "Area": data_type = 2
    if data_type == "Volume": data_type = 5
if not multiple_marks:
    if data_type == "Area": data_type = 2
    if data_type == "Volume": data_type = 3

# Create file to write data into.
out_path = directory + "\\" + output_file
output_writer = csv.writer(open(out_path, 'w'), delimiter='\t', quotechar='|',
                           quoting=csv.QUOTE_MINIMAL)

# Generate file headers and place them in the first row. Bins are "x and 
# smaller"
bin_track = bin_size
header_row = ["Case", "Cell Type"]
while bin_track <= bin_max:
    header_row.append(str(bin_track) + " um")
    bin_track = bin_track + bin_size
output_writer.writerow(header_row)

nuc_data = nucleator_data(case_list)
cell_type_list = celltypes(nuc_data)

if number_file:
    number_data = number_data()

# Fill in the sheet with each output, starting in the second row for headers. 
for case in case_list:
    cell_sum = [0] * ((bin_max - bin_min) / bin_size)
    for cell_type in cell_type_list:    
        bin_output = bins(nuc_data, case, cell_type, data_type, bin_size, 
                          bin_min, bin_max)
        
        if number_file:
            nuc_cell_num = bintotal(bin_output)
            correction = num_correct(case, cell_type, bin_output, 
                                    nuc_cell_num, number_data)
            bin_out_corr = bin_correct(bin_output, correction)      
        else:
            bin_out_corr = bin_output
        
        # Write row for indvidual cell type    
        bin_track = 0
        cell_row = [case, cell_type]
        for nuc_bin in bin_out_corr:
            cell_row.append(nuc_bin)
            cell_sum[bin_track] = cell_sum[bin_track] + nuc_bin 
            bin_track += 1
        output_writer.writerow(cell_row)
    
    # Write row for all cell types within a case    
    case_output = [case, "All"]
    for nuc_bin in cell_sum:
        case_output.append(nuc_bin)
    output_writer.writerow(case_output)
    
print "Done"
