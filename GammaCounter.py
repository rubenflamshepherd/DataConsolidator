from xlrd import *
from xlwt import *
from tempfile import TemporaryFile
import time, os.path

def extract_cobra_data(filename):
    '''
    Given a raw .csv file from the gamma counter (Cobra Quantum 5003; 
        PerkinElmer Inc., Waltham, MA, USA), return a list where each entry is
        the split(',') contents of each line from the raw .csv file
    '''
    
    cpm_samples = [] # List to contain .csv contents
    # First entry in list is date raw .csv file created   
    cpm_samples.append(time.ctime(os.path.getctime(filename)))
    # Reading the file and grabbing info
    data_file = open(filename, 'r')
    # Default values for list entries
    cpm = 'N/A'
    index = 7    
    line_raw = 'start'

    while line_raw != '': # End of raw .csv file        
        line_raw = data_file.readline ().strip ("\n")
        line = line_raw.split (",")    
        cpm = line
        cpm_samples.append(cpm)       
           
    return cpm_samples

def extract_wizard_data(filename):
    '''
    Similiar functionality to extractor() function above, but refactored for
        use with newer gamma counter (Wallac 1480 Wizard 30; PerkinElmer Inc.,
        Turku, Finland)
    '''
    
    cpm_samples = [] # List of W_sample objects
   
    cpm_samples.append (time.ctime(os.path.getctime(filename)))
    # Reading the file and grabbing info
    for line_raw in open (filename, 'r'):
        line = line_raw.split()   
        cpm_samples.append (line)
           
    return cpm_samples

def add_sheet(output_file, samples_list, directory, filename):
    '''
    Given a list of samples (samples_list) and open excel file (output_file),
        add a new sheet to output_file. New sheet contains samples_list,
        with each list item written into its own row (in order).
        Sheet name is the filename
    '''    
    # Set the style for specific cells that will delineate data on the page
    borders = Borders()
    borders.bottom = Borders.THIN
    style_bot = XFStyle()
    style_bot.borders = borders 
    
    sheet_name = filename.strip(".txt")
    output_sheet = output_file.add_sheet(sheet_name, cell_overwrite_ok=True)

    # Writing the headers in the sheet with the extracted data
    # Data is present in a 2D where each list entry should be its own cell
    for x in range (1, len(samples_list)):
        for y in range (0, len(samples_list[x])):
            try:
                output_sheet.write(x-1, y, float(samples_list[x][y]))
            except:
                output_sheet.write(x-1, y, samples_list[x][y])
            
    output_sheet.write (
        len (samples_list) - 1 , 0, "Data collected on " + samples_list [0]) 
    output_sheet.write (
        len (samples_list), 0,
        "Data extracted on " + time.strftime ("%a %b %d %H:%M:%S %Y"))
              
