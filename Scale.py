from xlrd import *
from xlwt import *
from tempfile import TemporaryFile
import time

def spacer (string, length):
    ''' get a string and return a string that is length spaces long and 
    added spaces to make up the difference are empty spaces ' '
    '''
    counter = length - len (string)

    while counter > 0:
        string += ' '
        counter -= 1

    return string

class Sample:
    def __init__ (self, date, time, sample_num, weight):
        self.date = date
        self.time = time
        self.sample_num = sample_num
        self.weight = weight
    
    def __str__ (self):
        date = spacer (self.date, 15)
        time = spacer (self.time, 10)
        sample_num = spacer (self.sample_num, 5)
        weight = spacer (self.weight, 7)

def extract_data(filename):
    weight_samples = [] # List of W_sample objects
        
    # Create a stop point at the end of the file
    data_file = open(filename, 'a')
    data_file.write('\n Mozart')
    data_file.close()
    
    # Reading the file and grabbing info
    data_file = open(filename, 'r')

    # Default values for Sample object
    date, time, sample_num, weight = 'N/A', 'N/A', 'N/A', 'N/A'
    
    line = []
    while 'Mozart' not in line:
        
        line = data_file.readline().strip("\n").split()    

        # Getting our 'global values' (date/time weights were taken)
        if 'Weighing' in line:
            date_raw = data_file.readline ().strip ("\n").split ()   
            # assert (time == date_raw [len (date_raw) - 1], "time = 0/0/0")
            date = date_raw [0] + " " + date_raw [1]
        
        # Getting actual data IF it is in the line
        elif 'g' in line:            
            end_index = line.index('g')
            weight = float (line[end_index - 1])
            if len(line) == 3 or len(line) == 4:
                sample_num = int(line [0])
                
            sample = Sample(date, time, sample_num, weight)
            weight_samples.append(sample)
            # Reset default values - will be overwritten (theoretically)
            weight = 'N/A'
            sample_num = 'N/A'
       
    return weight_samples

def add_sheet(output_file, samples_list, directory, filename):
    '''
    Given a list of samples (samples_list) and open excel file (output_file),
        add a new sheet to output_file. New sheet contains samples_list,
        with each list item written into its own row (in order).
        Sheet name is the filename
    '''   
    
    # Set the style for specific cells that will delineate data on the page
    borders = Borders ()
    borders.bottom = Borders.THIN
    style_bot = XFStyle ()
    style_bot.borders = borders 
    
    sheet_name = filename.strip(".txt")
    output_sheet = output_file.add_sheet(sheet_name, cell_overwrite_ok=True)
    
    # Writing the headers in the sheet with the extracted data
    # Two seperate data sets are written; Samples + Quality Control measures
    header_titles = ['Sample #', 'Weight'] 
    
    for x in range(len(header_titles)):
                output_sheet.write (0, x, header_titles [x], style_bot)   
                
    sample_row = 1            
    for obj in samples_list:
        output_sheet.write (sample_row, 0, obj.sample_num)
        output_sheet.write (sample_row, 1, obj.weight)
        sample_row += 1
    
    sample_row += 1
    output_sheet.write (
        sample_row, 0,
        "Data collected on " + samples_list[0].date + " " + samples_list[0].time)
    output_sheet.write (
        sample_row + 1, 0,
        "Data extracted on " + time.strftime ("%d.%b %Y %H:%M"))
