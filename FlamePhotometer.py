from xlrd import *
from xlwt import *
from tempfile import TemporaryFile
import time

def spacer(string, length):
    ''' get a string and return a string that is length spaces long and 
    added spaces to make up the difference are empty spaces ' '
    '''
    counter = length - len (string)

    while counter > 0:
        string += ' '
        counter -= 1

    return string

class Sample:
    def __init__ (self, date, sample_num, mode, time, K, Na):
        self.date = date
        self.time = time
        self.sample_num = sample_num
        self.K = K
        self.Na = Na
        self.mode = mode
    
    def __str__ (self):
        date = spacer (self.date, 15)
        time = spacer (self.time, 10)
        sample_num = spacer (self.sample_num, 5)
        K = spacer (self.K, 7)
        Na = spacer (self.Na, 7)
        mode = spacer (self.mode, 7)
        
        return date + time + sample_num + mode + 'Na: ' + Na + 'K: ' + K

def extract_data(filename):
    
    tissue_samples = [] # List of Sample objects
        
    # Create a stop point at the end of the file
    data_file = open (filename, 'a')
    data_file.write ('\n Mozart')
    data_file.close ()
    
    # Reading the file and grabbing info
    data_file = open (filename, 'r')

    # Values for first TA_sample (null object)
    date = 'N/A'
    sample_num = 'N/A'
    mode = 'N/A'
    time = 'N/A'
    Na = 'N/A'
    K = 'N/A'   
    
    line = ''
    while line.find ('Mozart') == -1:
        
        line = data_file.readline ().strip ("\n")    
            
        if line.find ('Date:') != -1:
            start_index = line.find (':') + 1
            date = line [start_index:]              
        
        elif line.find ('Sample:') != -1:
            sample = Sample(date, sample_num, mode, time, K, Na)
            tissue_samples.append(sample)
            # Reset default values for Na/K measurement-should be overwritten
            Na = 'N/A'
            K = 'N/A'
            sample_num = line.split ()[1]
            
        elif line.find ('Mode:') != -1:
            mode = line.split ()[1]
        elif line.find ('Time:') != -1:
            time = line.split ()[1]
        elif line.find ('Na') != -1:
            if line.find ('E') == -1:
                Na = line.split ()[1]              
            else:
                start_index = line.find ('=') + 1
                Na = line [start_index:]
                            
        elif line.find ('K') != -1:
            if line.find ('E') == -1:
                K = line.split ()[2]              
            else:
                start_index = line.find ('=') + 1
                K = line [start_index:]                      
                
    # This is neccissary because finding the word 'sample' is usually required 
    # to initiate saving of a previous sample 
    # (This grabs the last sample that is ignored by the main loop)
    sample = Sample(date, sample_num, mode, time, K, Na)
    tissue_samples.append (sample)        
    tissue_samples.pop (0)
        
    return tissue_samples

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
    
    output_sheet = output_file.add_sheet(
        filename.strip (".txt"), cell_overwrite_ok=True)
    
    # Writing the headers in the sheet with the extracted data
    # Two seperate data sets are written; Samples + Quality Control measures
    header_titles = ['Time','Mode', 'Sample #', 'Na', 'K'] 
    nonsample_header_titles = ['Time','Mode','Sample Type','B4 Sample #','Na','K']

    for x in range(len(header_titles)):
                output_sheet.write (0,x + 6, header_titles [x], style_bot)   
    for x in range(len(nonsample_header_titles)):
                output_sheet.write (
                    0,x, nonsample_header_titles [x], style_bot)    
    
    sample_row = 1            
    nonsample_row = 1
    before_sample_num = 0
    for obj in samples_list:
                        
        if obj.sample_num [0] == '<':
            obj_things = [obj.time, obj.mode, obj.sample_num,\
                before_sample_num, obj.Na, obj.K]
            for col in range (len(obj_things)):
                output_sheet.write (nonsample_row, col, obj_things [col])
                col += 1
            nonsample_row += 1            
            
        else:
            obj_things = [obj.time, obj.mode, obj.sample_num, obj.Na, obj. K]
            for col in range (len(obj_things)):
                output_sheet.write (sample_row, col + 6, obj_things [col])
                col += 1
            sample_row += 1
            before_sample_num = float(obj.sample_num) + 1
            
    output_sheet.write (
        nonsample_row + 1, 0, "Data collected on " + samples_list [0].date)
    output_sheet.write (
        nonsample_row + 2, 0, "Data extracted on " + time.strftime ("%d-%b-%Y"))    
     
    