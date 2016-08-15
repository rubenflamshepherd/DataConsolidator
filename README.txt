Instructions for Capturing Serial Output from the Flame Photometer

Inital Setup (Only to be done the first time capturing serial output)

1. Install the serial-to-USB adapter driver (PL2303_Vista_32_64_332102.exe in the Flame Photometer Data Extractor's (FPDE) installation folder)
2. Make sure your computer recognizes the adapter when it’s plugged in by itself
	i.Found under ‘Start’>’Devices and Printers’
3. Install Realterm (Realterm_2.0.0.70_setup.exe in the FPDE's installation folder or from http://realterm.sourceforge.net/)

Capturing Data

1. With the flame photometer off, use the Serial-to-USB adapter to connect the serial port on the flame	photometer to the USB port on your computer 
	i. You will probably have to use the USB extension cord to do this (both should already be attached in series to the flame photometer)
2. Open Realterm
3. Under the ‘Port’ Tab
	i. change ‘Baud’ to ‘9600’
	ii. Change ‘Hardware Flow Control’ from ‘None’ to ‘RTS/CTS’
	iii. Press ‘Change’
4. Under the ‘Capture’ Tab
	i. change the ‘File’ entry field to where you would like data to be saved in a text file
		- NOTE: Realterm will not create the directory/file you have specified. To make sure this works you must create the directory/file outside of Realterm.
	ii. Uncheck the box next to ‘Direct Capture’ in order to see the output of the flame photometer on the Realterm output screen
	iii. Press ‘Start: Overwrite’ just before you turn on the flame photometer to start capturing its output.
		- NOTE: You will have to specify these settings (steps 2-5) every time you open Realterm, they do not save.
5. Turn on the flame photometer, aspirate standards/samples pressing 'Measure' to send the serial output of the data to RealTerm to be captured.
		- NOTE: Be sure to confirm that RealTerm is recognizing the data and saving ti the desired file

Extracting Data

1. Open the DataConsolidator
2. Select the file that was created by RealTerm using the serial output from the flame photometer.
3. Select whether you would like the data extracted into one column or two columns.
4. Press 'Extract Data'

Helpful Hints:

- Holding the 'Measure' button until the machine beeps twice will save a measurement as a quality control with '<<QQQ>>' as its sample number
- Pressing the 'Measure' button twice quickly will save a measurement as a repeat with '<<RRR>>' as is sample number
	i. Repeat and Quality control measures are extracted by the FPDE into their own column with the sample number that they proceed indicated
- Pressing both the '+' buttons simultaneously will reset sample number to 1