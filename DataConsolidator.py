''' Change log

0.4 - Corrected to catch error messages (<<EXX>>) as values for K/Na
    - Changed output settings available (2 settings)
0.5 - Scope of data extraction expanded to scales and gamma counter outputs

'''
import os,sys
import time
import wx
from xlrd import *
from xlwt import *
from tempfile import TemporaryFile

import FlamePhotometer
import GammaCounter
import Scale

class AboutDialog(wx.Dialog):
    def __init__(self, parent, id, title):
        wx.Dialog.__init__(self, parent, id, title)
        
        self.rootPanel = wx.Panel(self)        
        innerPanel = wx.Panel(self.rootPanel,-1, size=(500,260), style=wx.ALIGN_CENTER)
        hbox = wx.BoxSizer(wx.HORIZONTAL) 
        vbox = wx.BoxSizer(wx.VERTICAL)
        innerBox = wx.BoxSizer(wx.VERTICAL)
        
        txt1 = wx.StaticText(
            innerPanel, id=-1,
            label="     Data Extractor     ",style=wx.ALIGN_CENTER, name="")
        txt2 = wx.StaticText(
            innerPanel, id=-1,
            label="     Version 1.0 as of August 15th, 2016     ",
            style=wx.ALIGN_CENTER, name="")
        txt3 = wx.StaticText(
            innerPanel, id=-1,
            label="     Copyright 2016 Ruben Flam-Shepherd. All rights reserved.     ",
            style=wx.ALIGN_CENTER, name="")
        txt7 = wx.StaticText(
            innerPanel, id=-1,
            label="     captured into a .txt file by RealTerm) and write the data into     ",
            style=wx.ALIGN_CENTER, name="")
        txt8 = wx.StaticText(
            innerPanel, id=-1,
            label="     an excel (.xls) file.     ",
            style=wx.ALIGN_CENTER, name="")
        txt9 = wx.StaticText(
            innerPanel, id=-1,
            label="     More detailed information about how to do this can be found     ",
            style=wx.ALIGN_CENTER, name="")
        txt10 = wx.StaticText(
            innerPanel, id=-1,
            label="     in the 'README.txt' file found in the installation directory     ",
            style=wx.ALIGN_CENTER, name="")
        btn1 = wx.Button(innerPanel, id=1, label="Close")
        
        line1 = wx.StaticLine(innerPanel, -1, style=wx.LI_HORIZONTAL)
        line2 = wx.StaticLine(innerPanel, -1, style=wx.LI_HORIZONTAL)
        
        self.Bind(wx.EVT_BUTTON, self.OnClose, id=1)
        
        innerBox.AddSpacer((150,10))
        innerBox.Add(txt1, 0, wx.CENTER)
        innerBox.AddSpacer((150,6))
        innerBox.Add(txt2, 0, wx.CENTER)
        innerBox.AddSpacer((150,6))
        
        innerBox.Add(line1, 0, wx.CENTER|wx.EXPAND)
        innerBox.AddSpacer((150,6))
        innerBox.Add(txt7, 0, wx.CENTER)
        innerBox.Add(txt8, 0, wx.CENTER)
        innerBox.AddSpacer((150,6))
        innerBox.Add(txt9, 0, wx.CENTER)
        innerBox.Add(txt10, 0, wx.CENTER)
        innerBox.AddSpacer((150,6))
        
        innerBox.Add(line2, 0, wx.CENTER|wx.EXPAND)
        innerBox.AddSpacer((150,6))
        innerBox.Add(txt3, 0, wx.CENTER)
        innerBox.AddSpacer((150,6))        
        innerBox.Add(btn1, 0, wx.CENTER)
        innerBox.AddSpacer((150,6))
        
        innerPanel.SetSizer(innerBox)

        hbox.Add(innerPanel, 0, wx.ALL|wx.ALIGN_CENTER)
        vbox.Add(hbox, 1, wx.ALL|wx.ALIGN_CENTER, 5)
        

        self.rootPanel.SetSizer(vbox)
        vbox.Fit(self)        
        
    def OnClose(self, event):
        self.Close()

class MyFrame(wx.Frame):
        
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title)

        self.rootPanel = wx.Panel(self)

        innerPanel = wx.Panel(
            self.rootPanel,-1,
            size=(500,160),
            style=wx.ALIGN_CENTER)
        hbox = wx.BoxSizer(wx.HORIZONTAL) 
        vbox = wx.BoxSizer(wx.VERTICAL)
        innerBox = wx.BoxSizer(wx.VERTICAL)
        buttonBox = wx.BoxSizer(wx.HORIZONTAL)
        radioBox = wx.BoxSizer(wx.HORIZONTAL)

        # I want this line visible in the CENTRE of the inner panel
        txt1 = wx.StaticText(
            innerPanel, id=-1,
            label="     Welcome to the Data Extractor!     ",
            style=wx.ALIGN_CENTER, name="")
        txt2 = wx.StaticText(
            innerPanel, id=-1,
            label="Please choose an option below:",
            style=wx.ALIGN_CENTER, name="")        
        txt3 = wx.StaticText(innerPanel, id=-1,
            label="Note: .xls output file will be written in the same folder",
            style=wx.ALIGN_CENTER, name="")
        txt4 = wx.StaticText(
            innerPanel, id=-1, 
            label="that data is being extracted from",
            style=wx.ALIGN_CENTER, name="")
        
        font3 = wx.Font (7, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        txt3.SetFont (font3)
        txt4.SetFont (font3)
        
        # Option Buttons        
        btn1 = wx.Button (innerPanel, id=1, label="Extract Data")
        btn2 = wx.Button (innerPanel, id=2, label="About")
        btn4 = wx.Button (innerPanel, id=4, label="Quit")
        
        self.Bind(wx.EVT_BUTTON, self.OnOpenFile, id=1)        
        self.Bind(wx.EVT_BUTTON, self.OnAbout, id=2)        
        self.Bind(wx.EVT_BUTTON, self.OnClose, id=4)
                
        innerBox.AddSpacer((150,15))
        innerBox.Add(txt1, 0, wx.CENTER)
        innerBox.AddSpacer((150,15))
        innerBox.Add(txt2, 0, wx.CENTER)
        innerBox.AddSpacer((150,15))
        
        # Main program options
        buttonBox.AddSpacer(7,10)        
        buttonBox.Add(btn1, 0, wx.CENTER)
        buttonBox.AddSpacer(7,10)
        buttonBox.Add(btn2, 0, wx.CENTER)
        buttonBox.AddSpacer(7,15)
        buttonBox.Add(btn4, 0, wx.CENTER)
        buttonBox.AddSpacer(7,10)        
        innerBox.Add(buttonBox, 0, wx.CENTER)
        innerBox.AddSpacer ((150,10))
        
        # Disclaimer of where data is extracted to
        innerBox.Add(txt3, 0, wx.CENTER)
        innerBox.Add(txt4, 0, wx.CENTER)
        innerBox.AddSpacer((150,10))        
        innerPanel.SetSizer(innerBox)

        hbox.Add(innerPanel, 0, wx.ALL|wx.ALIGN_CENTER)
        vbox.Add(hbox, 1, wx.ALL|wx.ALIGN_CENTER, 5)
        

        self.rootPanel.SetSizer(vbox)
        vbox.Fit(self)
    
    def OnClose(self, event):
        self.Close()
        
    def OnOpenFile(self, event):
        dlg = wx.FileDialog(self,
            "Choose a file to extract data from", os.getcwd(),
            "", "", wx.FD_MULTIPLE)
        if dlg.ShowModal() == wx.ID_OK:
            directory, filenames = dlg.GetDirectory(), dlg.GetFilenames()
                                    
            # format the directory (and path) to unicode w/ forward slash so
            # it can be passed between methods/classes w/o bugs
            directory = u'%s' %directory
            directory = directory.replace (u'\\', '/')
            
            # Create the file to save everything in
            output_file = Workbook ()
            
            # Saving data in files as seperate sheets in the output_file
            for filename in filenames:
                path = '/'.join ((directory, filename))                                              
                if filename.find('FP') != -1:
                    samples_list = FlamePhotometer.extract_data(path)
                    FlamePhotometer.add_sheet(
                        output_file, samples_list, directory, filename)
                elif filename.find('DW') != -1 or \
                        filename.find ('FW') != -1 or \
                        filename.find ('DI') != -1 or \
                        filename.find ('CATE') != -1 or \
                        filename.find ('TW') != -1:
                    samples_list = Scale.extract_data(path)
                    Scale.add_sheet(
                        output_file, samples_list, directory, filename)            
                elif filename.find('NA24') != -1 \
                        or filename.find ('K42') != -1:
                    samples_list = GammaCounter.extract_cobra_data(path)
                    GammaCounter.add_sheet(
                        output_file, samples_list, directory, filename)
                elif filename.find('A07') != -1 \
                        or filename.find ('A05') != -1:
                    samples_list = GammaCounter.extract_wizard_data(path)
                    GammaCounter.add_sheet(
                        output_file, samples_list, directory, filename)                    
                    
            output_name = 'OUTPUT - ' + time.strftime("(%Y_%m_%d).xls")
            output_path = '/'.join((directory, output_name))    
        
            output_file.save(output_path)
            output_file.save(TemporaryFile())            

            dlg.Destroy()
            self.Close()           
    
    def OnAbout(self, event):
        dlg = AboutDialog(self, -1, 'About')
        dlg.SetIcon(wx.Icon('fire.ico', wx.BITMAP_TYPE_ICO))
        val = dlg.ShowModal()
        dlg.Destroy()
        
class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, -1, 'Data Extractor')
        frame.SetIcon(wx.Icon('fire.ico', wx.BITMAP_TYPE_ICO))
        frame.Show(True)
        frame.Center()
        self.SetTopWindow(frame)
        return True

if __name__ == '__main__':
    app = MyApp(0)
    app.MainLoop()