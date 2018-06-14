import pandas
import wx
import wx.lib.intctrl as INTCTRL
import re
import openpyxl
import os


class file_object(object):
    def __init__(self, *args, **kwargs):
        self.sheet_name = None
        self.start_row = 1
        self.file_name = None
        self.file_path = None
        self.sheet_names =[]
        self.file_type = None
        self.initial_dataframe = pandas.DataFrame()
        self.columns_type = None
        self.values_type = None
        self.directory_path = None



class Linearization_GUI(wx.Frame):
    
    
    def __init__(self, parent, title, *args, **kwargs):
        
        super(Linearization_GUI, self).__init__(parent, title=title, size = (400, 275), *args, **kwargs) 
        
        
        self.InitUI()
        self.Centre()
        self.Show()

        
    def InitUI(self):
        
        self.file_object = file_object()
        
        ##############
        ## Menu Bar
        ##############
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        
        ## Add items to file menu.
        open_item = fileMenu.Append(wx.ID_OPEN, "&Open", "Open Excel or .csv File")
        self.Bind(wx.EVT_MENU, self.OnOpen, open_item)
        
        fileMenu.AppendSeparator()
        
        quit_item = fileMenu.Append(wx.ID_EXIT, '&Quit', 'Quit Application')
        self.Bind(wx.EVT_MENU, self.OnQuit, quit_item)
        
        
        ## Add file menu to the menu bar.
        menubar.Append(fileMenu, "&File")
        ## Put menu bar in frame.
        self.SetMenuBar(menubar)



        panel = wx.Panel(self)
        
        vbox = wx.BoxSizer(wx.VERTICAL)

        settings_flexbox = wx.FlexGridSizer(3, 2, 10, 10)
        
        ## Add a status bar to the bottom for reporting messages.
        #self.statusbar = self.CreateStatusBar()
        
        
        ##############
        ## Current File Display
        ##############
        current_file_header = wx.StaticText(panel, label = "Current File:")
        self.current_file_label = wx.StaticText(panel, label = "")
        
        vbox.Add(current_file_header, flag=wx.ALIGN_LEFT | wx.LEFT | wx.TOP, border=10)
        vbox.Add(self.current_file_label, flag=wx.ALIGN_LEFT | wx.LEFT, border=10)

        
        ##############
        ## Current Sheet Name
        ##############
        current_sheet_header = wx.StaticText(panel, label = "Current Sheet:")
        self.current_sheet_label = wx.StaticText(panel, label = "")
        
        vbox.Add(current_sheet_header, flag=wx.ALIGN_LEFT | wx.LEFT | wx.TOP, border=10)
        vbox.Add(self.current_sheet_label, flag=wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        
        ##############
        ## Start Row
        ##############
        start_row_label = wx.StaticText(panel, label = "Start Row:")
        self.start_row_entry = INTCTRL.IntCtrl(panel, value = self.file_object.start_row, min = 1, limited = True)
        
        #self.start_row_entry.Bind(wx.EVT_TEXT, self.On_Start_Row)
        
        #vbox.Add(start_row_label, flag=wx.ALIGN_LEFT | wx.LEFT | wx.TOP, border=10)
        #vbox.Add(self.start_row_entry, flag=wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        
        ##############
        ## Columns Type
        ##############
        columns_type_label = wx.StaticText(panel, label = "Type of Columns:")
        self.columns_type_entry = wx.TextCtrl(panel, value = "Compounds")
        
        #self.columns_type_entry.Bind(wx.EVT_TEXT, self.On_Columns_Type)
        
        #vbox.Add(columns_type_label, flag=wx.ALIGN_LEFT | wx.LEFT | wx.TOP, border=10)
        #vbox.Add(self.columns_type_entry, flag=wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        
        ##############
        ## Values Type
        ##############
        values_type_label = wx.StaticText(panel, label = "Type of Values:")
        self.values_type_entry = wx.TextCtrl(panel, value = "Concentration")
        
        #self.values_type_entry.Bind(wx.EVT_TEXT, self.On_Values_Type)
        
        #vbox.Add(values_type_label, flag=wx.ALIGN_LEFT | wx.LEFT | wx.TOP, border=10)
        #vbox.Add(self.values_type_entry, flag=wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        
        settings_flexbox.AddMany([(start_row_label, wx.ALIGN_BOTTOM | wx.EXPAND), (self.start_row_entry), 
                                  (columns_type_label, wx.ALIGN_BOTTOM | wx.EXPAND), (self.columns_type_entry), 
                                  (values_type_label, wx.ALIGN_BOTTOM | wx.EXPAND), (self.values_type_entry)])
    
        vbox.Add(settings_flexbox, flag = wx.ALIGN_CENTER | wx.ALL, border = 10)

        
        
        ##############
        ## Linearize Data Button
        ##############
        linearize_button = wx.Button(panel, label = "Linearize Data")
        linearize_button.Bind(wx.EVT_BUTTON, self.OnLinearize)
        
        
        vbox.Add(linearize_button, flag = wx.ALIGN_CENTER | wx.ALL, border = 10)
        

        panel.SetSizer(vbox)
        panel.Fit()
        
        


    
    
    def OnQuit(self, e):
        self.Close()

        
        
    
    
    
    def OnLinearize(self, event):
        
        if self.file_object.initial_dataframe.empty:
            message = "No File Selected"
            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
        
        else:
            df_list = []
            df = self.file_object.initial_dataframe
            
            columns_type = self.file_object.columns_type
            values_type = self.file_object.values_type

            for i in range(len(df)):
                temp_df = pandas.DataFrame(index = range(len(df.columns)-1), columns = [df.columns[0], columns_type, values_type])
                
                temp_df.loc[:, df.columns[0]] = df.loc[:, df.columns[0]].iloc[i]
                temp_df.loc[:, columns_type] = df.columns[1:]
                temp_df.loc[:, values_type] = df.iloc[i,1:].values
                
                df_list.append(temp_df)
    
            final_form = pandas.concat(df_list)
            

      ###### CSV ########      
            
            if self.file_object.file_type == "csv":
                
                dlg = wx.FileDialog(
                self, message="Save file as ...", 
                defaultDir=self.file_object.directory_path, 
                defaultFile="", style=wx.FD_SAVE
                )
                
                if dlg.ShowModal() == wx.ID_OK:
                    path = dlg.GetPath()
                    dlg.Destroy()
                    
                    if not re.match(r".*\.csv$", path):
                        path = path + ".csv"
                
                    final_form.to_csv(path, index = False)
                else:
                    dlg.Destroy()
                    return
                

      ####### EXCEL #########      
            
            else:
                
                sheet_name = ""
                while sheet_name == "":
                    dlg = wx.TextEntryDialog(None, "Enter the name of the sheet to add:", "Sheet Name")
                    if dlg.ShowModal() == wx.ID_OK:
                        sheet_name = dlg.GetValue()
                        if sheet_name == "" or len(sheet_name) > 31 :
                            message = "Please enter a valid sheet name."
                            msg_dlg = wx.MessageDialog(None, message, "Invalid Sheet Name", wx.OK | wx.ICON_EXCLAMATION)
                            msg_dlg.ShowModal()
                            msg_dlg.Destroy()
                            
                        elif sheet_name in self.file_object.sheet_names:
                            message = "Sheet name is already in workbook. Overwrite the sheet?"
                            msg_dlg = wx.MessageDialog(None, message, "Sheet Name In Use", wx.YES_NO | wx.ICON_QUESTION)

                            if msg_dlg.ShowModal() == wx.ID_NO:
                                sheet_name = ""
                            msg_dlg.Destroy()
            
                        dlg.Destroy()
                    else:
                        dlg.Destroy()
                        return
                    
                    
                    
                book = openpyxl.load_workbook(self.file_object.file_path)
                if sheet_name in self.file_object.sheet_names:
                    book.remove_sheet(sheet_name)
                writer = pandas.ExcelWriter(self.file_object.file_path, engine = "openpyxl")
                writer.book = book
                final_form.to_excel(writer, sheet_name = sheet_name, startrow = 0, startcol = 0, index = False)
                writer.save()

        
        
        
        
  


    def OnOpen(self, event):
                
        
        if self.start_row_entry.GetValue() == 0:
            message = "Please enter a value for the Start Row."
            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            
        elif self.columns_type_entry.GetValue() == "":
            message = "Please enter a value for the Type of Columns."
            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            
        elif self.values_type_entry.GetValue() == "":
            message = "Please enter a value for the Type of Values."
            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            
        else:
        
            message = "Select Excel or .csv File"
            dlg = wx.FileDialog(None, message = message, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
    
            if dlg.ShowModal() == wx.ID_OK:
                filepath = dlg.GetPath()
                
                if re.match(r".*\.xls$", filepath) :
                    message = "The old .xls Excel format is not supported. Please convert to the more recent .xlsx format. "
                    msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                
                elif not re.match(r".*\.xlsx$|.*\.xlsm$|.*\.csv$", filepath) :
                    message = "Unknown File Type."
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                    
                else:
                    
                    directory_path = os.path.split(filepath)[0]
                    
                    if re.match(r".*\.csv", filepath):
                        
                        self.file_object = file_object()
                        
                        self.file_object.start_row = self.start_row_entry.GetValue()
                        
                        
                        df = pandas.read_csv(filepath, skiprows = self.file_object.start_row-1)
                        
                        if df.empty:
                            message = "There is no data in the selected file."
                            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                            msg_dlg.ShowModal()
                            msg_dlg.Destroy()
                            
                            return
                        
                        
                        
                        
                        self.file_object.columns_type = self.columns_type_entry.GetValue()
                        self.file_object.values_type = self.values_type_entry.GetValue()
                        self.file_object.file_type = "csv"
                        self.file_object.file_path = filepath
                        self.current_file_label.SetLabel(filepath)
                        self.current_sheet_label.SetLabel("N\\A")
                        self.file_object.directory_path = directory_path
                        
                        self.file_object.initial_dataframe = df
                        
                        if any([re.match("Unnamed", column_name) for column_name in self.file_object.initial_dataframe.columns]):
                            message = "There are unnamed columns in the data. Are you sure the Start Row value was correct?"
                            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                            msg_dlg.ShowModal()
                            msg_dlg.Destroy()
                        
                        
                    else:
                        workbook = pandas.ExcelFile(filepath)
                        
                        dlg = wx.SingleChoiceDialog(None, "Select a sheet:", "Sheets in the Excel File", workbook.sheet_names)
                    
                        if dlg.ShowModal() == wx.ID_OK:
                            
                            self.file_object = file_object()
                            
                            self.file_object.start_row = self.start_row_entry.GetValue()
                            
                            
                            df = pandas.read_excel(filepath, sheetname = dlg.GetStringSelection(), skiprows = self.file_object.start_row-1)
                        
                            if df.empty:
                                message = "There is no data in the selected sheet."
                                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                                msg_dlg.ShowModal()
                                msg_dlg.Destroy()
                                
                                return
                            
                            
                            
                            
                            self.file_object.columns_type = self.columns_type_entry.GetValue()
                            self.file_object.values_type = self.values_type_entry.GetValue()
                            self.file_object.sheet_name = dlg.GetStringSelection()
                            self.file_object.file_type = "Excel"
                            self.file_object.file_path = filepath
                            self.current_file_label.SetLabel(filepath)
                            self.current_sheet_label.SetLabel(dlg.GetStringSelection())
                            self.file_object.sheet_names = workbook.sheet_names
                            self.file_object.directory_path = directory_path
                            
                            self.file_object.initial_dataframe = df
                            
                            if any([re.match("Unnamed", column_name) for column_name in self.file_object.initial_dataframe.columns]):
                                message = "There are unnamed columns in the data. Are you sure the Start Row value was correct?"
                                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                                msg_dlg.ShowModal()
                                msg_dlg.Destroy()

                    
                    

    


def main():

    ex = wx.App(False)
    Linearization_GUI(None, title = "Table Linearization")
    ex.MainLoop()    


if __name__ == '__main__':
    main()







#df = pandas.read_excel("/Users/higashi/Desktop/2018_04_16_luminex rerun media samples.xls", sheet_name = "Average results", skiprows=1)
#
#df.dropna(inplace=True)
#
#df.drop(axis=1, columns = "Sample", inplace=True)
#
#df_list = []
#
#for i in range(len(df)):
#    temp_df = pandas.DataFrame(index = range(len(df.columns)-1), columns = ["Sample_Name", "Compound", "Concentration"])
#    
#    temp_df.loc[:, "Sample_Name"] = df.loc[:, "Sample_Name"].iloc[i]
#    temp_df.loc[:, "Compound"] = df.columns[1:]
#    temp_df.loc[:, "Concentration"] = df.iloc[i,1:].values
#    
#    df_list.append(temp_df)
#    
#    
#    
#final_form = pandas.concat(df_list)
#
#final_form.to_excel("/Users/higashi/Desktop/2018_04_19_luminex rerun media samples.xls")
    