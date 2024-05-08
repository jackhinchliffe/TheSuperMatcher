# Better Fuzzy Lookup

from thefuzz import fuzz
from thefuzz import process
import jaro
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import time
from threading import Thread


#########################################################################################
#                                                                                       #
#                                 The Super Matcher                                     #
#                                                                                       #
#                                By Jack Hinchliffe                                     #
#                             April 2024, Python 3.8.1                                  #
#                                                                                       #
# Requirements:                                                                         #
# - thefuzz 0.20.0                                                                      #
# - jaro-winkler 2.0.3                                                                  #
#                                                                                       #
# Features:                                                                             #
# - Launches a GUI for ease of use                                                      #
# - Read a single workbook into the application                                         #
# - Writes left table with matched data from right table if found to loaded workbook    #
# - Significantly faster than the Excel Fuzzy Matching Add-in for large data sets       #
# - Same customizable options as Excel Add-in (complete functional replacement)         #
# - Ability to enable smart matching (self deciding) on the results                     #
# - Self decide suggests which are actual matches based on selected column's similarity #
# - Assists in identifying similar data between large datasets quickly                  #
#                                                                                       #
#########################################################################################


################################################
# Classes
################################################

class TheSuperMatcher(tk.Tk):
    """
    Main window of application gui

    Top level class for script
    """

    # Frame construction
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.title("The Super Matcher")
        self.geometry('1000x750')
        self.frames = {}

        for F in ({MainPage}):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame(MainPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()    

class TextRedirector(object):
    """
    Redirects text from sys output to text widget
    """
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.see("end")

    def flush(self):
        pass

class Workbook:
    """
    Class for representing an excel workbook within this script

    Parameters
    ----------
    path : str
        a string path to the file that will be represented by this class

    Constructor will create instances of the Tables class for each sheet present in workbook
    """
    matchedText = 'Rows Matched: 0'

    def __init__(self, path:str) -> None:
        self.path = path
        self.tables = {}
        self.__read()
        print('Workbook Loaded!\n')
    
    def __read(self):
        with pd.ExcelFile(self.path) as xls:
            for i, sheet in enumerate(xls.sheet_names):
                print(f'Reading sheet \'{sheet}\'...')
                sheetdata = pd.read_excel(xls, sheet_name=sheet)
                self.tables[sheet] = Tables(sheet, i)
                self.tables[sheet].readData(sheetdata)

    def getSheets(self) -> list:
        return list(self.tables.keys())

    def numOfSheets(self) -> int:
        return len(self.tables)

    def getPath(self) -> str:
        return self.path
    
    def setMatchedCount(self, count) -> str:
        self.matchedText = f'Rows Matched: {count}'

    def getMatchedCount(self) -> str:
        return self.matchedText
    

class Tables:
    """
    Tables class that represents a workbook's sheets and the data stored in them

    Parameters
    ----------
    table_name : str
        A string name for this instance table
    table_id : int
        An interger id to index tables
    """
    def __init__(self, table_name, table_id) -> None:
        self.table_name = table_name
        self.table_id = table_id
        self.data = pd.DataFrame()

    def __readColumnHeader(self):
        """
        Gets column names from the sheet and adds sheet name to the column header
        """
        print(f"Extracting columns from \'{self.table_name}\'...\n")
        oldColList = self.data.columns.values.tolist()
        newColList = [f"{col}.{self.table_name}" for col in oldColList]
        oldNewDict = {}
        for old, new in zip(oldColList, newColList):
            oldNewDict[old] = new
        self.data.rename(columns=oldNewDict, inplace=True)
        self.columns = self.data.columns.values.tolist()
        if len(self.columns) == 0: 
            self.columns = ['']

    def readData(self, sheetData: pd.DataFrame) -> None:
        """
        Read the data from the selected sheet
        """
        print("Reading table data...")
        self.data = sheetData
        self.__readColumnHeader()
    
    def getHeaders(self) -> list:
        """
        Get the headers (column name) of this table

        Returns
        -------
        list of column names (str)
        """
        return self.columns
    
    def getData(self) -> pd.DataFrame:
        """
        Get the full data of this table
        """
        return self.data
    
    def getName(self) -> str:
        """
        Get the name of this table
        """
        return self.table_name

################################################
# Helper functions
################################################

def chooseFileHandler(filePathLabel:tk.Label) -> str:
    """
    Prompts user to select a file to load

    Parameters
    ----------
    filePathLabel : tk.Label
        the GUI label that should be updated to display the filepath
    -------
    Returns
    file path string
    """
    excelFile = askopenfilename()
    if(excelFile == ""):
        filePathLabel['text'] = ""
    elif(".xlsx" not in excelFile):
        tk.messagebox.showerror("Error", "Incorrect File Format")
        filePathLabel['text'] = ""
        excelFile = ""
    return excelFile

def writeData(fluData:pd.DataFrame, excelFilePath:str, sheet_name:str, decidedTable:pd.DataFrame, decidedSheet_name:str='') -> None:
    """
    Writes dataframe to an excel sheet of file

    Parameters
    ----------
    data : pandas dataframe
        dataframe to write to file
    excelFilePath : str
        path to excel file
    sheet_name : str
        name of sheet that will be created and wrote to
    """
    print("Updating Workboook (This takes a while!)...\n")
    options = {}
    options['strings_to_formulas'] = False
    options['strings_to_urls'] = False

    name = sheet_name
    dname = decidedSheet_name

    if len(sheet_name) > 30: # Shorten name to prevent breaking workbook. Max is 31 char, but need room incase a number is added for page name duplicates.
        name = sheet_name[:30]
    if len(decidedSheet_name) > 30: # Shorten name to prevent breaking workbook
        dname = decidedSheet_name[:30]

    with pd.ExcelWriter(excelFilePath,
                        mode='a',
                        engine='openpyxl', #using openpyxl because xlsxwritter does not support append mode. Openpyxl is super slow though.
                        if_sheet_exists='new',
    ) as writer:
        writer.workbook = load_workbook(excelFilePath)
        fluData.to_excel(excel_writer=writer, sheet_name=name, index=False)
        if not decidedTable.empty:
            decidedTable.to_excel(excel_writer=writer, sheet_name=dname, index=False)
        print("Data Updated.\nClosing Workbook (This takes a while, too!)...\n") 
    print("Workbook Saved and Closed.\n")

def runButtonHandler(Workbook: Workbook, sheet_selection:list, column_selection:list, similarity_threshold:int, matchLimit:int, doSelfDeciding:bool, selfDecideParams:dict=None) -> None:
    """
    Handler function to execute main program logic

    Parameters
    -----------
    Workbook : Worbook
        A workbook object instance from this file's workbook class
    sheet_selection : list
        list of strings, length 2
    column_selection : list
        list of strings, length 2
    similarity_threshold : int
        between 0 and 100
    matchLimit : int
        number of matches to find
    doSelfDediciding : bool
        should run self deciding function
    selfDecideParams : dict
        parameter values for the self decide function
        dict {'parameter': [int, str, str, bool]}
    """
    start_time = time.perf_counter()
    tableData = []
    names = []
    decidedTable = pd.DataFrame()

    for sheet in sheet_selection:
        tableData.append(Workbook.tables[sheet].getData())
        names.append(Workbook.tables[sheet].getName())
    
    matchedtable = fuzzyMatch(tableData, column_selection, similarity_threshold, matchLimit, Workbook)
    if doSelfDeciding and selfDecideParams != None:
        decidedTable = selfDecide(matchedtable, selfDecideParams)

    writeData(matchedtable, Workbook.getPath(), f"FLU_{names[0]}_{names[1]}", decidedTable, f'SD_{names[0]}_{names[1]}')
    print("Runtime Finished.\nexecution time = %s seconds" % (time.perf_counter() - start_time))

#original
def fuzzyMatch(tabledata:list, column_selection:list, similarity_threshold:int, matchLimit:int, wb:Workbook) -> pd.DataFrame:
    """
    Preforms the fuzzy match between two tables

    Parameters
    ----------
    tabledata : list[pd.DataFrame, pd.DataFrame]
        list of the two dataframes to match
    column_selection : list[str,str]
        list of columns to match on
    similarity_threshold : int
        threshold used to decide whether to reject or keep match result
    matchLimit : int
        maximum number of matches that should be returned
    wb : Worbook
        the current Workbook object

    Returns
    ----------
    dataframe of left table with matches from right table and a similarity score column
    """
    left_table:pd.DataFrame = tabledata[0].fillna("")
    right_table:pd.DataFrame = tabledata[1].fillna("")
    matchOnLeft = column_selection[0]
    matchOnRight = column_selection[1]

    left_table_len = len(left_table.index)

    global matchCountText
    matchedCounter = 0
    print('Beginning Fuzzy Matching...\n')

    combinedTable = pd.DataFrame(columns=[*left_table.columns.values, *right_table.columns.values, 'Similarity Score'])
    
    for index, val in left_table.iterrows():
        returnedValues = process.extract(val[matchOnLeft], right_table[matchOnRight], scorer=fuzz.token_sort_ratio, limit=matchLimit)
        filteredValues = []
        skippedPairs = 0
        for pair in returnedValues:
            if pair[1] >= similarity_threshold or len(returnedValues) - skippedPairs == 1:
                filteredValues.append(pair)
            else:
                skippedPairs += 1

        for pair in filteredValues:
            if len(filteredValues) == 1 and pair[1] < similarity_threshold:
                pair = ("", 0)
                combinedrow = left_table.iloc[index]
                combinedrow['Similarity Score'] = pair[1]
                combinedTable = combinedTable.append(combinedrow, ignore_index=True) # NOTE using pd.append() is depreciated
            else:
                row_from_left = left_table.iloc[index]
                row_from_right = right_table.iloc[pair[2]]
                combinedrow = pd.concat([row_from_left, row_from_right])
                combinedrow['Similarity Score'] = pair[1]
                combinedTable = combinedTable.append(combinedrow, ignore_index=True) # NOTE using pd.append() is depreciated
                matchedCounter += 1
                matchCountText.set(f'Row Matches Found: {matchedCounter} | Row: {index+1}/{left_table_len}')
    print(f"Fuzzy Matching Complete, {matchedCounter} matches found.\n")
    return combinedTable


def selfDecide(data:pd.DataFrame,  selfDecideParams:dict) -> pd.DataFrame:
    """
    Self deciding logic

    Parameters
    -----------
    data : DataFrame
        Matched table to run decisions on
    selfDecideParams : dict
        parameters to make decisions with
        dict {'parameter': [int, str, str, bool]}
    
    Returns
    -----------
    Inputted Dataframe with calculated decisions in first column
    """

    print("Beginning Self Deciding On Results...\n")

    data = data.fillna('')

    decisionDF = data.copy()
    
    decisionList = []
    for i in range(len(data)):
    
        numMatched = 0
        matchedColNames = []

        for key, value in selfDecideParams.items():
            threshold = value[0]/100
            leftitem = str(data[value[1]].iloc[i]).upper() # To compare str of same case, upper because jaro.original_metric typo checker only compares capitals
            rightitem = str(data[value[2]].iloc[i]).upper()

            if not (leftitem == rightitem == '' and value[3]):
                ratio = jaro.original_metric(leftitem, rightitem)

                if ratio >= threshold and "" not in value:
                    numMatched +=1
                    matchedColNames.append(value[1])
        
        colStr = ', '.join(matchedColNames)

        if numMatched == 0:
            #decisionList.append(rowString + '0/1 match')
            decisionList.append(f'0/3, Not a Match')
        elif numMatched == 1:
            #decisionList.append(rowString + '1/3 match')
            decisionList.append(f"1/3, Possible Match: {colStr}")
        elif numMatched == 2:
            #decisionList.append(rowString + '2/3 match')
            decisionList.append(f"2/3, Likely Match, confirm: {colStr}")
        elif numMatched == 3:
            #decisionList.append(rowString + '3/3 match')
            decisionList.append(f"3/3, Definite Match")
        else:
            decisionList.append(f'Somehow matched more than possible')


    decisionDF.insert(0, "DECISION", decisionList)
    print('Self Deciding Complete.\n')
    return decisionDF

################################################
# GUI Window
################################################

matchCountText = None # This is a GLOBAL var that will be used to pass info from calculation threads to the GUI

class MainPage(tk.Frame):

    backgroundcolour = "#b2cfbc"

    def initValues(self):
        """
        Initializes class variables upon initial launch of app
        """
        self.similarity_threshold = 50
        self.num_matches = 1
        self.sheetOptions = []
        self.colOptions_1 = []
        self.colOptions_2 = []
        self.excelFilePath = ""
        self.matchLimit = 1
        self.selected_tables = ["",""]
        self.selected_columns = ["",""]
        self.sheetselectorbox_1_val = ''
        self.sheetselectorbox_2_val = ''
        self.selectedsheet1 = tk.StringVar()
        self.selectedsheet2 = tk.StringVar()
        self.selectedcol1 = tk.StringVar()
        self.selectedcol2 = tk.StringVar()
        self.doSmartMatch = tk.BooleanVar(value=False)
        self.matchCountVar = tk.StringVar()
        global matchCountText
        matchCountText = tk.StringVar()
        matchCountText.set('Waiting to run...')
        

    def matchingThreadFnc(self):
        """
        This is executed in seperate running/calculation thread, seperate from GUI thread.

        This thread is for running the background calculations for fuzzy match and self decide without freezing the GUI
        """
        self.runButton['state'] = 'disabled'
        self.info['state'] = 'disabled'
        runButtonHandler(self.thisWorkbook, self.selected_tables, self.selected_columns, self.similarity_threshold, self.matchLimit, self.doSmartMatch.get(), self.compressSelfDecideParams(self.sd_slider, self.sd_listBox, self.sd_compareEmpty))
        self.runButton['state'] = 'normal'
        self.info['state'] = 'disabled'

    def onRunPress(self):
        """
        This is called and executed by clicking the 'Run Program' button

        Creates a seperate thread for background calculations!
        - Specifically the matchingThreadFnc thread
        """
        self.similarity_threshold = int(self.similarity_slider.get())
        self.matchLimit = int(self.limit_spinbox.get())
        self.selected_columns[0] = self.colselectorbox_1.get()
        self.selected_columns[1] = self.colselectorbox_2.get()
        self.selected_tables[0] = self.sheetselectorbox_1.get()
        self.selected_tables[1] = self.sheetselectorbox_2.get()

        matchingThread = Thread(name='MatchingThread', daemon=True, target=self.matchingThreadFnc)
        matchingThread.start()

    def compressSelfDecideParams(self, slider, listbox, noMatchBlank):
        """
        Squishes the three seperate dicts into a single dict of paramter and values

        Returns either single dict of the parameters or None if not running self decide function
        """
        if self.doSmartMatch.get():
            dataForDecider = {
                "Parameter 1": [int, str, str, bool],
                "Parameter 2": [int, str, str, bool],
                "Parameter 3": [int, str, str, bool]
            }
            for key in dataForDecider:
                dataForDecider[key] = [slider[key][0].get(), listbox[key][0].get(listbox[key][0].curselection()), listbox[key][2].get(listbox[key][2].curselection()), noMatchBlank[key][0].get()]
        else:
            dataForDecider = None
        
        return dataForDecider

    def chooseThreadFnc(self):
        """
        This is executed in seperate running/calculation thread, seperate from GUI thread.

        This thread is for reading the worbook and creating this script's objects without freezing the GUI.
        """
        self.thisWorkbook = Workbook(self.excelFilePath)
        self.sheetOptions = self.thisWorkbook.getSheets()
        

        self.sheetselectorbox_1['values'] = self.sheetOptions
        self.sheetselectorbox_2['values'] = self.sheetOptions

        self.selectedsheet1 = self.sheetOptions[0]

        if self.thisWorkbook.numOfSheets() < 2:
            self.selectedsheet2 = self.selectedsheet1
        else:
            self.selectedsheet2 = self.sheetOptions[1]

        self.sheetselectorbox_1_val = self.selectedsheet1
        self.sheetselectorbox_2_val = self.selectedsheet2
        self.sheetselectorbox_1.set(self.selectedsheet1)
        self.sheetselectorbox_2.set(self.selectedsheet2)
        
        self.colselectorbox_1.set(self.thisWorkbook.tables[self.selectedsheet1].getHeaders()[0])
        self.colselectorbox_2.set(self.thisWorkbook.tables[self.selectedsheet2].getHeaders()[0])
        self.colselectorbox_1['values'] = self.thisWorkbook.tables[self.selectedsheet1].getHeaders()
        self.colselectorbox_2['values'] = self.thisWorkbook.tables[self.selectedsheet2].getHeaders()

        self.sheetselectorbox_1.bind('<<ComboboxSelected>>', self.onSheetSelect_1)
        self.sheetselectorbox_2.bind('<<ComboboxSelected>>', self.onSheetSelect_2)
        self.colselectorbox_1.bind('<<ComboboxSelected>>', self.checkToEnableRun)
        self.colselectorbox_2.bind('<<ComboboxSelected>>', self.checkToEnableRun)

        self.popSelfDecideWidgets()
        self.checkToEnableRun()
        self.chooseFileButton['state'] = 'normal'
        self.info["state"] = "disabled"

    def onChooseFilePress(self):
        """
        This is called and executed by clicking the 'Choose File' button

        Creates a seperate thread for background calculations!
        - Specifically the chooseThreadFnc thread
        """
        self.chooseFileButton['state'] = 'disabled'
        self.excelFilePath = chooseFileHandler(tk.Label(self))
        if self.excelFilePath != '':
            self.sourcefilelabeldisp['text'] = self.excelFilePath
            self.info['state'] = 'disabled'

            print(f'Reading Workbook \'{self.excelFilePath}\'')

            creatingWorkbookThread = Thread(target=self.chooseThreadFnc, name='ChooseThread', daemon=True)
            creatingWorkbookThread.start()
        else:
            self.chooseFileButton['state'] = 'normal'
        self.info['state'] = 'disabled'
        

    def onSheetSelect_1(self, event):
        """
        Event handler for what to do when the Left Table is selected from drop down
        """
        if self.excelFilePath == '': return
        self.sheetselectorbox_1_val = event.widget.get()
        #self.colOptions_1 = readColumns(self.excelFile, self.sheetselectorbox_1_val)
        self.colOptions_1 = self.thisWorkbook.tables[self.sheetselectorbox_1_val].getHeaders()
        self.colselectorbox_1['values'] = self.colOptions_1
        self.selectedcol1 = self.colOptions_1[0]
        self.colselectorbox_1.set(self.selectedcol1)
        self.checkToEnableRun()
        self.popSelfDecideWidgets()

    def onSheetSelect_2(self, event):
        """
        Event handler for what to do when the Right Table is selected from drop down
        """
        if self.excelFilePath == '': return
        self.sheetselectorbox_2_val = event.widget.get()
        #self.colOptions_2 = readColumns(self.excelFile, self.sheetselectorbox_2_val)
        self.colOptions_2 = self.thisWorkbook.tables[self.sheetselectorbox_2_val].getHeaders()
        self.colselectorbox_2['values'] = self.colOptions_2
        self.selectedcol2 = self.colOptions_2[0]
        self.colselectorbox_2.set(self.selectedcol2)
        self.checkToEnableRun()
        self.popSelfDecideWidgets()

    def checkToEnableRun(self, event=None):
        """
        Changes run button state by checking if all required info has been entered to enable program to run
        """
        if self.excelFilePath != '':
            if (self.sheetselectorbox_1.get() != self.sheetselectorbox_2.get()) and (self.colselectorbox_1.get() != '' or self.colselectorbox_2.get() != ''):
                self.runButton['state'] = 'normal'
            else:
                self.runButton['state'] = 'disabled'

    def changeWidgetCollectionState(self, obj:dict, state:str):
        """
        Changes state (enabled/disabled) of the widgets in dict based off parameter passed
        """
        for key in obj:
            for i in range(len(obj[key])):
                if isinstance(obj[key][i], tk.Widget):
                    obj[key][i]['state'] = state
    
    def unlockSelfDecideWidgets(self, event=None):
        """
        Event handler that changes state of self-decide related widgets if checkbox is clicked on/off
        """
        if self.doSmartMatch.get():
            self.changeWidgetCollectionState(self.sd_listBox, 'normal')
            self.changeWidgetCollectionState(self.sd_slider, 'normal')
            self.changeWidgetCollectionState(self.sd_compareEmpty, 'normal')
            self.popSelfDecideWidgets()
        else:
            self.changeWidgetCollectionState(self.sd_listBox, 'disabled')
            self.changeWidgetCollectionState(self.sd_slider, 'disabled')
            self.changeWidgetCollectionState(self.sd_compareEmpty, 'disabled')

    def clearlistbox(self, listbox):
        """
        Clears the listbox of values
        """
        for key in listbox:
            for i in range(len(listbox[key])):
                if type(listbox[key][i]) == tk.Listbox:
                    listbox[key][i].delete(0, listbox[key][i].size())

    def popSelfDecideWidgets(self):
        """
        Populates the self decide widgets based on left/right table selections.
        """
        self.clearlistbox(self.sd_listBox)
        try: # Try to populate the listbox. If workbook object hasn't been created yet (no file loaded), pass on AttributeError
            if self.doSmartMatch.get():
                for key in self.sd_listBox:
                    for i in range(len(self.thisWorkbook.tables[self.sheetselectorbox_1_val].getHeaders())):
                        self.sd_listBox[key][0].insert(i, self.thisWorkbook.tables[self.sheetselectorbox_1_val].getHeaders()[i])
                    for i in range(len(self.thisWorkbook.tables[self.sheetselectorbox_2_val].getHeaders())):
                        self.sd_listBox[key][2].insert(i, self.thisWorkbook.tables[self.sheetselectorbox_2_val].getHeaders()[i])
        except AttributeError:
            pass

    def __init__(self, parent, controller):
        """
        MainPage Frame constructor overloading.

        This is where widgets and tkinter vars are created and initialized.
        """

        global matchCountText
        tk.Frame.__init__(self, parent)
        self.initValues()


        ################################
        # Row 0
        ################################
        header = tk.Label(self, text="The Super Matcher", font="Arial 20 bold")
        header.configure(background=self.backgroundcolour)
        header.grid(column=1, row=0, pady=10)

        ################################
        # Row 1
        ################################

        self.chooseFileButton = tk.Button(self, text="Choose File", font="Arial 14", command=self.onChooseFilePress, cursor='hand2')
        self.chooseFileButton.grid(row=2, column=0)

        self.runButton = tk.Button(self, text="Start Matching", font="Arial 14", command=self.onRunPress, cursor='hand2')
        self.runButton.grid(row=2, column=2)
        self.runButton['state'] = 'disabled'

        ################################
        # Row 2
        ################################
        
        self.sourcefilelabel = tk.Label(self)
        self.sourcefilelabel.configure(background=self.backgroundcolour)
        self.sourcefilelabel.grid(column=0, row=3,sticky="nw", padx=15, pady=5)
        self.sourcefilelabel['text'] = 'File Path:'

        self.sourcefilelabeldisp = tk.Label(self)
        self.sourcefilelabeldisp.configure(background=self.backgroundcolour)
        self.sourcefilelabeldisp.grid(column=1, row=3,sticky="nw",pady=5)

        ################################
        # Row 3
        ################################

        self.similarity_slider = tk.Scale(self, from_=0, to=100, orient="horizontal", cursor='hand2')
        self.similarity_slider.grid(row = 6, column=1)
        self.similarity_slider.set(self.similarity_threshold)

        self.limit_spinbox = tk.Spinbox(self, from_=1, to=10) #Max 10 matches allowed
        self.limit_spinbox.grid(row=8, column=1)

        ################################
        # Row 4
        ################################
        
        #combo boxes
        self.sheetselectorbox_1 = ttk.Combobox(self, textvariable=self.selectedsheet1, width=30)
        self.sheetselectorbox_1['values'] = self.sheetOptions
        self.sheetselectorbox_1.state(["readonly"])
        self.sheetselectorbox_1.grid(row=6, column=0)

        self.sheetselectorbox_2 = ttk.Combobox(self, textvariable=self.selectedsheet2, width=30)
        self.sheetselectorbox_2['values'] = self.sheetOptions
        self.sheetselectorbox_2.state(["readonly"])
        self.sheetselectorbox_2.grid(row=6, column=2)

        self.tableLLabel = tk.Label(self)
        self.tableLLabel.configure(background=self.backgroundcolour)
        self.tableLLabel.grid(column=0, row=5,sticky="nw", padx=15)
        self.tableLLabel['text'] = 'Left Table'

        self.tableRLabel = tk.Label(self)
        self.tableRLabel.configure(background=self.backgroundcolour)
        self.tableRLabel.grid(column=2, row=5,sticky="nw", padx=15)
        self.tableRLabel['text'] = 'Right Table'

        self.colLLabel = tk.Label(self)
        self.colLLabel.configure(background=self.backgroundcolour)
        self.colLLabel.grid(column=0, row=7,sticky="nw", padx=15)
        self.colLLabel['text'] = 'Left Matching Column'

        self.colRLabel = tk.Label(self)
        self.colRLabel.configure(background=self.backgroundcolour)
        self.colRLabel.grid(column=2, row=7,sticky="nw", padx=15)
        self.colRLabel['text'] = 'Right Matching Column'

        self.colselectorbox_1 = ttk.Combobox(self, textvariable=self.selectedcol1, width=30)
        self.colselectorbox_1['values'] = self.colOptions_1
        self.colselectorbox_1.state(["readonly"])
        self.colselectorbox_1.bind('<<ComboboxSelected>>', self.checkToEnableRun)
        self.colselectorbox_1.grid(row=8, column=0, padx=15)

        self.colselectorbox_2 = ttk.Combobox(self, textvariable=self.selectedcol2, width=30)
        self.colselectorbox_2['values'] = self.colOptions_2
        self.colselectorbox_2.state(["readonly"])
        self.colselectorbox_2.bind('<<ComboboxSelected>>', self.checkToEnableRun)
        self.colselectorbox_2.grid(row=8, column=2, padx=15)

        self.sliderLabel = tk.Label(self)
        self.sliderLabel.configure(background=self.backgroundcolour)
        self.sliderLabel.grid(column=1, row=5,sticky="nwe")
        self.sliderLabel['text'] = 'Similarity Cutoff'

        self.sliderLabel = tk.Label(self)
        self.sliderLabel.configure(background=self.backgroundcolour)
        self.sliderLabel.grid(column=1, row=7,sticky="nwe")
        self.sliderLabel['text'] = 'Maximum Matches to return'

        #self.matchCountVar.set(f'Rows Matched: {matchCount}')
        self.matchedCountLabel = tk.Label(self, textvariable=matchCountText)
        self.matchedCountLabel.configure(background=self.backgroundcolour, font='TkDefaultFont 10 bold')
        self.matchedCountLabel.grid(column=0, row=10,sticky="nw", padx=15)
        

        fuzzyHeader = tk.Label(self, text="Fuzzy Matching Options", font="Arial 12 bold")
        fuzzyHeader.configure(background=self.backgroundcolour)
        fuzzyHeader.grid(column=1, row=1, pady=10)

        # Decider widgets
        decideHeader = tk.Label(self, text="Self Deciding Match Options", font="Arial 12 bold")
        decideHeader.configure(background=self.backgroundcolour)
        decideHeader.grid(column=1, row=11, pady=10)

        self.runSDfuncsBox = tk.Checkbutton(self, text='Apply Smart Matching Decisions To Results', font="Arial 10 bold", variable=self.doSmartMatch, onvalue=True, offvalue=False, command=self.unlockSelfDecideWidgets, cursor='hand2')
        self.runSDfuncsBox.configure(background=self.backgroundcolour)
        self.runSDfuncsBox.grid(column=1, row=12, pady=5)

        self.sd_slider = {"Parameter 1":[], "Parameter 2":[], "Parameter 3":[]}
        self.sd_listBox = {"Parameter 1":[], "Parameter 2":[], "Parameter 3":[]}
        self.sd_compareEmpty = {"Parameter 1":[], "Parameter 2":[], "Parameter 3":[]}

        c=0
        for key in self.sd_slider:
            self.sd_slider[key].append(tk.Scale(self, orient=tk.HORIZONTAL, cursor='hand2'))
            self.sd_slider[key][0].grid(column=c, row=14)
            self.sd_slider[key][0].set(75)
            self.sd_slider[key].append(tk.Label(self, text=key))
            self.sd_slider[key][1].configure(bg=self.backgroundcolour)
            self.sd_slider[key][1].grid(column=c, row=13)
            c += 1
        c=0
        for key in self.sd_listBox:
            self.sd_listBox[key].append(tk.Listbox(self, width=30, height=5, selectmode="single", exportselection=0))
            self.sd_listBox[key][0].grid(column=c, row=16, rowspan=4, pady=5)
            self.sd_listBox[key].append(tk.Label(self, text=key, font=("Boulder", 13)))
            self.sd_listBox[key].append(tk.Listbox(self, width=30, height=5, selectmode="single", exportselection=0))
            self.sd_listBox[key][2].grid(column=c, row=21, rowspan=4, pady=5)
            c += 1
        c=0
        for key in self.sd_compareEmpty:
            self.sd_compareEmpty[key].append(tk.BooleanVar(value=False))
            self.sd_compareEmpty[key].append(tk.Checkbutton(self, text="Ignore match if both blank", variable=self.sd_compareEmpty[key][0], onvalue=True, offvalue=False, cursor='hand2'))
            self.sd_compareEmpty[key][1].grid(column=c, row=15)
            self.sd_compareEmpty[key][1].configure(background=self.backgroundcolour)
            c += 1

        self.changeWidgetCollectionState(self.sd_listBox, 'disabled')
        self.changeWidgetCollectionState(self.sd_slider, 'disabled')
        self.changeWidgetCollectionState(self.sd_compareEmpty, 'disabled')

        self.info = tk.Text(self, height=5)
        self.info.configure(bg=self.backgroundcolour)
        self.info.grid(column=0, row=26, pady=10, padx=15, columnspan=3, sticky='ew')
        sys.stdout = TextRedirector(self.info, "stdout")
        self.info['state'] = 'disabled'
        self.info.see('end')
        

        self.configure(background=self.backgroundcolour)

        self.grid_columnconfigure(0, pad=10)
        self.grid_columnconfigure(1, weight=3)


# Run the app mainloop
if __name__ == '__main__':
    app = TheSuperMatcher()
    app.mainloop()    