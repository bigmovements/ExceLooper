# -*- coding: UTF-8 -*-


import xlwings as xw, tkinter, functools
from tkinter import messagebox


def quitGUI(master):
    """Closes GUI, terminates processing"""
    master.destroy()


def checkInputVals(minVal, maxVal, stepVal):
    try:
        minVal = float(minVal)
        maxVal = float(maxVal)
        stepVal = float(stepVal)
    except:
        tkinter.messagebox.showerror(
            "Input Error", "One or more input values are not numbers")
        return False
    else:
        return True


def radioButtonSet1Assumption(assumption2OptionMenu, assumption2CellEntry,
                              assumption2MinValEntry, assumption2MaxValEntry,
                              assumption2StepValEntry):
    assumption2OptionMenu.config(state = "disabled")
    assumption2CellEntry.config(state = "disabled")
    assumption2MinValEntry.config(state = "disabled")
    assumption2MaxValEntry.config(state = "disabled")
    assumption2StepValEntry.config(state = "disabled")

    
def radioButtonSet2Assumption(assumption2OptionMenu, assumption2CellEntry,
                              assumption2MinValEntry, assumption2MaxValEntry,
                              assumption2StepValEntry):
    assumption2OptionMenu.config(state = "normal")
    assumption2CellEntry.config(state = "normal")
    assumption2MinValEntry.config(state = "normal")
    assumption2MaxValEntry.config(state = "normal")
    assumption2StepValEntry.config(state = "normal")


def refresh(fileInputEntry, assumption1OptionMenu, assumption2OptionMenu,
            outputOptionMenu, assumption1SheetName, assumption2SheetName,
            outputSheetName):
    
    # Checking if Excel is open
    if len(xw.apps) == 0:
        tkinter.messagebox.showerror(
            "Error: Open Excel",
            "This application requires your file to be open in Excel")
        return
    
    # Checking user workbook is open
    elif '<Book [%s]>' % fileInputEntry.get() \
            not in [str(i) for i in xw.books]:
        tkinter.messagebox.showerror(
            "Error: Open Excel File",
            "This application requires your Excel file to be open")
        return

    # Refresh list of worksheets
    # First empty current lists, obtain worksheets, then refill list
    assumption1OptionMenu['menu'].delete(0, 'end')
    assumption2OptionMenu['menu'].delete(0, 'end')
    outputOptionMenu['menu'].delete(0, 'end')
    
    sheets = []
    wb = xw.Book(fileInputEntry.get())

    for i in range(len(wb.sheets)):
        sheets.append(str(wb.sheets[i].name))

    for choice in sheets:
        assumption1OptionMenu['menu'].add(
            label = choice,
            itemType = "radiobutton",
            variable = assumption1SheetName)
        assumption2OptionMenu['menu'].add(
            label = choice,
            itemType = "radiobutton",
            variable = assumption2SheetName)
        outputOptionMenu['menu'].add(
            label = choice,
            itemType = "radiobutton",
            variable = outputSheetName)

    assumption1SheetName.set(sheets[0])
    assumption2SheetName.set(sheets[0])
    outputSheetName.set(sheets[0])
        

def outputDataToWorkbookTwoAssumption(assumptionData1, assumptionData2,
                                      outputData, wb):
    # Creating a new sheet in workbook and inserting output data
    wb.sheets.add()
    wsDataOutput = wb.sheets.active

    for i in range(1, len(assumptionData1) + 1):
        wsDataOutput.range((1 + i, 1)).value = assumptionData1[i - 1]

    for i in range(1, len(assumptionData2) + 1):
        wsDataOutput.range((1, 1 + i)).value = assumptionData2[i - 1]


    for i in range(2, len(outputData) + 2):
        for j in range(2, len(outputData[0]) + 2):
            wsDataOutput.range((i, j)).value = outputData[i - 2][j - 2]


def varyAndReadTwoAssumption(minVal1, maxVal1, step1,
                             minVal2, maxVal2, step2,
                             wsAssumptions1, wsAssumptions2,
                             wsOutputs, assumptionCell1,
                             assumptionCell2, outputCell):
    # Varying assumptions between given values. Output data stored in 2-D array.
    
    assumptionData1 = []
    assumptionData2 = []
    outputData = []
    tempOutputData = []
    index1 = minVal1
    index2 = minVal2

    # No single pass through of second assumption values so done outside of algorithm
    # Also assigning first assumption values now because why not
    
    while index1 <= maxVal1:
        assumptionData1.append(index1)
        index1 += step1

    while index2 <= maxVal2:
        assumptionData2.append(index2)
        index2 += step2

    # Algorithm looping through all values contained within user inputted range

    index1 = minVal1
    index2 = minVal2
    while index1 <= maxVal1:
        wsAssumptions1.range(assumptionCell1).value = index1
        index2 = minVal2
        
        while index2 <= maxVal2:
            wsAssumptions2.range(assumptionCell2).value = index2
            tempOutputData.append(wsOutputs.range(outputCell).value)
            index2 += step2

        outputData.append(tempOutputData)
        tempOutputData = []
        index1 += step1

    
    return assumptionData1, assumptionData2, outputData


def outputDataToWorkbookOneAssumption(assumptionData, outputData, wb):
    # Creating a new sheet in workbook and inserting output data
    
    wb.sheets.add()
    wsDataOutput = wb.sheets.active
    
    for i in range(1, len(assumptionData) + 1):
        wsDataOutput.range((i, 1)).value = assumptionData[i - 1]
        wsDataOutput.range((i, 2)).value = outputData[i - 1]

    
def varyAndReadOneAssumption(minVal, maxVal, step,
                             wsAssumptions, wsOutputs,
                             assumptionCell, outputCell):
    # Varying assumption between given values and storing assumption input values
    # and cost outputs in seperate arrays
    
    assumptionData = []
    outputData = []
    
    index = minVal
    while index <= maxVal:
        assumptionData.append(index)
        index += step

    index = minVal
    while index <= maxVal:
        wsAssumptions.range(assumptionCell).value = index
        outputData.append(wsOutputs.range(outputCell).value)
        index += step

    return assumptionData, outputData


def initiateProcessing(assumption1OptionMenu, assumption1CellEntry,
                       assumption1SheetName, assumption1MinValEntry,
                       assumption1MaxValEntry, assumption1StepValEntry,
                       assumption2OptionMenu, assumption2CellEntry,
                       assumption2SheetName, assumption2MinValEntry,
                       assumption2MaxValEntry, assumption2StepValEntry,
                       outputOptionMenu, outputCellEntry, outputSheetName,
                       radioButtonVar, fileInputEntry):
    # Checking if Excel is open
    if len(xw.apps) == 0:
        tkinter.messagebox.showerror(
            "Error: Open Excel",
            "This application requires your file to be open in Excel")
        return

    # Checking user workbook is open
    elif '<Book [%s]>' % fileInputEntry.get() \
            not in [str(i) for i in xw.books]:
        tkinter.messagebox.showerror(
            "Error: Open Excel File",
            "This application requires your Excel file to be open")
        return

    # Assigning assumption 1 variables using get() method, minimising memory usage
    wb = xw.Book(fileInputEntry.get())
    
    wbAssumption1 = wb.sheets[assumption1SheetName.get()]
    wbOutput = wb.sheets[outputSheetName.get()]
    
    assum1Cell = assumption1CellEntry.get()
    outputCell = outputCellEntry.get()

    assum1MinVal = assumption1MinValEntry.get()
    assum1MaxVal = assumption1MaxValEntry.get()
    assum1StepVal = assumption1StepValEntry.get()
        
    if checkInputVals(assum1MinVal, assum1MaxVal, assum1StepVal) == True:
        assum1MinVal = float(assum1MinVal)
        assum1MaxVal = float(assum1MaxVal)
        assum1StepVal = float(assum1StepVal)
    else:
        return

    # Using active radio button to vary one or two assumptions
    if radioButtonVar.get() == 1:
        assumptionData, outputData = varyAndReadOneAssumption(
            assum1MinVal, assum1MaxVal, assum1StepVal,
            wbAssumption1, wbOutput, assum1Cell, outputCell)
        outputDataToWorkbookOneAssumption(
            assumptionData, outputData, wb)
    else:
        # Assinging assumption 2 variables only if two are being varied
        wbAssumption2 = wb.sheets[assumption2SheetName.get()]
        assum2Cell = assumption2CellEntry.get()

        assum2MinVal = assumption2MinValEntry.get()
        assum2MaxVal = assumption2MaxValEntry.get()
        assum2StepVal = assumption2StepValEntry.get()
        
        if checkInputVals(assum2MinVal, assum2MaxVal, assum2StepVal) == True:
            assum2MinVal = float(assum2MinVal)
            assum2MaxVal = float(assum2MaxVal)
            assum2StepVal = float(assum2StepVal)
        else:
            return
        
        assumptionData1, assumptionData2, outputData = varyAndReadTwoAssumption(
            assum1MinVal, assum1MaxVal, assum1StepVal,
            assum2MinVal, assum2MaxVal, assum2StepVal,
            wbAssumption1, wbAssumption2, wbOutput,
            assum1Cell, assum2Cell, outputCell)
        outputDataToWorkbookTwoAssumption(
            assumptionData1, assumptionData2, outputData, wb)                   


def main():
    # Creating GUI
    master = tkinter.Tk()
    master.wm_title("ExceLooper")

    title = tkinter.Label(
        master,
        text = "ExceLooper",
        width = 40,
        font = ("Courier", 14))
    disclaimer = tkinter.Label(
        master,
        pady = 7,
        text = "Note: I recommend not adjusting the excel document whilst " \
               "processing.\nRefresh this application after each process or " \
               "if sheet arrangement has changed." \
               "\nCreated by George Watson for Loowatt Ltd.") 

    # Widget Creation
    
    # Frames
    inputFrame = tkinter.Frame(master)
    fileInputFrame = tkinter.Frame(inputFrame)
    cellInputFrame = tkinter.Frame(inputFrame)
    
    assumption1ValuesFrame = tkinter.Frame(cellInputFrame, pady = 10)
    assumption2ValuesFrame = tkinter.Frame(cellInputFrame, pady = 10)

    radioButtonFrame = tkinter.Frame(inputFrame)
    buttonFrame = tkinter.Frame(inputFrame)

    # Creating Button/Entry/Label Widgets
    
    # fileInputFrame
    fileInputLabel = tkinter.Label(fileInputFrame, text = "Excel File Name:")
    fileInputEntry = tkinter.Entry(fileInputFrame, width = 35)
    fileInputEntry.insert(0, "file name.xlsx")

    # Assumption 1
    assumption1SheetName = tkinter.StringVar(cellInputFrame)
    assumption1Label = tkinter.Label(cellInputFrame,
                                     text = "Assumption 1:")
    assumption1OptionMenu = tkinter.OptionMenu(cellInputFrame,
                                               assumption1SheetName,
                                               "Select Sheet")
    assumption1CellEntry = tkinter.Entry(cellInputFrame, width = 5)
    assumption1SheetName.set("Select Sheet")
    assumption1CellEntry.insert(0, "A1")

    # Assumption 1 Values
    assumption1MinValLabel = tkinter.Label(assumption1ValuesFrame,
                                           text = "Min Value:")
    assumption1MinValEntry = tkinter.Entry(assumption1ValuesFrame,
                                           width = 7)
    assumption1MaxValLabel = tkinter.Label(assumption1ValuesFrame,
                                           text = "Max Value:")
    assumption1MaxValEntry = tkinter.Entry(assumption1ValuesFrame,
                                           width = 7)
    assumption1StepValLabel = tkinter.Label(assumption1ValuesFrame,
                                            text = "Step Size:")
    assumption1StepValEntry = tkinter.Entry(assumption1ValuesFrame,
                                            width = 5)

    # Assumption 2
    assumption2SheetName = tkinter.StringVar(cellInputFrame)
    assumption2Label = tkinter.Label(cellInputFrame,
                                     text = "Assumption 2:")
    assumption2OptionMenu = tkinter.OptionMenu(cellInputFrame,
                                               assumption2SheetName,
                                               "Select Sheet")
    assumption2CellEntry = tkinter.Entry(cellInputFrame, width = 5)
    assumption2SheetName.set("Select Sheet")
    assumption2CellEntry.insert(0, "A1")

    # Assumption 2 Values
    assumption2MinValLabel = tkinter.Label(assumption2ValuesFrame,
                                           text = "Min Value:")
    assumption2MinValEntry = tkinter.Entry(assumption2ValuesFrame,
                                           width = 7)
    assumption2MaxValLabel = tkinter.Label(assumption2ValuesFrame,
                                           text = "Max Value:")
    assumption2MaxValEntry = tkinter.Entry(assumption2ValuesFrame,
                                           width = 7)
    assumption2StepValLabel = tkinter.Label(assumption2ValuesFrame,
                                            text = "Step Size:")
    assumption2StepValEntry = tkinter.Entry(assumption2ValuesFrame,
                                            width = 5)

    # Output
    outputSheetName = tkinter.StringVar(cellInputFrame)
    
    outputLabel = tkinter.Label(cellInputFrame, text = "Output:")
    outputOptionMenu = tkinter.OptionMenu(cellInputFrame,
                                          outputSheetName,
                                          "Select Sheet")
    outputCellEntry = tkinter.Entry(cellInputFrame, width = 5)

    outputSheetName.set("Select Sheet")
    outputCellEntry.insert(0, "A1")

    # Radio Buttons (1 or 2 Assumptions)
    radioButtonVar = tkinter.IntVar()
    radioButtonVar.set(2)
    radioButton1Assumption = tkinter.Radiobutton(
            radioButtonFrame,
            text = "1 Assumption", variable = radioButtonVar, value = 1,
            command = functools.partial(radioButtonSet1Assumption,
            assumption2OptionMenu, assumption2CellEntry,
            assumption2MinValEntry, assumption2MaxValEntry,
            assumption2StepValEntry),
            pady = 10)
    radioButton2Assumption = tkinter.Radiobutton(
            radioButtonFrame,
            text = "2 Assumptions",variable = radioButtonVar, value = 2,
            command = functools.partial(radioButtonSet2Assumption,
            assumption2OptionMenu, assumption2CellEntry,
            assumption2MinValEntry, assumption2MaxValEntry,
            assumption2StepValEntry),
            pady = 10)

    # Quit and Process buttons (put refresh in here too?)
    quitButton = tkinter.Button(
            buttonFrame, text = "Quit", width = 8, pady = 10,
            command = functools.partial(quitGUI, master))
    refreshButton = tkinter.Button(
            buttonFrame, text = "Refresh", width = 12, pady = 10,
            command = functools.partial(refresh, fileInputEntry,
            assumption1OptionMenu, assumption2OptionMenu,
            outputOptionMenu, assumption1SheetName,
            assumption2SheetName, outputSheetName))
    processButton = tkinter.Button(
            buttonFrame, text = "Process", width = 14, pady = 10,
            command = functools.partial(initiateProcessing,
            assumption1OptionMenu, assumption1CellEntry,
            assumption1SheetName, assumption1MinValEntry,
            assumption1MaxValEntry, assumption1StepValEntry,
            assumption2OptionMenu, assumption2CellEntry,
            assumption2SheetName, assumption2MinValEntry,
            assumption2MaxValEntry, assumption2StepValEntry,
            outputOptionMenu, outputCellEntry, outputSheetName,
            radioButtonVar, fileInputEntry))

    # Displaying widgets using grid geometry method
    title.grid(column = 0, row = 0)

    fileInputLabel.grid(column = 0, row = 0)
    fileInputEntry.grid(column = 1, row = 0)
    fileInputFrame.grid(column = 0, row = 0)

    radioButton1Assumption.grid(column = 0, row = 0)
    radioButton2Assumption.grid(column = 1, row = 0)

    radioButtonFrame.grid(column = 0, row = 1)

    assumption1Label.grid(column = 0, row = 2)
    assumption1OptionMenu.grid(column = 1, row = 2)
    assumption1CellEntry.grid(column = 2, row = 2)

    assumption1MinValLabel.grid(column = 0, row = 0)
    assumption1MinValEntry.grid(column = 1, row = 0)
    assumption1MaxValLabel.grid(column = 2, row = 0)
    assumption1MaxValEntry.grid(column = 3, row = 0)
    assumption1StepValLabel.grid(column = 4, row = 0)
    assumption1StepValEntry.grid(column = 5, row = 0)

    assumption1ValuesFrame.grid(column = 0, row = 3, columnspan = 3)

    assumption2Label.grid(column = 0, row = 4)
    assumption2OptionMenu.grid(column = 1, row = 4)
    assumption2CellEntry.grid(column = 2, row = 4)

    assumption2MinValLabel.grid(column = 0, row = 0)
    assumption2MinValEntry.grid(column = 1, row = 0)
    assumption2MaxValLabel.grid(column = 2, row = 0)
    assumption2MaxValEntry.grid(column = 3, row = 0)
    assumption2StepValLabel.grid(column = 4, row = 0)
    assumption2StepValEntry.grid(column = 5, row = 0)

    assumption2ValuesFrame.grid(column = 0, row = 5, columnspan = 3)

    outputLabel.grid(column = 0, row = 6)
    outputOptionMenu.grid(column = 1, row = 6)
    outputCellEntry.grid(column = 2, row = 6)

    cellInputFrame.grid(column = 0, row = 2)

    quitButton.grid(column = 0, row = 0)
    refreshButton.grid(column = 1, row = 0)
    processButton.grid(column = 2, row = 0)

    buttonFrame.grid(column = 0, row = 3)
    
    inputFrame.grid(column = 0, row = 1)

    disclaimer.grid(column = 0, row = 2)

    master.mainloop()

   
if __name__ == '__main__':
    main()

