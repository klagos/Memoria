# Necessary imports
from tkinter import *
from tkinter import ttk

# Colors
colorActiveTab="#CCCCCC" # Color of the active tab
colorNoActiveTab="#EBEBEB" # Color of the no active tab
# Fonts
fontLabels='Calibri'
sizeLabels2=13

class ListboxEditable(object):
    """A class that emulates a listbox, but you can also edit a field"""
    # Constructor
    def __init__(self,frameMaster,list):
        # *** Assign the first variables ***
        # The frame that contains the ListboxEditable
        self.frameMaster=frameMaster
        # List of the initial items
        self.list=list
        # Number of initial rows at the moment
        self.numberRows=len(self.list)

        # *** Create the necessary labels ***
        ind=1
        for row in self.list:
            # Get the name of the label
            labelName='label'+str(ind)
            # Create the variable
            setattr(self, labelName, Label(self.frameMaster, text=self.list[ind-1], bg=colorActiveTab, fg='black', font=(fontLabels, sizeLabels2), pady=2, padx=2, width=10))

            # ** Bind actions
            # 1 left click - Change background
            getattr(self, labelName).bind('<Button-1>',lambda event, a=labelName: self.changeBackground(a))
            # Double click - Convert to entry
            getattr(self, labelName).bind('<Double-1>',lambda event, a=ind: self.changeToEntry(a))
            # Move up and down
            getattr(self, labelName).bind("<Up>",lambda event, a=ind: self.up(a))
            getattr(self, labelName).bind("<Down>",lambda event, a=ind: self.down(a))

            # Increase the iterator
            ind=ind+1

    # Place
    def placeListBoxEditable(self):
        # Go row by row placing it
        ind=1
        for row in self.list:
            # Get the name of the label
            labelName='label'+str(ind)
            # Place the variable
            getattr(self, labelName).grid(row=ind-1,column=0)

            # Increase the iterator
            ind=ind+1


    # Action to do when one click
    def changeBackground(self,labelNameSelected):
        # Ensure that all the remaining labels are deselected
        ind=1
        for row in self.list:
            # Get the name of the label
            labelName='label'+str(ind)
            # Place the variable
            getattr(self, labelName).configure(bg=colorActiveTab)

            # Increase the iterator
            ind=ind+1

        # Change the background of the corresponding label
        getattr(self, labelNameSelected).configure(bg=colorNoActiveTab)
        # Set the focus for future bindings (moves)
        getattr(self, labelNameSelected).focus_set()


    # Function to do when up button pressed
    def up(self, ind):
        if ind==1: # Go to the last
            # Get the name of the label
            labelName='label'+str(self.numberRows)
        else: # Normal
            # Get the name of the label
            labelName='label'+str(ind-1)

        # Call the select
        self.changeBackground(labelName)


    # Function to do when down button pressed
    def down(self, ind):
        if ind==self.numberRows: # Go to the last
            # Get the name of the label
            labelName='label1'
        else: # Normal
            # Get the name of the label
            labelName='label'+str(ind+1)

        # Call the select
        self.changeBackground(labelName)


    # Action to do when double-click
    def changeToEntry(self,ind):
        # Variable of the current entry
        self.entryVar=StringVar()
        # Create the entry
        #entryName='entry'+str(ind) # Name
        self.entryActive=ttk.Entry(self.frameMaster, font=(fontLabels, sizeLabels2), textvariable=self.entryVar, width=10)
        # Place it on the correct grid position
        self.entryActive.grid(row=ind-1,column=0)
        # Focus to the entry
        self.entryActive.focus_set()

        # Bind the action of focusOut
        self.entryActive.bind("<FocusOut>",lambda event, a=ind: self.saveEntryValue(a))

    
    # Action to do when focus out from the entry
    def saveEntryValue(self,ind):
        # Find the label to recover
        labelName='label'+str(ind)
        # Remove the entry from the screen
        self.entryActive.grid_forget()
        # Place it again
        getattr(self, labelName).grid(row=ind-1,column=0)
        # Change the name to the value of the entry
        getattr(self, labelName).configure(text=self.entryVar.get())
