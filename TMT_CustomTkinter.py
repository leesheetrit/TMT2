#https://customtkinter.tomschimansky.com/documentation/
#https://github.com/tomschimansky/customtkinter


#approach:
    #last working file: TMT_customTkinter before file split
    #get field mapping window loaded, no button functionality
    #move to seperate files <- here, getting into name error
    #keep building
    #build functionality for base tool as needed

import customtkinter
import tkinter as tk
import openpyxl as xl
from tkinter import ttk
from FieldMappingFrame import FieldMappingFrame

class NavigationFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        widgets = []
        #Lavel
        widgets.append(customtkinter.CTkLabel(self, text="Navigation Window"))
        #Buttons
        widgets.append(customtkinter.CTkButton(self, text="Read Excel TMT File", \
                                               command=self.read_excel_TMT))
        widgets.append(customtkinter.CTkButton(self, text="Map Tape Fields", \
                                               command=self.open_field_mapping_frame))
        widgets.append(customtkinter.CTkButton(self, text="Map Tape Values"))
        
        for i, widget in enumerate(widgets):
            widget.grid(row=i, column=0, padx=10, pady=5, sticky="ew")

    def read_excel_TMT(self):
        print('Reading Excel TMT...')
        #call loading screen class
        print('Excel TMT Read')
        #call excel loaded class
        
        print('Loading Excel File...')
        self.wb_path = 'I:/ServicingMSR/QRM/Trades Upload Files/Tape Mapping Tool/Tape Mapping Tool.xlsx'
        self.wb = xl.load_workbook(self.wb_path)
        self.dest_columns_sheet = self.wb['Destination Columns']
        self.tape_sheet = self.wb['Tape']
        self.dest_field_values_sheet = self.wb['Destination Field Values']
        self.values_mapping_sheet = self.wb['ValuesMapping']
        #Find the last used row on the Excel sheet
        print('Finding Last Row on Tape File...')
        self.last_row = self.tape_sheet.max_row
        self.source_columns = {}
        self.destination_columns = {}
        
        ####Read Source Columns 
        # Create a dictionary to store the source columns
        #stores the source column letter as the Dictionary item and the source Column Name as the dictionary value

        for cell in self.tape_sheet[1]:
            if cell.value is not None:
                self.source_columns[cell.column_letter] = cell.value
        print('Source Columns Read.')          
        # Sort the source column names alphabetically
        self.source_column_names = sorted(map(str, self.source_columns.values()))
        print('Source Columns Sorted.')          

        ####Read Destination Columns 
        #Read the destination columns into a dictionary called "destination_columns"
        #dictionarys are Key-Value pairs
        #stores the destination column name as the Dictionary Key
        for row in self.dest_columns_sheet.iter_rows(min_row=2, min_col=0, max_col=4):
            dest_col_name = row[0].value
            requires_mapping = row[1].value
            dest_value_example = row[2].value
            is_percentage = row[3].value
            
            #the value for the destination_columns dictionary is another dictionary
            #the key-value pairs of the child dictioary as shown below
            if dest_col_name is not None:
                self.destination_columns[dest_col_name] = {
                   'requires_mapping': requires_mapping,
                   'dest_value_example': dest_value_example,
                   'is_percentage': is_percentage,
                   }
        
    def open_field_mapping_frame(self):
        print('OpenFieldMappingFrame Method Activated')
        self.FieldMappingFrame= FieldMappingFrame(master=app)
        self.FieldMappingFrame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        
class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Tape Mapping Tool")
        self.geometry("700x700x5x5")
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(1,weight = 1) 
        self.navigation_frame  = NavigationFrame(master=self)
        self.navigation_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
      

if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")
    app = App()
    app.mainloop()


