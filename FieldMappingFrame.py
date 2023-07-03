# -*- coding: utf-8 -*-
"""
Created on Sun Jun 18 20:48:10 2023

@author: lee.sheetrit
"""
import customtkinter
import tkinter as tk
import datetime
from tkinter import ttk
import sqlalchemy


class FieldMappingFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, destination_columns, source_columns, smart_match = False, qrm_db = '',**kwargs):
        super().__init__(master, **kwargs)
        
        print('FieldMappingFrame class intiated')
        #Attributes
        self.column_mapping_rules = []
        self.DivideBy100Vars = []
        self.AllCurrentVar = None 
        self.destination_columns = destination_columns
        # print(destination_columns)
        self.source_columns = source_columns
        print(self.source_columns)
        self.column_mapping_rules = []
        self.smart_match = smart_match
        self.qrm_db = qrm_db
        
        #Call Methods
        if self.smart_match:
            self.query_past_mappings()
            
        self.generate_headers()
        self.generate_mapping_widgets()

    def query_past_mappings(self):
        self.query = "SELECT Destination_Field, Source_Field FROM Servicing.dbo.TMT_Field_Mappings"
        self.query_result = self.qrm_db.connect().execute(self.query)
        self.past_mappings = [(row['Destination_Field'], row['Source_Field']) for row in self.query_result]
        
        print("Past Mappings", self.past_mappings)

        
        
    def generate_headers(self):
        # add widgets onto the frame, for example:
        labels_data = [
        ("Destination Column", 0),
        ("Requires Mapping", 1),
        ("One Loan Example", 2),
        ("Source Column", 3),
        ("Special", 4),
        ('Hard Code Value', 5)]
        
        for text, column in labels_data:
            label = customtkinter.CTkLabel(self, text=text, font=customtkinter.CTkFont(size=20, weight="bold"))
            label.grid(row=0, column=column, padx=20)
            

    def generate_mapping_widgets(self):
        
        row_num = 2
        #method to be: generate mapping rows
        for item, child_dictionary in self.destination_columns.items():
            # Create a label and combo box for each Destination Column in the "destination_columns" 
                 
            self.label = customtkinter.CTkLabel(self, text=item)
            self.label.grid(row=row_num, column=0)
            
            RequiresMappingLabelText = child_dictionary['requires_mapping']\
                if child_dictionary['requires_mapping'] != 'FALSE' else ''
            self.label2 = customtkinter.CTkLabel(self, text=RequiresMappingLabelText)
            self.label2.grid(row=row_num, column=1)
            
            self.label3 = customtkinter.CTkLabel(self, text=child_dictionary['dest_value_example'])
            self.label3.grid(row=row_num, column=2)
            
            self.combobox_var = customtkinter.StringVar()
            #not sure if I want to use this or not
            # self.combo = customtkinter.CTkComboBox(self, variable=self.combobox_var, \
                                                   # values=list(self.source_columns), state="readonly", width = 400)
            self.combo = ttk.Combobox(self, textvariable=self.combobox_var, \
                                                       values=list(self.source_columns), state="readonly", width = 40, height = 20)
            # self.combo.configure(postcommand=self.limit_combobox_height)
            self.combo.grid(row=row_num, column=3)
            
            if self.smart_match:
                for past_destination, past_source in self.past_mappings:
                   if past_destination == item and past_source in self.source_columns:
                       self.combobox_var.set(past_source)
                       break
            
            # self.combo.bind('<<ComboboxSelected>>', combobox_selected)
            self.combo.focus_set()
            self.combo.bind('<KeyPress>', self.keyboard_first_letter_select)
     
            # combo.bind('<KeyPress>', keyboard_first_letter_select)
               
            # Check if any of the source column names match "item" and set it as the default value
            #comment this out if you don't want the default matching
            for source_column_name in self.source_columns:
                if source_column_name.lower() == item.lower():
                    self.combobox_var.set(source_column_name)
                    break
            
            
            if child_dictionary['is_percentage'] == True:
                # Create a variable to hold the state of the checkbox
                self.DivideBy100Var = tk.BooleanVar()
                # Create the checkbox
                self.checkbox = customtkinter.CTkCheckBox(self, text='Divide by 100?', variable=self.DivideBy100Var)
                self.checkbox.grid(row=row_num, column=4)
                self.DivideBy100Vars.append(self.DivideBy100Var) # Add the DivideBy100Var to the DivideBy100Vars list
            elif item == 'Delq':
                self.AllCurrentVar = tk.BooleanVar()
                checkbox = customtkinter.CTkCheckBox(self, text='All Current?', variable=self.AllCurrentVar)
                checkbox.grid(row=row_num, column=4)
                self.DivideBy100Vars.append(False)
            else:
                self.DivideBy100Vars.append(False)
        
            self.hard_code_var = customtkinter.StringVar()
            self.hard_code = customtkinter.CTkEntry(self, textvariable = self.hard_code_var)
            self.hard_code.grid(row=row_num, column=5)
        
            #add destination column, the selected variable, and whether or not the column requires mapping to the "column_mapping_rules" list
            self.column_mapping_rules.append((item, self.combobox_var, \
                                              child_dictionary['requires_mapping'],\
                                                  child_dictionary['is_percentage'],\
                                                      self.hard_code)) #adds item, var, and requires mapping as as tupple
        
            row_num += 1
            
           
        print('Labels created')
        
        print('Field Mapping Frame Loaded')
    def keyboard_first_letter_select(self, event):
        # print ('in keyboard_first_letter_select method')
        combo = event.widget

        # Get the typed letter from the event
        letter = event.char.lower()

        # Find the first list item that starts with the typed letter
        for i, item in enumerate(combo['values']):
            if item.lower().startswith(letter):
                combo.current(i)  # Select the found item
                break
        # print ('letter: ', letter)
        
        
    def get_field_mappings(self):
        return self.column_mapping_rules
    
    def get_DivideBy100Vars(self):
        return self.DivideBy100Vars

    def get_AllCurrentVar(self):
        return self.AllCurrentVar
    

        
    
if __name__ == "__main__":
    print('in test mode')
    
    #generate tk window
    #call the frame sending it a static destination columns and source coulmns 
    # customtkinter.set_appearance_mode("dark")
    # app = App()
    # app.mainloop()
    
    # destination_columns = {
    # 'LOAN NUMBER': {'requires_mapping': 'FALSE', 'dest_value_example': 1464829561, 'is_percentage': 'FALSE'},
    # 'ORIGINAL MORTGAGE AMOUNT': {'requires_mapping': 'FALSE', 'dest_value_example': 139425, 'is_percentage': 'FALSE'},
    # 'INVESTOR HEADER': {'requires_mapping': 'TRUE', 'dest_value_example': 'GNMA', 'is_percentage': 'FALSE'},
    # 'REMIT TYPE': {'requires_mapping': 'TRUE', 'dest_value_example': 'GNMA II', 'is_percentage': 'FALSE'},
    # 'LO TYPE DESCRIPTION': {'requires_mapping': 'TRUE', 'dest_value_example': 'FHA RESIDENTIAL', 'is_percentage': 'FALSE'},
    # 'FIRST PRINCIPAL BALANCE': {'requires_mapping': 'FALSE', 'dest_value_example': 12.99, 'is_percentage': 'FALSE'},
    # 'T AND I MONTHLY AMOUNT': {'requires_mapping': 'FALSE', 'dest_value_example': 231.5, 'is_percentage': 'FALSE'},
    # 'ESCROWED INDICATOR': {'requires_mapping': 'TRUE', 'dest_value_example': 'Y', 'is_percentage': 'FALSE'},
    # 'LOAN TERM': {'requires_mapping': 'FALSE', 'dest_value_example': 360, 'is_percentage': 'FALSE'},
    # 'ARM PLAN ID': {'requires_mapping': 'TRUE', 'dest_value_example': 'SOF5', 'is_percentage': 'FALSE'},
    # 'FIRST DUE DATE': {'requires_mapping': 'FALSE', 'dest_value_example': datetime.datetime(2020, 9, 1, 0, 0), 'is_percentage': 'FALSE'},
    # 'INTEREST PAID THROUGH DATE': {'requires_mapping': 'FALSE', 'dest_value_example': datetime.datetime(2023, 2, 1, 0, 0), 'is_percentage': 'FALSE'},
    # 'LOAN MATURES DATE': {'requires_mapping': 'FALSE', 'dest_value_example': datetime.datetime(2050, 8, 1, 0, 0), 'is_percentage': 'FALSE'},
    # 'LOAN CLOSING DATE': {'requires_mapping': 'FALSE', 'dest_value_example': datetime.datetime(2020, 7, 29, 0, 0), 'is_percentage': 'FALSE'},
    # 'ANNUAL INTEREST RATE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.03, 'is_percentage': True},
    # 'NET SERV FEE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.005, 'is_percentage': True},
    # 'PROPERTY ALPHA STATE CODE': {'requires_mapping': 'FALSE', 'dest_value_example': 'TN', 'is_percentage': 'FALSE'},
    # 'PROPERTY TYPE FNMA CODE DESCRIPTION': {'requires_mapping': 'TRUE', 'dest_value_example': 'SINGLE FAMILY DETACHED', 'is_percentage': 'FALSE'},
    # 'ORIG LOAN TO VALUE RATIO': {'requires_mapping': 'FALSE', 'dest_value_example': 0.982, 'is_percentage': True},
    # 'LOAN TO VALUE RATIO': {'requires_mapping': 'FALSE', 'dest_value_example': 0, 'is_percentage': True},
    # 'LOAN PURPOSE CODE DESCRIPTION': {'requires_mapping': 'TRUE', 'dest_value_example': 'PURCHASE', 'is_percentage': 'FALSE'},
    # 'OCCUPANCY CODE DESCRIPTION': {'requires_mapping': 'TRUE','dest_value_example': 'OWNER OCCUPANCY', 'is_percentage': 'FALSE'},
    # 'CS BORR CREDIT QUALITY CODE': {'requires_mapping': 'FALSE', 'dest_value_example': '635', 'is_percentage': 'FALSE'},
    # 'PREPAY PENALTY INDICATOR': {'requires_mapping': 'TRUE', 'dest_value_example': 'N', 'is_percentage': 'FALSE'},
    # 'FORECLOSURE STOP CODE DESCRIPTION': {'requires_mapping': 'TRUE', 'dest_value_example': 'NORMAL PROCESSING', 'is_percentage': 'FALSE'},
    # 'ARM IR MARGIN RATE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.0275, 'is_percentage': True},
    # 'ARM IR MAX LIFE FLOOR RATE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.0275, 'is_percentage': True},
    # 'ARM IR MAX LIFE CEILING RATE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.09125, 'is_percentage': True},
    # 'ARM INDEX CODE 1 DESCRIPTION': {'requires_mapping': 'TRUE', 'dest_value_example': '30 DAY AVERAGE SOFR', 'is_percentage': 'FALSE'},
    # 'ARM IR MAX INCREASE RATE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.02, 'is_percentage': True},
    # 'ARM IR MAX DECREASE RATE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.01375, 'is_percentage': True},
    # 'ARM IR CHANGE PERIOD': {'requires_mapping': 'FALSE', 'dest_value_example': 6, 'is_percentage': 'FALSE'},
    # 'ARM NEXT PI CALC DATE': {'requires_mapping': 'FALSE', 'dest_value_example': datetime.datetime(2027, 11, 1, 0, 0), 'is_percentage': 'FALSE'},
    # 'MP GUARANTY FEE RATE': {'requires_mapping': 'FALSE', 'dest_value_example': 0.06, 'is_percentage': 'FALSE'},
    # 'ACTIVE PANDEMIC FORBEARANCE': {'requires_mapping': 'TRUE', 'dest_value_example': 0, 'is_percentage': 'FALSE'},
    # 'Delq': {'requires_mapping': 'TRUE', 'dest_value_example': 'PREPAID OR CURRENT', 'is_percentage': 'FALSE'},
    # 'DTI': {'requires_mapping': 'FALSE', 'dest_value_example': 0.151466268, 'is_percentage': True},
    # 'Division': {'requires_mapping': 'TRUE', 'dest_value_example': 'RETAIL', 'is_percentage': 'FALSE'},
    # 'Refi Type': {'requires_mapping': 'TRUE', 'dest_value_example': 0, 'is_percentage': 'FALSE'}}
    
    


