# -*- coding: utf-8 -*-
"""
Created on Mon Jun 19 19:29:15 2023

@author: lee.sheetrit
"""

import customtkinter
import tkinter as tk
import pandas as pd
from tkinter import ttk

class ValueMappingFrame(customtkinter.CTkScrollableFrame):
    
    def __init__(self, master, last_row, tape_sheet,\
                 column_mapping_rules,dest_field_values_sheet,\
                 source_columns, wb_path,**kwargs):
        super().__init__(master, **kwargs)
        
        print('FieldMappingFrame class intiated')
        #Attributes
        self.last_row = last_row
        self.tape_sheet = tape_sheet
        self.column_mapping_rules = column_mapping_rules
        self.dest_field_values_sheet = dest_field_values_sheet
        self.source_columns = source_columns
        self.wb_path = wb_path
        self.row_num = 0
        self.value_mapping_rules = []
        
        #Call Methods
        self.find_last_row()
        self.generate_widgets()
       
        
        
    def find_last_row(self):      
        #set the max row for source field values
        #checks if the column mappings have already been written, then limits the max row to the end of the tape (ignors the new mappinss)
        print('Checking if Column Mappings have been Written...')
        self.cell_value = self.tape_sheet["A{}".format(self.last_row-1)].value
        if self.cell_value == 'LOAN NUMBER' :
              self.last_row_adjusted = self.last_row - 4
        else:
              self.last_row_adjusted = self.last_row
              
        print('Last Row for Source Values is ',self.last_row_adjusted)
        
    def read_unique_values_from_excel(self,filename, sheet_name, column_name, max_row, max_values=15):
        #### functions for Values Mapping Window  read unique source column values 
        df = pd.read_excel(filename, sheet_name=sheet_name, nrows=max_row)
        column_values = df[column_name].unique()  # Get unique values from the specified column
        capped_values = column_values[:max_values]  # Cap the values at the specified maximum   
        
        return capped_values.tolist()  # Return the capped values as a list
    
    
    def generate_widgets (self):
        # Loop through destination_columns and create labels for columns that requires value mapping
        #iterate through the destination columns list
        for i, (dest_field_name, var, requires_mapping,is_percentage,hard_code) in enumerate(self.column_mapping_rules):
         
         if requires_mapping == 'TRUE': 
             #create a label for each destination column
             self.dest_col_label = customtkinter.CTkLabel(self, text="Dest: " + dest_field_name)
             self.dest_col_label.grid(row=self.row_num, column=1)
            
             # List to store destination field values
             self.dest_field_values = []
            
             # Iterate through each row in the destination field values sheet
             print('Reading Destination Field Values for ', dest_field_name)
             for row in self.dest_field_values_sheet.iter_rows(min_row=2, values_only=True):  
                 xl_dest_field_name = row[0]  # Column A
                 dest_field_value1 = row[1]  # Column B
            
                 if xl_dest_field_name == dest_field_name:
                     self.dest_field_values.append(dest_field_value1)
            
             # Print the list of matching destination field values
             print('Identified Destination values: ', self.dest_field_values)
            
            
             print('Finding Matching Source Column...')
            #find the matching source columns
             for source_col_letter, source_col_name in self.source_columns.items():
                 if source_col_name == var.get():
                     print('Source Column Found')
                     #create a label for matchin source column
                     source_col_label = customtkinter.CTkLabel(self, text="Source: " + source_col_name)
                     source_col_label.grid(row=self.row_num, column=0)
                    
                     #create a unique list of source values, capped at 15 values
                     ####call read_unique_values_from_excel
                     print('Reading Source Values for ', source_col_name)
                     unique_source_values = self.read_unique_values_from_excel(filename= self.wb_path, sheet_name = 'Tape', column_name = source_col_name, max_row = self.last_row_adjusted,  max_values = 15)
                     print('Identifiend Source Values: ',unique_source_values)
                    
                     print('Creating Labels for ', source_col_name)
                     #create a label for each unique destination value
                     for unique_source_value1 in unique_source_values:
                         self.row_num += 1
                          # Create a label for each unique destination value
                         self.label = customtkinter.CTkLabel(self, text=unique_source_value1)
                         self.label.grid(row=self.row_num, column=0)
                        
                         # var = tk.StringVar(value_mapping_window)
                         var = customtkinter.StringVar()
                         # Create a combo box for each unique destination value, populate it with the source values
                         combo = customtkinter.CTkComboBox(self, variable=var, \
                                                                values=list(self.dest_field_values), width=400, state ="readonly",height = 20)
                         combo.grid(row=self.row_num, column=1)

                        
                         for dest_field_value1 in self.dest_field_values:
                             #error handling for float
                            if isinstance(unique_source_value1, float) or isinstance(dest_field_value1, float):
                                if unique_source_value1 == dest_field_value1:
                                    combo.set(dest_field_value1)
                                    break
                            else:
                                #case insensitive match
                                if unique_source_value1.lower() == dest_field_value1.lower():
                                    combo.set(dest_field_value1)
                                    break
          
                          # #add destination column, the selected variable, and whether or not the column requires mapping to the "column_mapping_rules" list
                         self.value_mapping_rules.append((source_col_letter,source_col_name,unique_source_value1,dest_field_name,var))
                         
                     break
          
             self.row_num += 1
             
    def get_field_mappings(self):
        return self.value_mapping_rules
            
        #button
        # button = tk.Button(self, text="Write Value Mapping", command=write_value_mappings)
        # button.grid(row=row_num, column=0, columnspan=2)
