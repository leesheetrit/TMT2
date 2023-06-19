# -*- coding: utf-8 -*-
"""
Created on Sun Jun 18 20:48:10 2023

@author: lee.sheetrit
"""
import customtkinter
import tkinter as tk



class FieldMappingFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        
        #Attributes
        self.column_mapping_rules = []
        self.DivideBy100Vars = []
        self.AllCurrentVar = None 
        
        #Call Methods
        self.generate_labels()
        self.generate_mapping_widgets()

    def generate_labels(self):
        # add widgets onto the frame, for example:
        labels_data = [
        ("Destination Column", 0),
        ("Requires Mapping", 1),
        ("One Loan Example", 2),
        ("Source Column", 3),
        ("Special", 4)]
        
        for text, column in labels_data:
            label = customtkinter.CTkLabel(self, text=text, font=customtkinter.CTkFont(size=20, weight="bold"))
            label.grid(row=0, column=column, padx=20)

     
    #this should be a class
    #with a label
    
    def generate_mapping_widgets(self):
        row_num = 2
        #method to be: generate mapping rows
        for item, child_dictionary in app.navigation_frame.destination_columns.items():
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
            self.combo = customtkinter.CTkComboBox(self, variable=self.combobox_var, values=list(app.navigation_frame.source_column_names), state="readonly")
            self.combo.grid(row=row_num, column=3)
         
            # combo.bind('<KeyPress>', keyboard_first_letter_select)
               
            # Check if any of the source column names match "item" and set it as the default value
            #comment this out if you don't want the default matching
            for source_column_name in app.navigation_frame.source_column_names:
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
        
            #add destination column, the selected variable, and whether or not the column requires mapping to the "column_mapping_rules" list
            self.column_mapping_rules.append((item, self.combobox_var, \
                                              child_dictionary['requires_mapping'],\
                                                  child_dictionary['is_percentage'])) #adds item, var, and requires mapping as as tupple
            
            row_num += 1
            
        print('Labels created')

        # # Create a button to trigger the write_selections function
        # # button = customtkinter.CTkButton(self, text="Implement Field Mapping Rules", command=write_field_mappings)
        # button = customtkinter.CTkButton(self, text="Implement Field Mapping Rules")
        # # button.grid(row=row_num, column=0, columnspan=2)
        # button.grid(row=row_num, column=0)

        
        # # button = customtkinter.CTkButton(self, text="Open Value Mapping Rules Window", command=open_value_mappings_window)
        # button = customtkinter.CTkButton(self, text="Open Value Mapping Rules Window")
        # button.grid(row=row_num, column=3, columnspan=2)
        
        ####Load Tape Mapping Tool Excel File
        # button = tk.Button(field_map_window, text="Save Mapping Rules", command=open_save_load_field_mappings_window)    
        # # button.grid(row=row_num, column=0, columnspan=2)
        # button.grid(row=row_num, column=1)
        