# -*- coding: utf-8 -*-
"""
Created on Thu Jun 29 07:50:31 2023

@author: lee.sheetrit
"""

import customtkinter
import tkinter as tk
import pandas as pd
from tkinter import ttk

class FieldMappingToDBFrame(customtkinter.CTkScrollableFrame):
    
    def __init__(self, master,**kwargs):
        super().__init__(master, **kwargs)
        
        print('ValueMappingFrame class intiated')
        #Attributes
     
        #Call Methods
        self.generate_widgets()
       
        
    def generate_widgets (self):
        
        label = customtkinter.CTkLabel(self, text="Enter Broker or Seller's Name")
        label.grid(row=1, column=0)
        
        self.seller_entry = customtkinter.CTkEntry(self, placeholder_text="seller1")
        self.seller_entry.grid(row=2, column=0)
   
        
        label = customtkinter.CTkLabel(self, text="Enter Mappings Name")
        label.grid(row=3, column=0)
        
        self.mapping_name_entry = customtkinter.CTkEntry(self, placeholder_text = 'newMapping1')
        self.mapping_name_entry.grid(row=4, column=0)
     
             
    def get_seller_name(self):
        return self.seller_entry.get()
    
    def get_mapping_name(self):
        return self.mapping_name_entry.get()
            
        #button
        # button = tk.Button(self, text="Write Value Mapping", command=write_value_mappings)
        # button.grid(row=row_num, column=0, columnspan=2)
