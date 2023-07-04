
class ValueMappingFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, ..., smart_match = False, qrm_db = '',**kwargs):
        super().__init__(master, **kwargs)
        
        #Attributes
     ...
        self.smart_match = smart_match
        self.qrm_db = qrm_db
        
        #Call Methods
        if self.smart_match:
            self.query_past_mappings()
            
        ...
    def query_past_mappings(self):
        self.query = "SELECT Destination_Value, Source_Value FROM Servicing.dbo.TMT_Value_Mappings"
        self.query_result = self.qrm_db.connect().execute(self.query)
        self.past_mappings = [(row['Destination_Value'], row['Source_Value']) for row in self.query_result]
        
        print("Past Mappings", self.past_mappings)

        

    def gen_widgets(self):
        ...
        if self.smart_match:
            for past_destination, past_source in self.past_mappings:
                if past_source == unique_source_value1 /
                and past_destination in self.dest_field_values: #change me!
                    self.combobox_var.set(past_destination)
                    break
                    
                    
