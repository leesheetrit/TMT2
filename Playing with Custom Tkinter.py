
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
        self.query = "SELECT Destination_Field, Source_Field FROM Servicing.dbo.TMT_Field_Mappings"
        self.query_result = self.qrm_db.connect().execute(self.query)
        self.past_mappings = [(row['Destination_Field'], row['Source_Field']) for row in self.query_result]
        
        print("Past Mappings", self.past_mappings)

        



if self.smart_match:
    for past_destination, past_source in self.past_mappings:
        if past_destination == item and past_source in self.source_columns:
            self.combobox_var.set(past_source)
            break