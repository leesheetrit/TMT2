



if self.smart_match:
    for past_destination, past_source in self.past_mappings:
        if past_destination == item and past_source in self.source_columns:
            self.combobox_var.set(past_source)
            break