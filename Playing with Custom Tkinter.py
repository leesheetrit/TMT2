

import customtkinter


class NavigationFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # add widgets onto the frame, for example:
        self.label = customtkinter.CTkLabel(self, text="Navigation Window")
        self.label.grid(row=0, column=0, padx=20)
        self.button1 = customtkinter.CTkButton(self, text="TMT Set Up")
        self.button1.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.button2 = customtkinter.CTkButton(self, text="Map Tape Fields to Destination Fields")
        self.button2.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.button3 = customtkinter.CTkButton(self, text="Button 3")
        self.button3.grid(row=3, column=0, padx=10, pady=5, sticky="ew")


class MyFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # add widgets onto the frame, for example:
        self.label = customtkinter.CTkLabel(self)
        self.label.grid(row=0, column=0, padx=20)


AppWindow = customtkinter.CTk()
AppWindow.geometry("1300x700x5x5")
AppWindow.grid_rowconfigure(0, weight=1)  # configure grid system
AppWindow.grid_columnconfigure(1,weight = 1) 

AppWindow.navigation_frame  = NavigationFrame(master=AppWindow)
AppWindow.navigation_frame.grid(row=0, column=0, pady=20, sticky="nsew")

AppWindow.my_frame = MyFrame(master=AppWindow)
AppWindow.my_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        

AppWindow.mainloop()