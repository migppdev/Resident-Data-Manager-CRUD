import tkinter as tk
from database import add_resident_db, get_residents, search_resident_db, delete_resident, get_data, update_data_db, close_db, import_excel_db, export_excel_db, delete_all_db
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
import locale

class Resident:
    def __init__(self, name, age, registration_date):
        self.name = name
        self.age = age
        self.registration_date = registration_date

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # Configure window 
        self.geometry("600x400")
        self.title("Resident Manager")
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=5)
        self.rowconfigure(2, weight=1)
        self.columnconfigure(0, weight=1)

        #* Menu
        self.menu_bar = tk.Menu(self)

        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="Import from Excel file", command=self.import_excel)
        self.file_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        self.file_menu.add_command(label="Exit", command=self.close_window)        
        self.edit_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Edit", menu=self.edit_menu)
        self.edit_menu.add_command(label="Delete all", command=self.delete_all_residents)

        self.config(menu=self.menu_bar)

        # Search Frame
        self.search_frame = tk.Frame(self)
        self.search_frame.rowconfigure(0, weight=1)
        self.search_frame.columnconfigure(0, weight=2)
        self.search_frame.columnconfigure(1, weight=1)
        self.search_frame.grid(row=0, column=0, sticky="nsew")

        self.search_entry = tk.Entry(self.search_frame, font=("Arial Black", 10), bg="#1e90b0")
        self.search_entry.grid(row=0, column=0, sticky="nsew")
        self.search_button = tk.Button(self.search_frame,font=("Arial", 14), text="🔎", bg="#1e81b0", command=lambda: self.search_resident(self.search_entry.get()))
        self.search_button.grid(row=0, column=1, sticky="nsew")

        # Buttons Frame
        self.buttons_frame = tk.Frame(self)
        self.buttons_frame.rowconfigure(0, weight=1)
        self.buttons_frame.columnconfigure(0, weight=1)
        self.buttons_frame.columnconfigure(1, weight=1)
        self.buttons_frame.columnconfigure(2, weight=1) 
        self.buttons_frame.columnconfigure(3, weight=1) 
        self.buttons_frame.grid(row=2, column=0, sticky="nsew")
        
        # List Creation
        self.residents_list = tk.Listbox(font=("Arial", 15), bg="#7fbfd0", selectbackground="#3d7979", highlightthickness=0, selectmode="single", activestyle="none", selectforeground="black")
        self.residents_list.grid(row=1, column=0, sticky="nsew")

        # Add buttons for the list
        self.view_resident_button = tk.Button(self.buttons_frame, bg="#1e81b0", text="View information", command=self.view_resident_information)
        self.view_resident_button.grid(row=0, column=0, sticky="nsew")
        
        self.new_resident_button = tk.Button(self.buttons_frame, bg="#1e81b0", text="Add resident", command=self.new_resident)
        self.new_resident_button.grid(row=0, column=1, sticky="nsew")

        self.delete_resident_button = tk.Button(self.buttons_frame, bg="#1e81b0", text="Delete resident", command=self.delete_resident)
        self.delete_resident_button.grid(row=0, column=2, sticky="nsew")

        self.edit_resident_button = tk.Button(self.buttons_frame, bg="#1e81b0", text="Edit resident", command=self.edit_resident)
        self.edit_resident_button.grid(row=0, column=3, sticky="nsew")

    def view_resident_information(self):
        try:
            selected_resident = self.residents_list.get(self.residents_list.curselection())
            self.view_info_window = tk.Toplevel()
            self.view_info_window.geometry("400x300")
            self.view_info_window.title(f"Information of {selected_resident}")
            self.view_info_window.rowconfigure(0, weight=1)
            self.view_info_window.rowconfigure(1, weight=1)
            self.view_info_window.rowconfigure(2, weight=1)
            self.view_info_window.columnconfigure(0, weight=1)
            self.view_info_window.columnconfigure(1, weight=1)

            tk.Label(self.view_info_window, text="Full Name").grid(row=0, column=0)
            self.name_label = tk.Label(self.view_info_window, text=selected_resident, highlightthickness=1, highlightbackground="black")
            self.name_label.grid(row=0, column=1)

            tk.Label(self.view_info_window, text="Age").grid(row=1, column=0)
            self.age_label = tk.Label(self.view_info_window, text=get_data(selected_resident)[2], highlightthickness=1, highlightbackground="black")
            self.age_label.grid(row=1, column=1)

            tk.Label(self.view_info_window, text="Registration Date").grid(row=2, column=0)
            self.date_label = tk.Label(self.view_info_window, text=get_data(selected_resident)[3], highlightthickness=1, highlightbackground="black")
            self.date_label.grid(row=2, column=1)

        except tk.TclError:
            messagebox.showerror(title="Error", message="You haven't selected any resident")

    def new_resident(self):

        #* Configure window
        self.new_resident_window = tk.Toplevel()
        self.new_resident_window.title("Add resident")
        self.new_resident_window.rowconfigure(0, weight=1)
        self.new_resident_window.rowconfigure(1, weight=1)
        self.new_resident_window.columnconfigure(0, weight=1)
        
        #* New resident data Frame
        self.data_frame = tk.Frame(self.new_resident_window)
        self.data_frame.grid(row=0, column=0)
        
        # Configure frame
        self.data_frame.rowconfigure(0, weight=1)
        self.data_frame.rowconfigure(1, weight=1)
        self.data_frame.rowconfigure(2, weight=1)
        self.data_frame.columnconfigure(0, weight=1)
        self.data_frame.columnconfigure(1, weight=1)
        self.data_frame.grid(row=0, column=0, sticky="nsew")
        
        # Fields for the data frame
        self.full_name_label = tk.Label(self.data_frame, text="Full Name")
        self.full_name_label.grid(row=0, column=0)
        self.full_name_entry = tk.Entry(self.data_frame)      # Name
        self.full_name_entry.grid(row=0, column=1, sticky="nsew")

        self.age_label = tk.Label(self.data_frame, text="Age")
        self.age_label.grid(row=1, column=0)
        self.age_entry = tk.Spinbox(self.data_frame, from_=18, to=150)        # Age
        self.age_entry.grid(row=1, column=1)
        
        # Calendar 
        locale.setlocale(locale.LC_TIME, '') 
        self.calendar = DateEntry(self.data_frame, date_pattern="dd/MM/yyyy")
        self.date_label = tk.Label(self.data_frame, text="Registration Date")
        self.date_label.grid(row=2, column=0)
        self.calendar.grid(row=2, column=1)

        # Button Frame
        self.add_button_frame = tk.Frame(self.new_resident_window)
        self.add_button_frame.rowconfigure(0, weight=1)
        self.add_button_frame.columnconfigure(0, weight=1)
        self.add_button_frame.grid(row=1, column=0, sticky="nsew")

        # Function to collect data from entries and add it to the database
        def confirm_resident():
            global resident
            resident = Resident(name=self.full_name_entry.get(), age=self.age_entry.get(), registration_date=self.calendar.get())      
            add_resident_db(resident.name, resident.age, resident.registration_date)  
            self.new_resident_window.destroy()
            self.update_list()
            
        self.add_button = tk.Button(self.add_button_frame, text="Confirm", command=confirm_resident)
        self.add_button.grid(row=0, column=0, sticky="nsew") 
        self.new_resident_window.mainloop()

    def edit_resident(self):
        try:
            selected_resident = self.residents_list.get(self.residents_list.curselection())
            self.edit_window = tk.Toplevel()
            self.edit_window.geometry("400x300")
            self.edit_window.title(f"Editing {selected_resident}")
            self.edit_window.rowconfigure(0, weight=4)
            self.edit_window.rowconfigure(1, weight=1)
            self.edit_window.columnconfigure(0, weight=1)

            self.info_frame = tk.Frame(self.edit_window)
            self.info_frame.grid(row=0, column=0, sticky="nsew")

            self.info_frame.rowconfigure(0, weight=1)
            self.info_frame.rowconfigure(1, weight=1)
            self.info_frame.rowconfigure(2, weight=1)
            self.info_frame.columnconfigure(0, weight=1)
            self.info_frame.columnconfigure(1, weight=1)

            tk.Label(self.info_frame, text="Name").grid(row=0, column=0)
            self.new_name = tk.Entry(self.info_frame)
            self.new_name.grid(row=0, column=1)

            tk.Label(self.info_frame, text="Age").grid(row=1, column=0)
            self.new_age = tk.Entry(self.info_frame)
            self.new_age.grid(row=1, column=1)

            tk.Label(self.info_frame, text="Registration Date").grid(row=2, column=0)
            # Keeping the date pattern for consistency, setting locale to neutral
            self.new_date = DateEntry(self.info_frame, date_pattern="dd/MM/yyyy") 
            self.new_date.grid(row=2, column=1)

            self.button_frame = tk.Frame(self.edit_window)
            self.button_frame.grid(row=1, column=0, sticky="nsew")
            self.button_frame.rowconfigure(0, weight=1)
            self.button_frame.columnconfigure(0, weight=1)

            # Add data to entries
            self.new_name.insert(0, get_data(selected_resident)[1])
            self.new_age.insert(0, get_data(selected_resident)[2])
            self.new_date.delete(0, tk.END)
            self.new_date.insert(0, get_data(selected_resident)[3])

            self.save_changes_button = tk.Button(self.button_frame,
                                                 text="Save changes",
                                                 command=lambda: self.update_data(self.new_name.get(), self.new_age.get(), self.new_date.get(), selected_resident))
            self.save_changes_button.grid(row=0, column=0, sticky="nsew")

        except tk.TclError:
            messagebox.showerror(title="Error", message="Select a resident")
            return 0

    def delete_resident(self):
        try:
            selection = self.residents_list.curselection()
            selected_resident = self.residents_list.get(selection)

            if messagebox.askyesno(title="Delete resident", message=f"Do you want to delete resident {selected_resident} from the database?"):
                delete_resident(full_name=selected_resident)
            else:
                pass

        except tk.TclError:
            messagebox.showerror(title="Error", message="Select a resident")

        finally:
            self.update_list()

    def update_list(self):
        self.residents_list.delete(0, tk.END)
        for resident in get_residents():
            self.residents_list.insert(tk.END, ' '.join(resident))

    def search_resident(self, query):
        self.residents_list.delete(0, tk.END)
        for resident in search_resident_db(query):
            self.residents_list.insert(tk.END, ' '.join(resident))

    def update_data(self, new_full_name, new_age, new_date, old_full_name):
        if messagebox.askyesno(title="Confirmation", message="Do you want to change the data?"):
            update_data_db(new_full_name, new_age, new_date, old_full_name)
            self.edit_window.destroy()
            self.update_list()
        else:
            pass

    def import_excel(self):
        excel_file_path = filedialog.askopenfilename()
        import_excel_db(file_path=excel_file_path)
        self.update_list()

    def export_to_excel(self):
        excel_file_path = filedialog.askopenfilename()
        export_excel_db(file_path=excel_file_path)
    
    def close_window(self):
        self.destroy()
        close_db()

    def delete_all_residents(self):
        if messagebox.askokcancel(title="Delete all data", message="You are about to delete all data from the database"):
            delete_all_db()
        else:
            pass


if __name__ == '__main__':
    app = App()
    app.update_list()
    app.protocol("WM_DELETE_WINDOW", app.close_window)

    app.mainloop()
