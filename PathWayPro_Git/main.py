if __name__ == "__main__":

    import customtkinter
    import os
    import pandas as pd
    from openpyxl import load_workbook
    import pathway_objects 
    import time
    import shutil
    import datetime

    import os
    import pandas as pd

    new_path = pathway_objects.FileProcessor()
    working_dir = os.getcwd() + "\\"

    #Button Functions

    def transfer_to_excel():
        csv = working_dir + csv_entry.get()
        excel = working_dir + excel_entry.get()
        sheet = working_dir + sheet_entry.get()
        saved = new_path.add_to_excel(csv, excel, sheet)
        results_box.delete("1.0", "end")
        results_box.insert("end", text=f"{saved}")
        results_box.insert("end", text=" \n Confirming... ")
        new_path.reset_master_rep(csv)
        results_box.insert("end", text=f"\n Data Added to {excel} - {sheet}\n Transfer file reset confirmed \n ******You May Now Close the Window*******")



    def process_files():
        results_box.delete("1.0", "end")
        returned = new_path.process()
        new_path.move_and_delete_excel_files()
        action = [i for i in returned]
        results_box.insert("end", text=f"\n".join(action) + "\n****Continue with Merge Process Data****")

    def merge_files():
        returned = new_path.merge()
        action = [i for i in returned]  # Collect items in the action list
        results_box.delete("1.0", "end")
        results_box.insert("end", " ".join(action) + "\n added to transition file")  # Join the list with newlines and insert


    # Setting themes for GUI
    customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
    customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

    #Building the window
    root = customtkinter.CTk()
    root.title("When Paths Meet")
    root.geometry("600x500")


    #widgets for the window
    process_button = customtkinter.CTkButton(master=root, text="Prepare Unprocessed Files", command=process_files)
    process_button.place(relx=0.85, rely=0.05, anchor=customtkinter.CENTER)

    add_to_master_button = customtkinter.CTkButton(master=root, text="Merge Processed Data", command=merge_files)
    add_to_master_button.place(relx=0.85, rely=0.12, anchor=customtkinter.CENTER)

    title_label = customtkinter.CTkLabel(master=root, text="PathWayPro", font=customtkinter.CTkFont(family="cybernetic_font", size=50, weight="bold", slant="italic"), text_color="green", width=220)
    title_label.place(relx=0.5, rely= 0.05, anchor=customtkinter.NE)

    querry_button = customtkinter.CTkButton(master=root, text="Targeted Querry")
    querry_button.place(relx=0.27, rely=0.2, anchor=customtkinter.NE)

    csv_entry = customtkinter.CTkEntry(master=root, width=150, height=25, text_color="green", border_width=.5, border_color="green", corner_radius=8)
    csv_entry.insert(0, "master_rep.csv")
    csv_entry.place(relx=0.3, rely=0.3, anchor=customtkinter.NE)

    Excel_button = customtkinter.CTkButton(master=root, text="Add to Excel", command=transfer_to_excel)
    Excel_button.place(relx=0.85, rely=0.19, anchor=customtkinter.CENTER)

    excel_entry = customtkinter.CTkEntry(master=root, width=150, height=25, text_color="green", border_width=.5, border_color="green", corner_radius=8)
    excel_entry.insert(0, "DDS hits for 2024.xlsx")
    excel_entry.place(relx=0.57, rely=0.3, anchor=customtkinter.NE)

    sheet_entry = customtkinter.CTkEntry(master=root, width=150, height=25, text_color="green", border_width=.5, border_color="green", corner_radius=8)
    sheet_entry.insert(0, "Summer calibrations")
    sheet_entry.place(relx=0.84, rely=0.3, anchor=customtkinter.NE)

    results_box = customtkinter.CTkTextbox(master=root, width=550, height=290, font=customtkinter.CTkFont(family="cybernetic_font", size=18, weight="bold", slant="italic"), text_color="green", border_width=.8, border_color="dark green")
    results_box.place(relx=0.5, rely=0.66, anchor=customtkinter.CENTER)



    backup_confirm = new_path.backup_excel_file()
    results_box.insert("0.0", text=backup_confirm)
    new_path.reset_master_rep()

    root.mainloop()



