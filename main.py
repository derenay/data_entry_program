import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import *
from PIL import ImageTk, Image
import customtkinter as ct
from customtkinter import *
import sys
import os.path
import openpyxl

file_path = "deneme1.xlsx"
data_global = pd.read_excel(file_path)
Export_file_path = "exported.xlsx"
existing_df = pd.read_excel(file_path)


def selam():
    root = tk.Tk()
    root.title("Forest")
    root.option_add("*tearOff", False)
    root.geometry("1320x600")
    root.resizable(False, False)

    Envanter_No = tk.StringVar()

    IZM_no = tk.StringVar()
    Bilgisayar_No = tk.StringVar()
    entry_var = tk.StringVar()
    User_var = tk.StringVar()
    Text_var = tk.StringVar()

    check_frame_color = "#5DADE2"
    bg_color_root = "#F8F8F8"
    list_path = "list.xlsx"

    # Create lists for the Comboboxes

    Envanter_tipi_list = []
    Isletim_sistemi_list = []
    Marka_list = []
    Model_list = []
    CPU_list = []
    RAM_list = []
    Disk_list = []
    Disk_turu_list = []
    Group_id_list = []
    bit_check_list = ["evet", "hayır"]

    df = pd.read_excel(list_path)

    for index, row in df.iterrows():
        if pd.notna(row['Envanter_tipi']):
            Envanter_tipi_list.append(str(row['Envanter_tipi']))
        if pd.notna(row['Isletim_sistemi']):
            Isletim_sistemi_list.append(str(row['Isletim_sistemi']))
        if pd.notna(row['Marka']):
            Marka_list.append(str(row['Marka']))
        if pd.notna(row['Model']):
            Model_list.append(str(row['Model']))
        if pd.notna(row['CPU']):
            CPU_list.append(str(row['CPU']))
        if pd.notna(row['RAM']):
            RAM_list.append(str(row['RAM']))
        if pd.notna(row['Disk']):
            Disk_list.append(str(row['Disk']))
        if pd.notna(row['Disk_turu']):
            Disk_turu_list.append(str(row['Disk_turu']))
        if pd.notna(row['Group_id']):
            Group_id_list.append(str(row['Group_id']))

    image = ct.CTkImage(light_image=Image.open("lisi-aerospace-FTB.png"), size=(190, 106))

    panel = CTkLabel(master=root, corner_radius=8, bg_color=bg_color_root, height=150)
    panel.grid(row=0, column=0)

    photo_panel = CTkLabel(master=root, image=image, text="")
    photo_panel.place(x=1130, y=493)

    def submit():

        envanter_no = Envanter_No.get()
        bilgisayar_adi = Bilgisayar_No.get()
        izm_no = IZM_no.get()
        kullanıcı = User_var.get()
        envanter_tipi = Envanter_tipi.get()
        isletim_sistemi = Isletim_tur.get()
        sMarka = marka.get()
        model = Model.get()
        cpu = CPU.get()
        ram = RAM.get()
        disk = Disk.get()
        disk_turu = Disk_turu.get()
        group_id = Group_ID.get()
        text_var = Text_var.get()
        Bit_var = bit_check.get()

        # data dict
        data_dict = {'Envanter No': [envanter_no],
                     'Bağlandığı PC': [izm_no],
                     'Bilgisayar Adı': [bilgisayar_adi],
                     'Envanter Tipi': [envanter_tipi],
                     'İşletim Sistemi': [isletim_sistemi],
                     'Marka': [sMarka],
                     'Model': [model],
                     'CPU': [cpu],
                     'RAM': [ram],
                     'Disk': [disk],
                     'Disk Türü': [disk_turu],
                     'Kullanıcı': [kullanıcı],
                     'Lokasyon': [group_id],
                     'Not': [text_var],
                     '64x': [Bit_var]
                     }
        columns = ["Envanter No", "Bilgisayar Adı", "Bağlandığı PC", "Envanter Tipi", "İşletim Sistemi",
                   "Marka", "Model", "CPU", "RAM", "Disk", "Disk Türü", "Kullanıcı", "Lokasyon", "Not", "64x"]
        data = pd.DataFrame(data_dict, columns=columns)

        if envanter_no == "q":
            # Çıkış yapıldı
            root.destroy()
        elif envanter_no != "":
            # buradan kayıt yapılacak
            save_to_excel(data)
        else:
            # bu kısımda kayıt alınmaz
            my_data(existing_df)

        # Clear the form fields
        Envanter_No.set("")
        IZM_no.set("")
        Bilgisayar_No.set("")
        User_var.set("")
        Envanter_tipi.set("Envanter Tipi(nan)")
        Isletim_tur.set("İşletim Türü(nan)")
        marka.set("Marka(nan)")
        Model.set("Model(nan)")
        CPU.set("CPU(nan)")
        RAM.set("RAM(nan)")
        Disk.set("Disk(nan)")
        Disk_turu.set("Disk Türü(nan)")
        Group_ID.set("Lokasyon(nan)")
        bit_check.set("64x(nan)")
        Text_var.set("")

    def remove_from_excel():

        # df.shape[0] row count
        remove_no_var = entry_var.get()
        existing_df = pd.read_excel(file_path)

        existing_df = existing_df.loc[existing_df['Envanter No'] != remove_no_var]
        existing_df.to_excel(file_path, index=False)
        entry_var.set("")
        my_data(existing_df)

    def save_to_excel(data):

        try:
            existing_df = pd.read_excel(file_path)
            updated_df = pd.concat([existing_df, data], ignore_index=True)
            updated_df.to_excel(file_path, index=False)
            my_data(existing_df)
        except FileNotFoundError:
            data.to_excel(file_path, index=False)

    def search():

        tree.delete(*tree.get_children())

        existing_df = pd.read_excel(file_path)
        Search_var = entry_var.get()
        DfResults = pd.DataFrame(columns=existing_df.columns)
        Search_var = Search_var.replace("\\", "\\\\")
        connection_values = []

        for _, row in existing_df.iterrows():
            # checking all data to if there is a match
            search_result = row.astype(str).str.contains(Search_var, na=False, case=False)

            if search_result.any():
                connection_values.append(str(row.iloc[0]))
                # check if my searched value's row first item if its connected other rows
        DfResults = existing_df[
            existing_df.iloc[:, 2].isin(connection_values) | existing_df.iloc[:, 0].isin(connection_values)]
        # reduces duplicates
        DfResults = DfResults.drop_duplicates()
        shorted_DfResults = DfResults.sort_values(DfResults.columns[0])
        # parent child function
        for _, row in shorted_DfResults.iterrows():
            first_column_value = row.iloc[0]

            # parent
            parent_id = tree.insert("", "end", values=tuple(row))
            selected_rows = shorted_DfResults[(shorted_DfResults.iloc[:, 2] == first_column_value)]

            for _, row_second in selected_rows.iterrows():
                child_first_column_value = row_second.iloc[0]

                if first_column_value != child_first_column_value:
                    # child
                    tree.insert(parent_id, "end", values=row_second.tolist(), tags=('color2'))

        entry_var.set("")

    def my_data(existing_df):

        tree.delete(*tree.get_children())
        # pd.read_excel(file_path)
        # existing_df = pd.read_excel(file_path)

        tree.tag_configure('color1', background='white')
        tree.tag_configure('color2', background='purple')

        # short data
        # shorted_existing_df = existing_df.sort_values(existing_df.columns[0])

        for index, row in existing_df.iterrows():
            first_column_value = row.iloc[0]
            fourth_column = row.iloc[3]

            # view what you want in on fourth column
            if fourth_column == "Notebook" or \
                    fourth_column == "Desktop" \
                    :

                parent_id = tree.insert("", "end", values=row.tolist())
                selected_rows = existing_df[(existing_df.iloc[:, 2] == first_column_value)]

                for _, row_second in selected_rows.iterrows():

                    child_first_column_value = row_second.iloc[0]

                    if first_column_value != child_first_column_value:
                        tree.insert(parent_id, "end", values=row_second.tolist(), tags=('color2'))

    def update_record():
        # see update records on tree
        selected = tree.focus()
        selected_values = tree.item(
            selected,
            text="",
            values=(
                Envanter_No.get(),
                Bilgisayar_No.get(),
                IZM_no.get(),
                Envanter_tipi.get(),
                Isletim_tur.get(),
                marka.get(),
                Model.get(),
                CPU.get(),
                RAM.get(),
                Disk.get(),
                Disk_turu.get(),
                User_var.get(),
                Group_ID.get(),
                Text_var.get(),
                bit_check.get()
            ))

        selected_values = tree.item(selected, "values")

        first_old_value = selected_values[0]

        envanter_no = Envanter_No.get()

        existing_df = pd.read_excel(file_path)
        # location of ......
        existing_df = existing_df.loc[existing_df['Envanter No'] != first_old_value]
        existing_df.to_excel(file_path, index=False)

        update_df = pd.read_excel(file_path)

        # Note for myself I need to organizing here ###############################
        updated_envanter_no = Envanter_No.get()
        updated_bilgisayar_adi = Bilgisayar_No.get()
        updated_izm_no = IZM_no.get()
        updated_kullanıcı = User_var.get()
        updated_envanter_tipi = Envanter_tipi.get()
        updated_isletim_sistemi = Isletim_tur.get()
        updated_sMarka = marka.get()
        updated_model = Model.get()
        updated_cpu = CPU.get()
        updated_ram = RAM.get()
        updated_disk = Disk.get()
        updated_disk_turu = Disk_turu.get()
        updated_group_id = Group_ID.get()
        updated_text_var = Text_var.get()
        updated_bit_var = bit_check.get()

        updated_data_dict = {
            'Envanter No': [updated_envanter_no],
            'Bağlandığı PC': [updated_izm_no],
            'Bilgisayar Adı': [updated_bilgisayar_adi],
            'Envanter Tipi': [updated_envanter_tipi],
            'İşletim Sistemi': [updated_isletim_sistemi],
            'Marka': [updated_sMarka],
            'Model': [updated_model],
            'CPU': [updated_cpu],
            'RAM': [updated_ram],
            'Disk': [updated_disk],
            'Disk Türü': [updated_disk_turu],
            'Kullanıcı': [updated_kullanıcı],
            'Lokasyon': [updated_group_id],
            'Not': [updated_text_var],
            '64x': [updated_bit_var]
        }

        updated_columns = ["Envanter No", "Bilgisayar Adı", "Bağlandığı PC", "Envanter Tipi", "İşletim Sistemi",
                           "Marka", "Model", "CPU", "RAM", "Disk", "Disk Türü", "Kullanıcı", "Lokasyon", "Not", "64x"]

        updated_data = pd.DataFrame(updated_data_dict, columns=updated_columns)
        save_to_excel(updated_data)

        Envanter_entry.delete(0, END)
        Bilgisayar_entry.delete(0, END)
        IZM_entry.delete(0, END)
        user_entry.delete(0, END)
        Text_field.delete(0, END)
        Envanter_tipi.set("Envanter Tipi(nan)")
        Isletim_tur.set("İşletim Türü(nan)")
        marka.set("Marka(nan)")
        Model.set("Model(nan)")
        CPU.set("CPU(nan)")
        RAM.set("RAM(nan)")
        Disk.set("Disk(nan)")
        Disk_turu.set("Disk Türü(nan)")
        Group_ID.set("Lokasyon(nan)")
        bit_check.set("64x(nan)")

    def select_record():

        Envanter_entry.delete(0, END)
        Bilgisayar_entry.delete(0, END)
        IZM_entry.delete(0, END)
        user_entry.delete(0, END)
        Text_field.delete(0, END)

        # select when pressed
        selected = tree.focus()
        selected_values = tree.item(selected, "values")

        Envanter_entry.insert(0, selected_values[0])
        Bilgisayar_entry.insert(0, selected_values[1])
        IZM_entry.insert(0, selected_values[2])
        user_entry.insert(0, selected_values[11])
        Text_field.insert(0, selected_values[13])
        Envanter_tipi.set(selected_values[3])
        Isletim_tur.set(selected_values[4])
        marka.set(selected_values[5])
        Model.set(selected_values[6])
        CPU.set(selected_values[7])
        RAM.set(selected_values[8])
        Disk.set(selected_values[9])
        Disk_turu.set(selected_values[10])
        Group_ID.set(selected_values[12])
        bit_check.set(selected_values[14])

    def sort_by_columns(col, order=False):
        global data_global

        data_global[col] = data_global[col].astype(str)
        data_global = data_global.sort_values(by=[col], ascending=not order)

        for i in tree.get_children():
            tree.delete(i)
        my_data(data_global)
        tree.heading(col, command=lambda: sort_by_columns(col, not order))

    def Export():
        file_exists = os.path.exists("exported.xlsx")

        if file_exists:
            data_global.to_excel(Export_file_path, index=False)
        else:
            exported = openpyxl.Workbook()
            exported.save(Export_file_path)
            data_global.to_excel(Export_file_path, index=False)

    # Create a style
    style = ttk.Style(root)

    # Make the app responsive
    root.columnconfigure(index=0, weight=1)
    root.columnconfigure(index=1, weight=1)
    root.columnconfigure(index=2, weight=1)
    root.rowconfigure(index=0, weight=1)
    root.rowconfigure(index=1, weight=1)
    root.rowconfigure(index=2, weight=1)
    root.configure(bg=bg_color_root)

    # Create control variables
    a = tk.BooleanVar()
    b = tk.BooleanVar(value=True)
    c = tk.BooleanVar()
    d = tk.IntVar(value=2)
    f = tk.BooleanVar()
    g = tk.DoubleVar(value=75.0)
    h = tk.BooleanVar()

    # Create a Frame for the text box
    check_frame = CTkFrame(master=panel, fg_color=check_frame_color)
    check_frame.grid(row=0, column=0, padx=(20, 10), pady=(20, 10), sticky="nsew")

    # Text field for not
    Text_field = CTkEntry(master=check_frame, textvariable=Text_var, font=('calibre', 10, 'normal'),
                          border_width=0, corner_radius=10)
    Text_field.grid(row=2, column=6)

    #######Inventory#############
    Envanter_label = CTkLabel(master=check_frame, text='Envanter No:', font=('calibre', 10, 'bold'),
                              fg_color=check_frame_color)
    Envanter_label.grid(row=2, column=4, padx=5, pady=10, sticky="ew")
    Envanter_entry = CTkEntry(master=check_frame, textvariable=Envanter_No, font=('calibre', 10, 'normal'),
                              border_width=0, corner_radius=10)
    Envanter_entry.grid(row=2, column=5, padx=5, pady=10, sticky="ew")

    #######Computer#############
    Bilgisayar_label = CTkLabel(master=check_frame, text='Bilgisayar No:', font=('calibre', 10, 'bold'),
                                fg_color=check_frame_color)
    Bilgisayar_label.grid(row=1, column=2, padx=5, pady=10, sticky="ew")
    Bilgisayar_entry = CTkEntry(master=check_frame, textvariable=Bilgisayar_No, font=('calibre', 10, 'normal')
                                , border_width=0, corner_radius=10)
    Bilgisayar_entry.grid(row=1, column=3, padx=5, pady=10, sticky="ew")

    #######IZM#############
    IZM_label = CTkLabel(master=check_frame, text='Bağlandığı PC:', font=('calibre', 10, 'bold'),
                         fg_color=check_frame_color)
    IZM_label.grid(row=1, column=4, padx=5, pady=10, sticky="ew")
    IZM_entry = CTkEntry(master=check_frame, textvariable=IZM_no, font=('calibre', 10, 'normal'),
                         border_width=0, corner_radius=10)
    IZM_entry.grid(row=1, column=5, padx=5, pady=10, sticky="ew")

    #######Entry#############
    entrybox_label = CTkLabel(master=check_frame, text='Search/Remove:', font=('calibre', 10, 'bold'),
                              fg_color=check_frame_color)
    entrybox_label.grid(row=2, column=2, sticky="ew")
    entrybox_label = CTkEntry(master=check_frame, textvariable=entry_var, font=('calibre', 10, 'normal'),
                              border_width=0, corner_radius=10)
    entrybox_label.grid(row=2, column=3, sticky="ew")

    #######User#############
    user_label = CTkLabel(master=check_frame, text='Kullanıcı:', font=('calibre', 10, 'bold'),
                          fg_color=check_frame_color)
    user_label.grid(row=1, column=6, sticky="ew")
    user_entry = CTkEntry(master=check_frame, textvariable=User_var, font=('calibre', 10, 'normal'),
                          border_width=0, corner_radius=10)
    user_entry.grid(row=1, column=7, sticky="ew")

    ###################

    # Checkbuttons
    # Envanter_tipi
    Envanter_tipi = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=Envanter_tipi_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    Envanter_tipi.set("Envanter Tipi(nan)")
    Envanter_tipi.grid(row=0,
                       column=0,
                       padx=5,
                       pady=10,
                       sticky="ew"
                       )
    # İşletim sistemi
    Isletim_tur = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=Isletim_sistemi_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white"
    )
    Isletim_tur.set("İşletim Türü(nan)")
    Isletim_tur.grid(row=0,
                     column=1,
                     padx=5,
                     pady=10,
                     sticky="ew"
                     )
    # Marka
    marka = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=Marka_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    marka.set("Marka(nan)")
    marka.grid(row=0,
               column=2,
               padx=5,
               pady=10,
               sticky="ew"
               )
    # Model
    Model = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=Model_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    Model.set("Model(nan)")

    Model.grid(
        row=0,
        column=3,
        padx=5,
        pady=10,
        sticky="ew"
    )
    # CPU
    CPU = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=CPU_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    CPU.set("CPU(nan)")
    CPU.grid(
        row=0,
        column=4,
        padx=5,
        pady=10,
        sticky="ew"
    )
    # RAM
    RAM = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=RAM_list,
        border_width=0,
        button_color="#787878",
        corner_radius=10
    )
    RAM.set("RAM(nan)")
    RAM.grid(row=0,
             column=5,
             padx=5,
             pady=10,
             sticky="ew"
             )
    # Disk
    Disk = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=Disk_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    Disk.set("Disk(nan)")
    Disk.grid(
        row=0,
        column=6,
        sticky="ew"
    )
    # Disk_türü
    Disk_turu = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=Disk_turu_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    Disk_turu.set("Disk Türü(nan)")
    Disk_turu.grid(
        row=0,
        column=7,
        padx=5,
        pady=10,
        sticky="ew"
    )

    # Group_ID
    Group_ID = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=Group_id_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    Group_ID.set("Lokasyon(nan)")
    Group_ID.grid(
        row=1,
        column=0,
        padx=5,
        pady=10,
        sticky="ew"
    )
    # bit_check
    bit_check = CTkComboBox(
        master=check_frame,
        state="readonly",
        values=bit_check_list,
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    bit_check.set("64x(nan)")
    bit_check.grid(
        row=1,
        column=1,
        padx=5,
        pady=10,
        sticky="ew"
    )
    # empty_3
    empty_3 = CTkComboBox(
        master=check_frame,
        state="readonly",
        values="",
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    empty_3.set("empty_3")
    empty_3.grid(
        row=2,
        column=0,
        padx=5,
        pady=10,
        sticky="ew"
    )
    # empty_4
    empty_4 = CTkComboBox(
        master=check_frame,
        state="readonly",
        values="",
        border_width=0,
        button_color="#787878",
        button_hover_color="white",
        corner_radius=10
    )
    empty_4.set("empty_4")
    empty_4.grid(
        row=2,
        column=1,
        padx=5,
        pady=10,
        sticky="ew"
    )

    style.configure("CTkComboBox", background="white")

    # Create a Frame for input widgets
    widgets_frame = CTkFrame(master=panel, fg_color="#EEEEAD")
    widgets_frame.grid(row=1, column=0, padx=10, pady=(30, 10), sticky="nsew")
    widgets_frame.columnconfigure(index=0, weight=1)

    # Tree

    columns_tree = ["Envanter No", "Bilgisayar Adı", "Bağlandığı PC", "Envanter Tipi", "İşletim Sistemi",
                    "Marka", "Model", "CPU", "RAM", "Disk", "Disk Türü", "Kullanıcı", "Lokasyon", "Not", "64x"]
    tree = ttk.Treeview(widgets_frame, columns=columns_tree, show='headings', selectmode="extended")
    tree['height'] = 12
    y_treeScroll = Scrollbar(widgets_frame, orient="vertical", command=tree.yview, highlightcolor="white",
                             elementborderwidth=0, highlightbackground="black", jump=0)
    y_treeScroll.config(command=tree.yview)

    ttk.Style().configure("Treeview", background="#8B8B65", foreground="white",
                          yscrollcommand=y_treeScroll.set)
    for column in columns_tree:
        tree.column(column, anchor=CENTER, stretch=NO, width=84)
        tree.heading(column, text=column, command=lambda column=column: sort_by_columns(column))

    tree.pack(side='left')
    y_treeScroll.pack(side='left', fill='y')

    # Add Button
    button = CTkButton(
        master=check_frame,
        text="Add/Relist",
        command=submit,
        fg_color=("#787878"),
        hover_color="white",
        corner_radius=10
    )
    button.grid(row=2, column=7, padx=5, pady=10, sticky="nsew")

    # remove Button
    remove = CTkButton(
        master=root,
        text="Remove ",
        command=remove_from_excel,
        fg_color=("#787878"),
        hover_color="white",
        corner_radius=10
    )
    remove.place(x=919, y=191)

    # Search Button
    Search_button = CTkButton(
        master=root,
        text="Search",
        command=search,
        fg_color=("#787878"),
        hover_color="white",
        corner_radius=10
    )
    Search_button.place(x=1069, y=191)

    # Update Button
    Update_button = CTkButton(
        master=root,
        text="Update",
        fg_color=("#787878"),
        hover_color="white",
        corner_radius=10,
        command=update_record
    )
    Update_button.place(x=769, y=191)

    # Select Button
    Select_button = CTkButton(
        master=root,
        text="Select",
        fg_color=("#787878"),
        hover_color="white",
        corner_radius=10,
        command=select_record
    )
    Select_button.place(x=619, y=191)

    # export Button
    export_button = CTkButton(
        master=root,
        text="Export",
        fg_color=("#787878"),
        hover_color="white",
        corner_radius=10,
        command=Export
    )
    export_button.place(
        x=29,
        y=191
    )

    my_data(data_global)
    # Start the main loop
    root.mainloop()


if __name__ == "__main__":
    if "--logged_in" in sys.argv:
        selam()

    else:
        print("Please log in")
        selam()

